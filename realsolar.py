import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference
import calendar
import urllib.parse
import requests
import urllib3
from datetime import datetime

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Variables globales
resultats = []
lat, lon = 48.874, 2.295  #arc de triomphe
startyear, endyear = 2020, 2023

prix_kw = 0.2016  # €/kWh
date_departcumul = pd.Timestamp("2020-01-01")
conso_chauffeau_journalière = 2.2  # kWh



def build_pvgis_url(lat, lon, angle, aspect,libelle, startyear, endyear, MaxPower=1.00, tech='crystSi', db='PVGIS-SARAH3'):
    base_url = "https://re.jrc.ec.europa.eu/api/v5_3/seriescalc"
    params = {
        "lat": lat,
        "lon": lon,
        "raddatabase": db,
        "outputformat": "csv",
        "usehorizon": 1,
        "angle": angle,
        "aspect": aspect,
        "startyear": startyear,
        "endyear": endyear,
        "mountingplace": "free",
        "trackingtype": 0,              #Tracker 2 axe
        "pvcalculation": 1,
        "pvtechchoice": tech,
        "peakpower": MaxPower,
        "loss": 14
    }

    url = base_url + "?" + urllib.parse.urlencode(params)
    filename = f"data_{angle}deg_{aspect}deg_{startyear}_{endyear}_PMax{MaxPower:.2f}.csv"
    return url, filename


def telecharger_csv(url, filename):
    if not os.path.exists(filename):
        print(f"Téléchargement depuis : {url}")
        response = requests.get(url, verify=False)
        if response.ok:
            with open(filename, 'wb') as f:
                f.write(response.content)
            print(f"✅ Fichier enregistré sous : {filename}")
        else:
            print(f"❌ Erreur {response.status_code}")
    else:
        print(f"✔️ Fichier déjà existant : {filename}")


def traitement_csv(filename, maxpower):
    with open(filename, 'r', encoding='utf-8') as f:
        lignes = f.readlines()
    for i, ligne in enumerate(lignes):
        if ligne.strip().startswith("time,P"):
            header_idx = i
            break

    df = pd.read_csv(filename, skiprows=header_idx)
    df.columns = df.columns.str.strip().str.lower()
    df[['date', 'time']] = df['time'].str.split(":", expand=True)
    df = df[df['date'].astype(str).str.match(r'^\d{8}$')]
    df['date'] = pd.to_datetime(df['date'], format='%Y%m%d')
    df['time'] = df['time'].astype(str).str.zfill(4)
    df['hour'] = df['time'].str[:-2].astype(int)
    df['minute'] = df['time'].str[-2:].astype(int)
    df['datetime'] = df['date'] + pd.to_timedelta(df['hour'], unit='h') + pd.to_timedelta(df['minute'], unit='m')
    df['p'] = pd.to_numeric(df['p'], errors='coerce')

    return df






def traitement_cumul(filename,conso_maison_W):
    df = pd.read_excel(filename)
    conso_maison_kW = conso_maison_W / 1000

    df['date'] = pd.to_datetime(df['datetime']).dt.date
   
    df['p_dispo'] = 0.0
    df['p_maison'] = 0.0
    df['p_chauffeau'] = 0.0
    df['p_perdue'] = 0.0
  
    for date_jour, df_jour in df.groupby(df['datetime'].dt.date):
        energie_chauffeau_cumul = 0.0
        for i, row in df_jour.iterrows():
            p_dispo = row['p']
            df.at[i, 'p_dispo']=p_dispo
            if p_dispo >= conso_maison_W:
                df.at[i, 'p_maison'] = conso_maison_W
                p_rest = p_dispo - conso_maison_W
            else:
                df.at[i, 'p_maison'] = p_dispo
                p_rest = 0

            if energie_chauffeau_cumul < conso_chauffeau_journalière:
                reste = (conso_chauffeau_journalière - energie_chauffeau_cumul) * 1000
                chauffe_eau = min(p_rest, reste)
                df.at[i, 'p_chauffeau'] = chauffe_eau
                energie_chauffeau_cumul += chauffe_eau / 1000
                p_rest -= chauffe_eau

            df.at[i, 'p_perdue'] = p_rest

    df['date'] = pd.to_datetime(df['datetime'].dt.date)

    # Étape 1 : Identifier dynamiquement les colonnes 'scenario_*'
    scenario_cols = [col for col in df.columns if col.startswith('scenario_')]

    # Étape 2 : Ajouter les autres colonnes fixes
    other_cols = ['p_dispo', 'p_maison', 'p_chauffeau', 'p_perdue']

    # Construction du dictionnaire d'aggregation
    agg_dict = {col: 'sum' for col in scenario_cols + other_cols}

    # Étape 3 : Groupby + agg
    daily = df.groupby('date').agg(agg_dict) / 1000

    daily['manque'] = (
            (conso_chauffeau_journalière + conso_maison_kW * df.groupby('date').size() - 1.5  ) -         # on envele 2 heure
            (daily['p_chauffeau'] + daily['p_maison'])
        ).clip(lower=0)



    # Ajustement: si 'energie_perdue' > 0, alors 'manque' = 0
    daily.loc[daily['p_perdue'] > 0, 'manque'] = 0

    daily = daily.rename(columns={
        'p_maison': 'energie_maison',
        'p_chauffeau': 'energie_chauffeau',
        'p_perdue': 'energie_perdue'
    })
    daily['energie_recuperee'] = daily['energie_maison'] + daily['energie_chauffeau']

    daily['gain (€)'] = daily['energie_recuperee'] * prix_kw
    daily['perte (€)'] = daily['energie_perdue'] * prix_kw
    daily['manque (€)'] = daily['manque'] * prix_kw

    daily['cumul_gain (€)'] = daily[daily.index >= date_departcumul]['gain (€)'].cumsum()
    daily['cumul_perte (€)'] = daily[daily.index >= date_departcumul]['perte (€)'].cumsum()

    daily['cumul_manque (€)'] = daily[daily.index >= date_departcumul]['manque (€)'].cumsum()


    daily.loc[daily.index < date_departcumul, ['cumul_gain (€)', 'cumul_perte (€)']] = float('nan')

    
    return daily


from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference
import pandas as pd
import calendar

def ajouter_graphique_excel(fichier_excel, df_resultat):
    wb = load_workbook(fichier_excel)
    ws = wb.active

    # Recherche de l’index (colonne) de 'energie_maison'
    header = [cell.value for cell in ws[1]]
    try:
        col_energie_maison = header.index("energie_maison") + 1  # Excel est 1-based
    except ValueError:
        print("⚠️ Colonne 'energie_maison' non trouvée dans le fichier Excel.")
        return

    # Conversion des dates pour grouper par mois
    df_resultat["date"] = pd.to_datetime(df_resultat.index)
    df_resultat["mois"] = df_resultat["date"].dt.to_period("M")

    start_row = 2
    row_offset = 10
    chart_height = 5
    chart_width = 20

    for i, mois in enumerate(df_resultat["mois"].unique()):
        df_mois = df_resultat[df_resultat["mois"] == mois]
        excel_start = df_resultat.index.get_loc(df_mois.index[0]) + 2
        excel_end = excel_start + len(df_mois) - 1

        chart = BarChart()
        chart.type = "col"
        chart.grouping = "stacked"
        chart.overlap = 100
        chart.title = f"{calendar.month_name[mois.month]} {mois.year}"
        chart.y_axis.title = "Énergie (kWh)"
        chart.x_axis.title = "Jour"

        # Utilise la colonne 'energie_maison' et celle juste après (si tu veux plusieurs)
        data = Reference(ws,
                         min_col=col_energie_maison,
                         max_col=col_energie_maison + 3,  # Ajuste selon besoin
                         min_row=excel_start - 1,
                         max_row=excel_end)

        # Catégories (date/jour)
        cats = Reference(ws, min_col=1, max_col=1, min_row=excel_start, max_row=excel_end)

        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        chart.height = chart_height
        chart.width = chart_width

        ws.add_chart(chart, f"Q{start_row + i * row_offset}")

    wb.save(fichier_excel)




# === PARAMÈTRES MULTIPLES ===
import pandas as pd

# Définition des scénarios    
scenarios = [
    [         
        {"libelleScenario": "UnPanneau_optimal", "investissement": 120+ 80+126 ,"conso_maison_W":100},
        {"angle": 37, "aspect": -2, "MaxPower": 0.5, "libelle": "optimal"},
    ],

    [  
        {"libelleScenario": "1xBalcon_72", "investissement": 120+80+126 ,"conso_maison_W":100 },
        {"angle": 72, "aspect": -13.2, "MaxPower": 0.5, "libelle": "balcon 72"},
    ],
        [     
        {"libelleScenario": "1xBalcon_72_eauseul", "investissement": 120+80+45 ,"conso_maison_W":0 },
        {"angle": 72, "aspect": -13.2, "MaxPower": 0.5, "libelle": "balcon 72"},
    ],

    [  
        {"libelleScenario": "2xToiture_estouest", "investissement": 120+80+80+190 ,"conso_maison_W":100},
        {"angle": 23, "aspect": -103.2, "MaxPower": 0.5, "libelle": "toitest"},
        {"angle": 23, "aspect": 76.8, "MaxPower": 0.5, "libelle": "toitwest"},
    ],
    [  
        {"libelleScenario": "2xToiture_ouest", "investissement": 120+80+80+190,"conso_maison_W":100},
        {"angle": 23, "aspect": 76.8, "MaxPower": 0.5, "libelle": "ouest"},
        {"angle": 23, "aspect": 76.8, "MaxPower": 0.5, "libelle": "ouest"},
    ],

    [  
        {"libelleScenario": "2xBalcon_72", "investissement": 120+80+80+190,"conso_maison_W":100},
        {"angle": 72, "aspect": -13.2, "MaxPower": 0.5, "libelle": "balcon 72"},
        {"angle": 72, "aspect": -13.2, "MaxPower": 0.5, "libelle": "balcon 72"},
    ],
    [  
        {"libelleScenario": "2xGarage_Ouest", "investissement": 120+80+80+190},
        {"angle": 23, "aspect": 47, "MaxPower": 0.5, "libelle": "garage ouest"},
        {"angle": 23, "aspect": 47, "MaxPower": 0.5, "libelle": "garage ouest"},
    ],
    [  # mix optimal + balcon 
        {"libelleScenario": "2xMix_optimal_balcon72", "investissement": 120+80+80+126+126,"conso_maison_W":100},
        {"angle": 37, "aspect": -2, "MaxPower": 0.5, "libelle": "optimal"},
        {"angle": 72, "aspect": -13.2, "MaxPower": 0.5, "libelle": "balcon 72"},
    ],
    [  # optimal x2  
        {"libelleScenario": "2xOptimal", "investissement": 120+80+80+190,"conso_maison_W":100},
        {"angle": 37, "aspect": -2, "MaxPower": 0.5, "libelle": "optimal"},
        {"angle": 37, "aspect": -2, "MaxPower": 0.5, "libelle": "optimal"},
    ],

        [  # optimal x3
        {"libelleScenario": "3xoptimal", "investissement": 120+80+80+80+190+125,"conso_maison_W":100},
        {"angle": 37, "aspect": -2, "MaxPower": 0.5, "libelle": "optimal"},
        {"angle": 37, "aspect": -2, "MaxPower": 0.5, "libelle": "optimal"},
        {"angle": 37, "aspect": -2, "MaxPower": 0.5, "libelle": "optimal"},
    ],

        [  #  mix 
        {"libelleScenario": "1xOptimal_x2balcon72", "investissement": 120+80+80+190 +80+125,"conso_maison_W":100},
        {"angle": 72, "aspect": -13.5, "MaxPower": 0.5, "libelle": "balcon 72"},
        {"angle": 72, "aspect": -13.5, "MaxPower": 0.5, "libelle": "balcon 72"},
        {"angle": 37, "aspect": -2, "MaxPower": 0.5, "libelle": "optimal"},
 
    ],
    
         [  # optimal 
        {"libelleScenario": "4xOptimal", "investissement": 120+80+80+80+80+ 280,"conso_maison_W":100},
        {"angle": 37, "aspect": -2, "MaxPower": 0.52, "libelle": "optimal"},
        {"angle": 37, "aspect": -2, "MaxPower": 0.52, "libelle": "optimal"},
        {"angle": 37, "aspect": -2, "MaxPower": 0.52, "libelle": "optimal"},
         {"angle": 37, "aspect": -2, "MaxPower": 0.52, "libelle": "optimal"},
       
    ],   
]









# Traitement de chaque scénario
for scenario in scenarios:
    meta = scenario[0]
    configs = scenario[1:]

    print(f"\n▶️ Scénario: {meta['libelleScenario']}")
    for config in configs:
        print(f"  - Panneau: angle={config['angle']}°, aspect={config['aspect']}°, power={config['MaxPower']} kW")

    df_cumul = pd.DataFrame()
    dfs_scenarios = {}

    for i, params in enumerate(configs, 1):
        url, filename = build_pvgis_url(lat, lon, **params, startyear=startyear, endyear=endyear)
        telecharger_csv(url, filename)

        

        brut = traitement_csv(filename, params["MaxPower"])
        brut["datetime"] = pd.to_datetime(brut["datetime"])
        brut = brut[["datetime", "p"]]

        nom_finalbrut = f"energie_brut_{params['libelle']}_{i}_{params['angle']}_{params['aspect']}_{params['MaxPower']}.xlsx"
        brut.to_excel(nom_finalbrut, index=False)

        brut.set_index('datetime', inplace=True)
        dfs_scenarios[f"scenario_{i}"] = brut.copy()

        df_cumul = brut if df_cumul.empty else df_cumul.add(brut, fill_value=0)

    for name, df in dfs_scenarios.items():
        df_cumul[f"{name}"] = df["p"]

    nom_final = f"energie_cumulee_scenarios_{meta['libelleScenario']}.xlsx"
    df_cumul.to_excel(nom_final)

    df_cumulfin = traitement_cumul(nom_final ,  meta['conso_maison_W'] )
    df_cumulfin.to_excel(nom_final)
    ajouter_graphique_excel(nom_final, df_cumulfin)

    last_row = df_cumulfin[["energie_perdue"	,"manque",	"energie_recuperee",  "cumul_gain (€)", "cumul_perte (€)" ,  "cumul_manque (€)" ]].iloc[-1]
    gain = last_row["cumul_gain (€)"]
    perte = last_row["cumul_perte (€)"]
    manque =last_row["cumul_manque (€)"]
    
    investissement = meta["investissement"]

    ratio = perte / gain if gain else None
    gain_sur_invest = gain / investissement if investissement else None

    resultats.append({
        "scenario": meta["libelleScenario"],
        "investissement": investissement,
         "BesoinTotal(Kwh)": (gain + manque)/prix_kw,
        "gain(Kwh)": gain / prix_kw ,
        "perte(Kwh)":perte / prix_kw,
        "manque(Kwh)": manque / prix_kw,
        
        "BesoinTotal(€)": (gain + manque),
        "gain(€)": gain,
        "perte(€)": perte,
        "manque(€)": manque,
 
        "ratio_perte_gain": ratio,
        "gain_sur_invest": gain_sur_invest,
      "Depense": investissement+ manque,
      "Gain_Sur_depense": gain/ (investissement+ manque),


        
        "lien_excel": f'=HYPERLINK("{nom_final}", "Ouvrir Excel")'
    })

# Résumé final
print("\n📊 Résumé des scénarios :\n")
for res in resultats:
    print(f"🔹 {res['scenario']}")
    print(f"    Investissement     : {res['investissement']} €")
    print(f"    Gain cumulé        : {res['gain(€)']:.2f} €")
    print(f"    Perte cumulée      : {res['perte(€)']:.2f} €")
    print(f"    manque cumulée      : {res['manque(€)']:.2f} €")
    print(f"    Ratio perte/gain   : {res['ratio_perte_gain']:.2%}" if res['ratio_perte_gain'] is not None else "    ⚠️ Ratio perte/gain : gain = 0")
    print(f"    Gain / Invest      : {res['gain_sur_invest']:.2%}" if res['gain_sur_invest'] is not None else "    ⚠️ Impossible de calculer le rendement")
    print("")

# Export final des résultats
df_resultats = pd.DataFrame(resultats)
df_resultats.to_excel("resultats_scenarios.xlsx", index=False)
print("✅ Résumé global exporté dans 'resultats_scenarios.xlsx'")

