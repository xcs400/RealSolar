# RealSolarCost  (en python3)
Calcul l'amortissement de différentes options de pose de panneaux solaires.

Basé sur les données d'ensoleillement PVGIS-SARAH3    https://re.jrc.ec.europa.eu  .
- supposition du calcul  :
   - on veut effacer "le bruit de fond" de la consommation de la maison  ( ex: 100w  )   et chauffer un ballon d'eau chaude sanitaire .
   - la production va en priorité à la conso de la maison et puis au ballon d'eau chaude .  Le surplus s'il y en a, va au réseau (considéré comme perdu).
   - on suppose donc l'existence d'un routeur solaire pour envoyer ce qui est disponible dans le ballon,  et ne pas le chauffer avec le réseau.
   - le "manque" journalier est intégré dans le calcul (car il faudra le payer au fournisseur d'électricité pour compléter les apports nécessaires,  Dépense=investissement+manque) . 
    Ce "manque" est considéré que pendant 12 heures de conso de la maison + chauffe-eau complet ( afin de comparer l'effet de l'ajout de panneaux supplémentaires).    

  Le calcul se fait heure par heure, jour par jour sur la période considérée à partir des données d'ensoleillement réelles observées dans la base PVGIS-SARAH pour un lieu donné
  (plus de 20 ans d'historique).  l'amortissement se calcule à partir d'une date donnée, à renseigner (3 ou 5 dernières années par exemple ..)

- il faut renseigner les variables suivantes dans le code.

  - lat, lon = 45.xxxx, 2.xxxx       #position de la maison
  - date_departcumul = pd.Timestamp("2020-01-01")   # date de début du calcul ;  date de pose des panneaux
  - startyear, endyear = 2020, 2023  # date de la période du calcul (doit au minimum couvrir date_departcumul)
  - prix_kw = 0.2016  # €/kWh
  - conso_chauffeau_journalière = 2.2            # kWh /Jour  . correspond à deux douches journalières chez moi.
  - un ou des profils d'horizons s'il existe des masques d'ombrage importants (voir plus bas).
                                                  
- Définir les différents scénarios de pose et d'orientation des panneaux dans le tableau : scenarios
ex:

    -    [ 
     -   {"libelleScenario": "1xPortail_2xBalcon", "investissement": 120+80+80+190 +80+125,"conso_maison_W":conso_maison_W},
     -   {"angle": 72, "aspect": -18, "MaxPower": 0.5, "TrackerType": 0 , "horizonType": 1, "libelle": "balcon_72"},
     -   {"angle": 72, "aspect": -18, "MaxPower": 0.5, "TrackerType": 0 , "horizonType": 1, "libelle": "balcon_72"},
     -   {"angle": 37, "aspect": 0, "MaxPower": 0.5, "TrackerType": 0 , "horizonType": 2, "libelle": "optimal"},
       ]

    Dans investissement prévoir le coût: panneau(x), micro-onduleur(s), routeur solaire , fixation, câble..etc

   - 'conso_maison_W'  est la conso "bruit de fond" de la maison à couvrir .
   - 'angle' est l'orientation de pose du panneau (90 est vertical),
   - 'aspect' (mal nommé)  est l'azimut (0= sud;  -15:15 degrés vers l'est,  90:plein ouest) 
   - 'MaxPower' est la puissance crête du panneau .
   - 'libellé' préfixe du nom du fichier intermédiaire.
   - 'TrackerType'  tracking    0= fixe   2= 2 axes    3= axe vertical qui tourne     5= axe incliné qui tourne  (voir PVGIS)
   - 'horizonType' 0 = use default location   autre indice dans la table des profils d'horizon.

* le calcul génère des fichiers excel avec les données journalières pour chaque panneau et un cumul global de tous les panneaux ( plus un graphique journalier pour chaque mois).
* un fichier resultats_scenarios.xlsx est généré avec le résumé des calculs et amortissements avec un lien vers chaque fichier excel des données+graphiques.

![image](https://github.com/user-attachments/assets/0e66fc52-4cb5-41b3-a0c5-251aa8ea181c)

on retrouve:
 * scenario	investissement
 * BesoinTotal(Kwh)
 * gain(Kwh)
 * perte(Kwh)
 * manque(Kwh)
 * BesoinTotal(€)
 * gain(€)
 * perte(€)          (non récupéré, retour au réseau)
 *  manque(€)
 * ratio_perte_gain
 * gain_sur_invest
 * Dépense          (investissement + manque) 
 * Gain_Sur_dépense
 * lien_excel
 * 
 Pour chaque scénario le lien excel pointe vers le fichier datas et graphiques:
![image](https://github.com/user-attachments/assets/32c4af2c-8623-498b-9990-a247e0c09bab)

Des onglets supplémentaires sont generé pour résumer les elements du calcul:
  - les paramètres scénarios.
     
    ![image](https://github.com/user-attachments/assets/49bb57f0-c3a2-422e-91ac-1f77757ad5ba)

  - les diagrammes solaires avec les masques de hauteur d'horizon incluant les masques réels  (générés en sphérique cartésien , un onglet excel par graphique et par horizons )
    
![image](https://github.com/user-attachments/assets/79a7ed9d-4400-4d1a-a777-eb7f83e98b6e)
  
![image](https://github.com/user-attachments/assets/ba4fdb90-431f-43fb-8eda-35f1018bbce5)

## Comment relever un profil d'horizon ?
Pour tracer des profils d'horizon, utiliser le script extractprofile.py ( à adapter avec vos paramètres) .
Il faut être capable de relever des azimuts avec une boussole et des 'élévations' de points remarquables avec inclinometre.
Ceci peut se faire avec un simple téléphone et un peu de méthode (qui dispose de boussole et d'inclinomètre) , il faut attacher le télephone à une règle permettant de 'viser et pointer' un objet particulier.
- prendre une photo en mode panoramique du lieu du/des panneaux avec votre téléphone portable.
  Il est impératif de rester horizontal pendant toute la rotation du téléphone... il faut réaliser un support ou utiliser un trépied rotatif.
- enregistrer/renommer le fichier en .BMP (couleur 24bit).
- noter le nom du fichier dans 'image_path'
- tracer un trait rouge pur (255,0,0) dans l'image pour délimiter les masques (de gauche à droite de l'image sans revenir en arrière, sans discontinuité) .
- tracer deux points bleus purs(0,0,255 ) sur deux objets et noter leur azimut (à mesurer sur le terrain avec la boussole en tenant compte de la déclinaison magnétique) dans 'azimut_start' , 'azimut_end'  (ceci indique l'échelle horizontale).
- mesurer et noter l'élévation des points haut et bas de votre trait rouge (mesurer sur le terrain avec le clinomètre et noter dans 'hauteur_min' et 'hauteur_max' ) .
- noter dans     Max_azMin    ,    Max_azMax   la zone où le max de la courbe rouge doit être recherché.
- noter dans     Min_azMin    ,    Min_azMax   la zone où le min de la courbe rouge doit être recherché.
- lancer la moulinette extractprofile.py,  inclure les points générés dans le fichier .xlsx ('output_excel' = "votrenom.xlsx") sous forme de chaîne dans le tableau de profil de realsolarcost.py
example d'image modifiée.
![image](https://github.com/user-attachments/assets/3f6b9138-7ab1-428a-8b33-867da9b4214a)





# translated    
# RealSolarCost  (in Python 3)
Calculates the amortization of different solar panel installation options.

Based on PVGIS-SARAH3 sunlight data https://re.jrc.ec.europa.eu.
- Assumptions for the calculation:
   - The goal is to offset the household's "background noise" consumption (e.g., 100W) and heat a domestic hot water tank.
   - Production is first used for household consumption, then for heating water. Any surplus (if any) is sent to the grid (considered as lost).
   - It is assumed that a solar router is installed to send the available energy to the water tank and prevent it from being heated by the grid.
   - The daily "shortage" is included in the calculation (since it must be purchased from the electricity provider to cover the shortfall. Expense = investment + shortage).
     This "shortage" is only considered during the 12 hours of household consumption + full water heating (to evaluate the effect of adding more panels).

  The calculation is performed hour by hour, day by day over the selected period, using real historical sunlight data from the PVGIS-SARAH database for a given location
  (more than 20 years of data). Amortization is calculated from a given starting date (e.g., last 3 or 5 years).

- You need to set the following variables in the code:

  - lat, lon = 45.xxxx, 2.xxxx       # house location
  - date_departcumul = pd.Timestamp("2020-01-01")   # calculation start date; panel installation date
  - startyear, endyear = 2020, 2023  # period of calculation (must at least cover date_departcumul)
  - prix_kw = 0.2016  # €/kWh
  - conso_chauffeau_journalière = 2.2  # kWh /day. Corresponds to two daily showers in my case.
  - One or more horizon profiles if significant shading/masking exists (see below).

- Define different panel installation and orientation scenarios in the `scenarios` list.
Example:

    -    [ 
     -   {"libelleScenario": "1xGate_2xBalcony", "investissement": 120+80+80+190 +80+125,"conso_maison_W":conso_maison_W},
     -   {"angle": 72, "aspect": -18, "MaxPower": 0.5, "TrackerType": 0 , "horizonType": 1, "libelle": "balcony_72"},
     -   {"angle": 72, "aspect": -18, "MaxPower": 0.5, "TrackerType": 0 , "horizonType": 1, "libelle": "balcony_72"},
     -   {"angle": 37, "aspect": 0, "MaxPower": 0.5, "TrackerType": 0 , "horizonType": 2, "libelle": "optimal"},
       ]

    Investment should include costs for: panel(s), micro-inverter(s), solar router, mounting, cables, etc.

   - 'conso_maison_W' is the household's background consumption to be covered.
   - 'angle' is the tilt of the panel (90 means vertical),
   - 'aspect' (misnamed) is the azimuth (0 = south; -15:15 degrees toward east, 90 = due west)
   - 'MaxPower' is the panel’s peak power.
   - 'libelle' is the prefix used for naming intermediate files.
   - 'TrackerType': 0 = fixed, 2 = dual axis, 3 = vertical axis tracker, 5 = inclined axis tracker (see PVGIS)
   - 'horizonType': 0 = use default location; any other index refers to the custom horizon profile table.

* The calculation generates Excel files with daily data for each panel and a global cumulative file (with a daily graph per month).
* A `resultats_scenarios.xlsx` file is generated summarizing the amortization and performance with a link to each data+graph Excel file.

![image](https://github.com/user-attachments/assets/0e66fc52-4cb5-41b3-a0c5-251aa8ea181c)

Included:
 * scenario, investment
 * TotalNeed(KWh)
 * gain(KWh)
 * loss(KWh)
 * shortage(KWh)
 * TotalNeed(€)
 * gain(€)
 * loss(€)      (not recovered, returned to grid)
 * shortage(€)
 * loss_gain_ratio
 * gain_on_investment
 * TotalExpense   (investment + shortage)
 * Gain_on_Expense
 * excel_link
 
Each scenario’s Excel link points to the data and graph file:
![image](https://github.com/user-attachments/assets/32c4af2c-8623-498b-9990-a247e0c09bab)

Additional tabs are generated to summarize the calculation elements:
  - scenario parameters.
     
    ![image](https://github.com/user-attachments/assets/49bb57f0-c3a2-422e-91ac-1f77757ad5ba)

  - solar charts with actual horizon shading masks (generated in spherical Cartesian form, one Excel tab per chart and per horizon)
    
![image](https://github.com/user-attachments/assets/79a7ed9d-4400-4d1a-a777-eb7f83e98b6e)
  
![image](https://github.com/user-attachments/assets/ba4fdb90-431f-43fb-8eda-35f1018bbce5)

## How to record a horizon profile?
To draw horizon profiles, use the `extractprofile.py` script (to be adapted with your settings).
You must be able to record azimuths with a compass and elevation angles of landmarks using an inclinometer.
This can be done with a smartphone and a bit of method (smartphones have a compass and inclinometer). Attach the phone to a straight ruler to "aim and point" at a specific object.
- Take a panoramic photo of the panel area with your smartphone.
  It is crucial to stay level during the entire rotation... Use a tripod or build a rotating stand.
- Save/rename the file as .BMP (24-bit color).
- Write the filename in `image_path`.
- Draw a pure red line (255,0,0) on the image to outline the masks (from left to right with no gaps or backtracking).
- Draw two pure blue points (0,0,255) on two landmarks and record their azimuth (measured on-site with compass, considering magnetic declination) in `azimut_start` and `azimut_end` (this defines the horizontal scale).
- Measure and note the elevation of the top and bottom of your red line (with inclinometer) and enter in `hauteur_min` and `hauteur_max`.
- Enter the zone where the red curve's max should be found: `Max_azMin`, `Max_azMax`.
- Enter the zone for min: `Min_azMin`, `Min_azMax`.
- Run `extractprofile.py`, then include the generated points in your `.xlsx` file (`output_excel` = "yourname.xlsx") as a string in the horizon profile table of `realsolarcost.py`.

Example of a modified image:
![image](https://github.com/user-attachments/assets/3f6b9138-7ab1-428a-8b33-867da9b4214a)
