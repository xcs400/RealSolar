# RealSolar  (en python3)
Calcul l'amortissement de differentes options de pose de panneaux solaires. 

Basé sur les données d'ensoleilement PVGIS-SARAH3    https://re.jrc.ec.europa.eu  .
- supposition du calcul  :
   - on veux effacer "le bruit de fond" de la consomation de la maison  ( ex: 100w  )   et chauffer un ballon d'eau chauffe sanitaire .
   - la production va en priorité à la conso de la maison et puis au ballon d'eau chaude .  Le surplus si il y en a, va au réseau (consideré comme perdu).
   - on suppose donc l'existance d'un routeur solaire pour envoyer ce qui est disponible dans le ballon,  et ne pas le chauffer avec le réseau.
   - le "manque" jounalier est integré dans le calcul (car il faudra le payer au fournisseur d'electricité pour completer les apports necessaires) . 
    Ce "manque" est consideré que pendant la durée d'ensoleillement ( afin de comparer l'effet de l'ajout d'un panneau supplementaire).    
     
  Le calcul se fait heure par heure, jour par jour sur la periode consideré à partir des donnée d'ensoleilement reel observe dans la base PVGIS-SARAH pour un lieu donné.
  (plus de 20 ans d'historique).  l'amortissement se calcul à partir d'une date donnée, à renseigné (3 ou 5 derniere années par example ..)

- il faut renseigner les variables suivantes dans le code.

  - lat, lon = 45.xxxx, 2.xxxx       #position de la maison
  - date_departcumul = pd.Timestamp("2020-01-01")   # date de debut du calcul ;  date de pose des paneaux
  - startyear, endyear = 2020, 2023  # date de la periode du calcul (doit au minimum couvrir date_departcumul)
  - prix_kw = 0.2016  # €/kWh
  - conso_chauffeau_journalière = 2.2            # kWh /Jour  . correspond a deux douches jounalieres chez moi.
                                                  
- Definir les differents scenarios de pose et d'orientation des panneaux dans le tableau : scenarios
ex:
   -  [  # 2 panneaux toiture vers l'est et ouest 
    -  {"libelleScenario": "2xPanneaux_Toiture_VersOuest", "investissement": 120+80+80+190,"conso_maison_W":100},
    -  {"angle": 23, "aspect": -90, "MaxPower": 0.5, "libelle": "est"},
    -  {"angle": 23, "aspect": 90 , "MaxPower": 0.5, "libelle": "ouest"},
       ],

    Dans investissement prevoir le coût: panneau(x), micro-onduleur(s),routeur solaire , fixation, cable..etc

   - 'conso_maison_W'  est la conso "bruit de fond" de la maison à couvrir .
   - 'angle' est l'orientation de pose du panneau (90 est vertical),
   - 'aspect' (mal nommé)  est l'azimute (0= sud;  -15:15 degree vers l'est,  90:plein ouest) 
   - 'MaxPower' est la puissance crête du panneau .
   - 'libellé' prefix du nom du fichier intermedaire.

* le calcul genere des fichiers excel avec les données journalieres pour chaque panneau et un cumul globlal de tous les panneaux ( plus un graphique journalier pour chaque mois).
* un fichier resultats_scenarios.xlsx est generé avec le resumé des calculs et ammortissements avec un lien vers chaque fichier excel des données+graphiques.

![image](https://github.com/user-attachments/assets/0e66fc52-4cb5-41b3-a0c5-251aa8ea181c)


on retrouve:
 * scenario	investissement
 * BesoinTotal(Kwh)
 * gain(Kwh)
 * perte(Kwh)
 * manque(Kwh)
 * BesoinTotal(€)
 * gain(€)
 * perte(€)          (non recuperé, retour au reseau)
 *  manque(€)
 * ratio_perte_gain
 * gain_sur_invest
 * Depense          (investissement + manque) 
 * Gain_Sur_depense
 * lien_excel
 * 
 Pour chaque scenario le lien excel pointe vers le fichier datas et graphiques:
![image](https://github.com/user-attachments/assets/32c4af2c-8623-498b-9990-a247e0c09bab)


Des onglets supplementaires resumes:
  - les parametres scenarios.
  -    
    ![image](https://github.com/user-attachments/assets/49bb57f0-c3a2-422e-91ac-1f77757ad5ba)

  - les diagrammes solaire avec les masques de hauteur d'horizon
   ![image](https://github.com/user-attachments/assets/978bf3ad-2aef-4475-af41-dc7414103b84)




  
    
