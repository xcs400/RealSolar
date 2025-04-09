import cv2
import numpy as np
from openpyxl import Workbook
from openpyxl.chart import ScatterChart, Series, Reference
from scipy.interpolate import interp1d


# prendre une photo en mode panoramique  ( imperatif de rester horizontal pendant la rotation ..) 
# enregistrer en bpm
# noter le nom du fichier dans 'image_path'
# tracer un trait rouge pur (255,0,0) pour delimiter les masques (de gauche a droite de l'image sans revenir en arriere)
# tracer deux points bleu pur(0,0,255 ) sur deux objets et notez leur azimut ci dessous      azimut_start , azimut_end
# mesurer sur le terrain et notez la hauteur min et max de la courbe touge tracé
# noter dans     Max_azMin    ,    Max_azMax   la zone ou le max de la courbe rouge doit etre rechercher.
# noter dans     Min_azMin    ,    Min_azMax   la zone ou le min de la courbe rouge doit etre rechercher.
#lanver la molinette

def extract_red_contour_interpolated(image_path, output_excel, nbvaleur_max=360):


 #   azimut_start = 61   #portail
 #   azimut_end = 260
 #   hauteur_max = 20
#    hauteur_min = 0


    azimut_start = 118   #bas terrain
    azimut_end = 327
    hauteur_max = 25
    hauteur_min = 0
    
    img = cv2.imread(image_path)
    if img is None:
        print(f"Erreur: Impossible de charger l'image {image_path}")
        return

    height, width = img.shape[:2]
    print(f"Image: {width}px x {height}px")

    # === Recherche des deux points bleus (255, 0, 0) ===
    first_blue = None
    second_blue = None

    for x in range(width):
        for y in range(height):
            b, g, r = img[y, x]
            if (b, g, r) == (255, 0, 0):
                first_blue = (x, y)
                break
        if first_blue:
            break

    for x in range(width - 1, -1, -1):
        for y in range(height):
            b, g, r = img[y, x]
            if (b, g, r) == (255, 0, 0):
                second_blue = (x, y)
                break
        if second_blue:
            break

    if not first_blue or not second_blue:
        print("❌ Points bleus non trouvés → azimut inchangé.")
        return

    x1 = first_blue[0]
    x2 = second_blue[0]

    if x1 == x2:
        print("⚠️ Les deux points bleus sont sur la même colonne ! Impossible de calculer l'échelle.")
        return

    def x_to_azimut(x):
        return azimut_start + ((x - x1) / (x2 - x1)) * (azimut_end - azimut_start)

    # === Extraction des points rouges ===
    lower_red = np.array([0, 0, 250])
    upper_red = np.array([50, 50, 255])
    mask = cv2.inRange(img, lower_red, upper_red)

    raw_data = []
    all_red_y = []
    for x in range(width):
        red_pixels = np.where(mask[:, x] > 0)[0]
        if len(red_pixels) > 0:
            y = red_pixels[0]
            all_red_y.append(y)
            azimuth = x_to_azimut(x)
            raw_data.append((azimuth, y))

    if not raw_data:
        print("Aucun point rouge détecté.")
        return

    y_min_red = min(all_red_y)
    y_max_red = max(all_red_y)
    print(f"Courbe rouge détectée entre les lignes {y_min_red} (haut) et {y_max_red} (bas)")

    # Recherche du max et du min dans des zones restreintes
  #  Max_azMin = 100   portail
 #   Max_azMax = 250
 #   Min_azMin = 100
 #   Min_azMax = 250

    Max_azMin = 110   #sapin jy
    Max_azMax = 120
    Min_azMin = 180
    Min_azMax = 300

    
    max_zone = [pt for pt in raw_data if Max_azMin <= pt[0] <= Max_azMax]
    max_point = min(max_zone, key=lambda t: t[1]) if max_zone else min(raw_data, key=lambda t: t[1])

    min_zone = [pt for pt in raw_data if Min_azMin <= pt[0] <= Min_azMax]
    min_point = max(min_zone, key=lambda t: t[1]) if min_zone else max(raw_data, key=lambda t: t[1])

    print(f"Zone de recherche Min: {Min_azMin}° à {Min_azMax}°")
    print(f"→ Point min détecté: azimut={min_point[0]:.2f}°, y={min_point[1]}, élévation brute (avant échelle)=?.")

    y_max_detected = max_point[1]
    y_min_detected = min_point[1]

    if y_max_detected == y_min_detected:
        print("⚠️ Impossible de recalculer l'échelle (min = max) !")
        return

    def y_to_elevation(y):
        return hauteur_min + ((y - y_min_detected) / (y_max_detected - y_min_detected)) * (hauteur_max - hauteur_min)

    raw_data = [(az, y_to_elevation(y)) for az, y in raw_data]

    extended_data = list(raw_data)
    min_point_scaled = (min_point[0], y_to_elevation(min_point[1]))
    max_point_scaled = (max_point[0], y_to_elevation(max_point[1]))

    for special_point in [min_point_scaled, max_point_scaled]:
        if special_point[0] not in [az for az, _ in extended_data]:
            extended_data.append(special_point)

    extended_data.sort()
    azimuths_ext, elevations_ext = zip(*extended_data)

    interpolator = interp1d(azimuths_ext, elevations_ext, kind='linear', bounds_error=False, fill_value=None)

    azimut_first = min(azimuths_ext)
    azimut_last = max(azimuths_ext)
    elevation_first = float(interpolator(azimut_first))
    elevation_last = float(interpolator(azimut_last))

    delta_az = (360 - azimut_last + azimut_first) if azimut_last > azimut_first else (azimut_first - azimut_last)
    slope = (elevation_first - elevation_last) / delta_az

    def extrapolate(az):
        if az > azimut_last:
            delta = az - azimut_last
            return elevation_last + slope * delta
        elif az < azimut_first:
            delta = (360 - azimut_last) + az
            return elevation_last + slope * delta
        else:
            return float(interpolator(az))

    azimut_interp = np.linspace(0, 360, nbvaleur_max)
    elevation_interp = [extrapolate(az) for az in azimut_interp]

    # Écriture Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "Interpolation"

    step_deg = 360 / nbvaleur_max
    ws.append([f"Image: {width}px x {height}px"])
    ws.append([f"Azimut visible: {azimut_first:.2f}° à {azimut_last:.2f}°",
               f"Hauteur: {hauteur_min}° à {hauteur_max}°"])
    ws.append([f"Nb points interpolés: {nbvaleur_max}", f"Pas: {step_deg:.3f}°"])
    ws.append([])
    ws.append(['Azimut (°)', 'Hauteur (°)'])

    for az, el in zip(azimut_interp, elevation_interp):
        ws.append([az, el])

    # Ajout graphique
    chart = ScatterChart()
    chart.title = "Ligne rouge - Interpolée (0 à 360°)"
    chart.x_axis.title = "Azimut (°)"
    chart.y_axis.title = "Hauteur (°)"
    chart.y_axis.scaling.orientation = "minMax"

    filtered_data = [(az, el) for az, el in zip(azimut_interp, elevation_interp) if el is not None]
    start_row = ws.max_row + 2
    ws.append([])
    ws.append(['Azimut_clean', 'Hauteur_clean'])

    for az, el in filtered_data:
        ws.append([az, el])

    min_r = start_row + 1
    max_r = start_row + len(filtered_data)

    x_vals = Reference(ws, min_col=1, min_row=min_r, max_row=max_r)
    y_vals = Reference(ws, min_col=2, min_row=min_r, max_row=max_r)
    series = Series(y_vals, x_vals, title=None)
    series.marker.symbol = "none"
    series.graphicalProperties.line.width = 20000
    chart.series.append(series)
    ws.add_chart(chart, "E8")

    wb.save(output_excel)
    print(f"Fichier Excel généré : {output_excel}")

    csv_path = output_excel.replace('.xlsx', '_pvgis.csv')
    with open(csv_path, 'w', encoding='utf-8') as f:
        for el in elevation_interp:
            f.write(f"{el:.4f}\n")

    print(f"Fichier CSV pour PVGIS généré : {csv_path}")

if __name__ == "__main__":
 #   image_path = "IMG_5655trait.bmp"  #portail
 #   output_excel = "interpolationportail_360deg.xlsx"

    image_path = "IMG_5667basTrait.bmp"  #bas terrain
    output_excel = "interpolationpBasTerrain_360deg.xlsx"
    extract_red_contour_interpolated(image_path, output_excel)
