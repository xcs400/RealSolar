import cv2
import numpy as np
from openpyxl import Workbook
from openpyxl.chart import ScatterChart, Series, Reference
from scipy.interpolate import interp1d

def extract_red_contour_interpolated(image_path, output_excel, nbvaleur_max=360):
    # Paramètres angulaires
    azimut_start = 50
    azimut_end = 256
    hauteur_max = 20
    hauteur_min = 0

    # Charger l'image
    img = cv2.imread(image_path)
    if img is None:
        print(f"Erreur: Impossible de charger l'image {image_path}")
        return

    height, width = img.shape[:2]
    print(f"Image: {width}px x {height}px")

    # Rouge avec tolérance
    lower_red = np.array([0, 0, 250])
    upper_red = np.array([50, 50, 255])
    mask = cv2.inRange(img, lower_red, upper_red)

    raw_data = []
    for x in range(width):
        red_pixels = np.where(mask[:, x] > 0)[0]
        if len(red_pixels) > 0:
            y = red_pixels[0]
            azimuth = azimut_start + (x / (width - 1)) * (azimut_end - azimut_start)
            elevation = hauteur_max - (y / (height - 1)) * (hauteur_max - hauteur_min)
            raw_data.append((azimuth, elevation))

    # Tri des points par azimut
    raw_data.sort()
    azimuths, elevations = zip(*raw_data)

    # Interpolation linéaire
    interpolator = interp1d(azimuths, elevations, kind='linear', bounds_error=False, fill_value=None)

    # Générer une grille de 0 à 360° avec pas constant
    azimut_interp = np.linspace(0, 360, nbvaleur_max)
    elevation_interp = []

    # Obtenir les hauteurs aux extrémités connues
    elevation_start = float(interpolator(azimut_start))
    elevation_end = float(interpolator(azimut_end))

    # Calculer pente et ordonnée à l'origine de la droite qui relie azimut_end à azimut_start en passant par 360→0
    if azimut_end > azimut_start:
        # Long chemin autour du cercle
        delta_az = (360 - azimut_end) + azimut_start
    else:
        delta_az = azimut_start - azimut_end

    slope = (elevation_start - elevation_end) / delta_az

    def extrapolate(az):
        if az > azimut_end:
            delta = az - azimut_end
            return elevation_end + slope * delta
        elif az < azimut_start:
            delta = (360 - azimut_end) + az
            return elevation_end + slope * delta
        else:
            return float(interpolator(az))  # zone interpolée normalement

    elevation_interp = [extrapolate(az) for az in azimut_interp]


    # Écriture Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "Interpolation"

    step_deg = 360 / nbvaleur_max
    ws.append([f"Image: {width}px x {height}px"])
    ws.append([f"Azimut visible: {azimut_start}° à {azimut_end}°",
               f"Hauteur: {hauteur_min}° à {hauteur_max}°"])
    ws.append([f"Nb points interpolés: {nbvaleur_max}", f"Pas: {step_deg:.3f}°"])
    ws.append([])
    ws.append(['Azimut (°)', 'Hauteur (°)'])

    for az, el in zip(azimut_interp, elevation_interp):
        ws.append([az, el])

    # Graphique (sans les None)
    chart = ScatterChart()
    chart.title = "Ligne rouge - Interpolée (0 à 360°)"
    chart.x_axis.title = "Azimut (°)"
    chart.y_axis.title = "Hauteur (°)"
    chart.y_axis.scaling.orientation = "minMax"

    # On crée une version "nettoyée" des données pour le graphique, sans valeurs None
    filtered_data = [(az, el) for az, el in zip(azimut_interp, elevation_interp) if el is not None]
    start_row = ws.max_row + 2
    ws.append([])
    ws.append(['Azimut_clean', 'Hauteur_clean'])

    for az, el in filtered_data:
        ws.append([az, el])

    # Ajouter la courbe unique
    chart = ScatterChart()
    chart.title = "Ligne rouge - Interpolée (0 à 360°)"
    chart.x_axis.title = "Azimut (°)"
    chart.y_axis.title = "Hauteur (°)"
    chart.y_axis.scaling.orientation = "minMax"

    min_r = start_row + 1
    max_r = start_row + len(filtered_data)

    x_vals = Reference(ws, min_col=1, min_row=min_r, max_row=max_r)
    y_vals = Reference(ws, min_col=2, min_row=min_r, max_row=max_r)
    series = Series(y_vals, x_vals, title=None)
    series.marker.symbol = "none"
    series.graphicalProperties.line.width = 20000  # plus épais (optionnel)
    chart.series.append(series)
    ws.add_chart(chart, "E8")



    wb.save(output_excel)
    print(f"Fichier Excel généré : {output_excel}")
    print(f"Interpolation sur {nbvaleur_max} points réguliers de 0° à 360°")


    # === Export CSV pour PVGIS ===
    csv_path = output_excel.replace('.xlsx', '_pvgis.csv')

    with open(csv_path, 'w', encoding='utf-8') as f:
        
        # Ligne 6 et suivantes : uniquement les hauteurs
        for el in elevation_interp:
            f.write(f"{el:.4f}\n")

    print(f"Fichier CSV pour PVGIS généré : {csv_path}")



if __name__ == "__main__":
    image_path = "IMG_5655trait.bmp"
    output_excel = "interpolationportail_360deg.xlsx"
    extract_red_contour_interpolated(image_path, output_excel)
