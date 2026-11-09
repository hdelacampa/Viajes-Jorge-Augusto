Página interactiva de los viajes de Jorge y Augusto.
Se actualiza con `python3 build.py source.xlsx` (o `refresh.sh`) a partir de la pestaña Viajes.

Cada crucero es un viaje: sus tramos, días y millas se suman. Un mismo barco en otro año es otro viaje. El generador no lee la pestaña Resumen.

GitHub Pages publica `index.html`. Si ese archivo está en la misma carpeta, `build.py` lo reescribe junto con `viajes_jorge_augusto.html`.
