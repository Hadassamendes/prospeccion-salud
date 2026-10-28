#!/usr/bin/env python3
"""Genera las búsquedas del día rotando barrios/zonas para no repetir siempre los mismos negocios.

Cada línea sale en el formato que entiende el scraper:
    odontólogo en El Poblado, Medellín #!# Medellín|odontólogo|El Poblado
Lo que va después de #!# vuelve en la columna input_id y nos dice ciudad, nicho y zona.
"""
import datetime
import json
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent


def main() -> None:
    config = json.loads((BASE / "config.json").read_text(encoding="utf-8"))
    salida = Path(sys.argv[1]) if len(sys.argv) > 1 else BASE / "datos" / "busquedas_hoy.txt"
    salida.parent.mkdir(parents=True, exist_ok=True)

    dia = datetime.date.today().toordinal()
    por_dia = max(1, int(config.get("zonas_por_dia", 1)))

    lineas = []
    for ciudad, zonas in config["ciudades"].items():
        if not zonas:
            zonas = [""]
        inicio = (dia * por_dia) % len(zonas)
        zonas_hoy = [zonas[(inicio + i) % len(zonas)] for i in range(min(por_dia, len(zonas)))]
        for zona in zonas_hoy:
            lugar = f"{zona}, {ciudad}" if zona else ciudad
            for nicho in config["nichos"]:
                lineas.append(f"{nicho} en {lugar} #!# {ciudad}|{nicho}|{zona}")

    salida.write_text("\n".join(lineas) + "\n", encoding="utf-8")
    print(f"{len(lineas)} búsquedas para hoy -> {salida}")
    for linea in lineas[:: len(config['nichos'])]:
        print("  zona:", linea.split("#!#")[1].strip())


if __name__ == "__main__":
    main()
