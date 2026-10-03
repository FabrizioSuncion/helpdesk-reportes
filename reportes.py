"""Genera un reporte en Excel con estadisticas de los tickets de HelpDesk.

Se conecta a la base de datos HelpDeskDb (SQL Server) y crea un archivo .xlsx
con un resumen, tablas por estado, prioridad y categoria, graficos y el
detalle de todos los tickets.
"""

import argparse
from datetime import datetime
from pathlib import Path

import pandas as pd
import pyodbc

SERVIDOR = r"(localdb)\MSSQLLocalDB"
BASE_DATOS = "HelpDeskDb"
ZONA_HORARIA = "America/Lima"

PRIORIDADES = {1: "Baja", 2: "Media", 3: "Alta", 4: "Critica"}
ESTADOS = {1: "Abierto", 2: "En proceso", 3: "Resuelto", 4: "Cerrado"}

CONSULTA = """
SELECT t.Id, t.Titulo, t.Prioridad, t.Estado, t.FechaCreacion, t.FechaCierre,
       c.Nombre AS Categoria,
       u.NombreCompleto AS Creador,
       tec.NombreCompleto AS Tecnico
FROM Tickets t
JOIN Categorias c ON c.Id = t.CategoriaId
JOIN AspNetUsers u ON u.Id = t.CreadorId
LEFT JOIN AspNetUsers tec ON tec.Id = t.TecnicoId
ORDER BY t.Id
"""


def elegir_driver():
    disponibles = pyodbc.drivers()
    for nombre in ("ODBC Driver 18 for SQL Server", "ODBC Driver 17 for SQL Server"):
        if nombre in disponibles:
            return nombre
    raise RuntimeError("No se encontro un controlador ODBC de SQL Server (17 o 18).")


def obtener_tickets(servidor, base_datos):
    cadena = (
        f"DRIVER={{{elegir_driver()}}};SERVER={servidor};DATABASE={base_datos};"
        "Trusted_Connection=yes;Encrypt=no;TrustServerCertificate=yes;"
    )
    conexion = pyodbc.connect(cadena)
    try:
        cursor = conexion.cursor()
        cursor.execute(CONSULTA)
        columnas = [columna[0] for columna in cursor.description]
        filas = [list(fila) for fila in cursor.fetchall()]
    finally:
        conexion.close()
    return pd.DataFrame(filas, columns=columnas)


def a_hora_local(serie):
    return (
        pd.to_datetime(serie)
        .dt.tz_localize("UTC")
        .dt.tz_convert(ZONA_HORARIA)
        .dt.tz_localize(None)
    )


def preparar(df):
    df = df.copy()
    df["Prioridad"] = df["Prioridad"].map(PRIORIDADES)
    df["Estado"] = df["Estado"].map(ESTADOS)
    df["Tecnico"] = df["Tecnico"].fillna("Sin asignar")
    df["FechaCreacion"] = a_hora_local(df["FechaCreacion"])
    df["FechaCierre"] = a_hora_local(df["FechaCierre"])
    df["HorasResolucion"] = (
        df["FechaCierre"] - df["FechaCreacion"]
    ).dt.total_seconds() / 3600
    return df


def construir_resumen(df):
    resueltos = df[df["FechaCierre"].notna()]
    promedio = resueltos["HorasResolucion"].mean()
    indicadores = [
        ("Total de tickets", len(df)),
        ("Tickets abiertos", int((df["Estado"] == "Abierto").sum())),
        ("Tickets en proceso", int((df["Estado"] == "En proceso").sum())),
        ("Tickets resueltos o cerrados", int(df["Estado"].isin(["Resuelto", "Cerrado"]).sum())),
        ("Tickets sin tecnico asignado", int((df["Tecnico"] == "Sin asignar").sum())),
        (
            "Tiempo promedio de resolucion (horas)",
            round(float(promedio), 2) if pd.notna(promedio) else "Sin datos",
        ),
        ("Fecha del reporte", datetime.now().strftime("%d/%m/%Y %H:%M")),
    ]
    return pd.DataFrame(indicadores, columns=["Indicador", "Valor"])


def contar(df, columna, orden=None):
    conteo = df[columna].value_counts()
    if orden is not None:
        conteo = conteo.reindex(orden, fill_value=0)
    return conteo.rename_axis(columna).reset_index(name="Tickets")


def escribir_excel(df, ruta):
    resumen = construir_resumen(df)
    por_estado = contar(df, "Estado", list(ESTADOS.values()))
    por_prioridad = contar(df, "Prioridad", list(PRIORIDADES.values()))
    por_categoria = contar(df, "Categoria")
    detalle = df[
        [
            "Id", "Titulo", "Categoria", "Prioridad", "Estado", "Creador",
            "Tecnico", "FechaCreacion", "FechaCierre", "HorasResolucion",
        ]
    ].copy()
    detalle["HorasResolucion"] = detalle["HorasResolucion"].round(2)

    with pd.ExcelWriter(ruta, engine="xlsxwriter", datetime_format="dd/mm/yyyy hh:mm") as escritor:
        libro = escritor.book
        formato_encabezado = libro.add_format(
            {"bold": True, "bg_color": "#1F3864", "font_color": "#FFFFFF", "border": 1}
        )

        def escribir_tabla(datos, nombre_hoja, fila=0, columna=0):
            datos.to_excel(
                escritor, sheet_name=nombre_hoja, index=False, startrow=fila, startcol=columna
            )
            hoja = escritor.sheets[nombre_hoja]
            for i, nombre in enumerate(datos.columns):
                hoja.write(fila, columna + i, nombre, formato_encabezado)
                largos = [len(str(nombre))] + [len(str(valor)) for valor in datos.iloc[:, i]]
                hoja.set_column(columna + i, columna + i, min(max(largos) + 2, 45))
            return hoja

        escribir_tabla(resumen, "Resumen")

        hoja_estadisticas = escribir_tabla(por_estado, "Estadisticas", 0, 0)
        escribir_tabla(por_prioridad, "Estadisticas", 0, 3)
        escribir_tabla(por_categoria, "Estadisticas", 0, 6)

        def grafico(titulo, datos, columna, celda):
            n = len(datos)
            g = libro.add_chart({"type": "column"})
            g.add_series(
                {
                    "name": titulo,
                    "categories": ["Estadisticas", 1, columna, n, columna],
                    "values": ["Estadisticas", 1, columna + 1, n, columna + 1],
                    "data_labels": {"value": True},
                }
            )
            g.set_title({"name": titulo})
            g.set_legend({"none": True})
            g.set_size({"width": 420, "height": 280})
            hoja_estadisticas.insert_chart(celda, g)

        grafico("Tickets por estado", por_estado, 0, "A8")
        grafico("Tickets por prioridad", por_prioridad, 3, "A24")
        grafico("Tickets por categoria", por_categoria, 6, "A40")

        hoja_detalle = escribir_tabla(detalle, "Tickets")
        hoja_detalle.autofilter(0, 0, len(detalle), len(detalle.columns) - 1)
        hoja_detalle.freeze_panes(1, 0)


def main():
    parser = argparse.ArgumentParser(description="Genera un reporte en Excel de los tickets de HelpDesk.")
    parser.add_argument("--servidor", default=SERVIDOR, help="Servidor de SQL Server")
    parser.add_argument("--base", default=BASE_DATOS, help="Nombre de la base de datos")
    parser.add_argument("--salida", default="reportes", help="Carpeta donde se guarda el reporte")
    args = parser.parse_args()

    df = obtener_tickets(args.servidor, args.base)
    if df.empty:
        print("No hay tickets en la base de datos. No se genero el reporte.")
        return

    df = preparar(df)
    carpeta = Path(args.salida)
    carpeta.mkdir(exist_ok=True)
    ruta = carpeta / f"reporte_tickets_{datetime.now():%Y-%m-%d_%H%M}.xlsx"
    escribir_excel(df, ruta)
    print(f"Reporte generado: {ruta.resolve()}")


if __name__ == "__main__":
    main()
