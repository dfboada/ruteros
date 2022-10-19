# -*- coding: utf-8 -*-
# Triqui estándar

import base64
from zipfile import ZipFile

# Importo paquetes
import numpy as np
import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.units import cm, inch
from reportlab.platypus import *
from reportlab.platypus.flowables import Flowable


class verticalText(Flowable):

    # Rotates a text in a table cell

    def __init__(self, text):
        Flowable.__init__(self)
        self.text = text

    def draw(self):
        canvas = self.canv
        canvas.rotate(90)
        fs = canvas._fontsize
        canvas.translate(1, -fs / 1.2)  # canvas._leading?
        canvas.drawString(0, 0, self.text)

    def wrap(self, aW, aH):
        canv = self.canv
        fn, fs = canv._fontname, canv._fontsize
        return canv._leading, 1 + canv.stringWidth(self.text, fn, fs)


# Set title
st.title("Generador de Triquis", anchor=None)

# Importo el excel subido con rutero
dataset = st.file_uploader(
    "Cargue rutero con 3 columnas: Vendedor (Zona), Cliente, día de la semana (1-7) ",
    type=["xlsx"],
)
if dataset is not None:
    dfrut = pd.read_excel(dataset)
    st.sidebar.write(" ### Rows and Columns:", dfrut.shape)

# Importo el excel subido con las SKUs a medir
datasetsku = st.file_uploader(
    "Cargue nombre referencias a medir (1era columna: max 20 SKUs (max 15 caracteres), 2nda columna: 1 si pertenece a Champions, 0 de lo contrario)",
    type=["xlsx"],
)
if datasetsku is not None:
    dfsku = pd.read_excel(datasetsku)
    st.sidebar.write(" ### Rows and Columns:", dfsku.shape)

# Agrego el Boton de generar
resultado = st.button("Generar")  # Devuleve True cuando el usuario hace click

if resultado:

    # dfrut = pd.read_excel("data/Rutero.xlsx")
    # dfsku = pd.read_excel("data/Productos.xlsx")

    # Rutero
    dfrut.rename(
        columns={
            list(dfrut)[0]: "Vendedor",
            list(dfrut)[1]: "Cliente",
            list(dfrut)[2]: "Dia",
        },
        inplace=True,
    )
    dfrut["Cliente"] = dfrut["Cliente"].str[:20]

    dfsku.rename(columns={list(dfsku)[0]: "SKU"}, inplace=True)
    dfsku.rename(columns={list(dfsku)[1]: "Champions"}, inplace=True)
    dfsku["SKU"] = dfsku["SKU"].str[:15]

    # Zonas/Vendedores unicos
    zonas = dfrut.iloc[:, 0].unique()
    cantvendedores = zonas.shape[0]

    # Productos Unicos
    productos = dfsku.iloc[:, 0].unique()
    cantprod = productos.shape[0]

    dfunico = pd.DataFrame()

    days = {1: "Lun", 2: "Mar", 3: "Mie", 4: "Jue", 5: "Vie", 6: "Sab", 7: "Dom"}

    for i in range(1, cantvendedores + 1):
        # Segmento cada vendedor y organizo por alfabeto y día
        dfunico = dfrut[(dfrut["Vendedor"] == zonas[i - 1])]
        dfunico = dfunico.sort_values(by=["Dia", "Cliente"], ascending=True)

        for num in range(7):
            dfunico["Dia"] = np.where(
                dfunico["Dia"] == num + 1, days[num + 1], dfunico["Dia"]
            )
            dfunico["Dia"] = np.where(
                dfunico["Dia"] == f"{num+1}", days[num + 1], dfunico["Dia"]
            )

        # Agrego los productos hacia la derecha
        for k in range(0, min(cantprod, 20)):
            dfunico[productos[k]] = ""

        dfunico = dfunico.drop(["Vendedor"], axis=1)

        # Imprimir en PDF con reportlab
        canvas = SimpleDocTemplate(
            f"Zona {zonas[i - 1]}.pdf",
            pagesize=landscape(letter),
            topMargin=1 * cm,
            bottomMargin=1 * cm,
        )
        titulo = f"Zona {zonas[i - 1]}"
        c_width = [4 * cm] + [1 * cm] * 16  # width of the columns
        r_width = [3 * cm] + [0.5 * cm] * (dfunico.shape[0])
        data1 = dfunico.values.tolist()

        headcols = []
        for i, namecol in enumerate(dfunico.columns):
            if i <= 1:
                headcols.append(str(namecol))
            else:
                headcols.append(verticalText(str(namecol)))

        data1.insert(0, headcols)
        t = Table(data1, colWidths=c_width, repeatRows=1, rowHeights=r_width)
        t.setStyle(
            TableStyle(
                [
                    ("FONTSIZE", (0, 0), (-1, -1), 9),
                    ("BACKGROUND", (0, 0), (-1, 0), colors.lightblue),
                    ("VALIGN", (0, 0), (-1, 0), "BOTTOM"),
                ]
            )
        )

        for a in range(cantprod):
            if dfsku["Champions"][a]:
                t.setStyle(
                    TableStyle(
                        [
                            ("FONTSIZE", (0, 0), (-1, -1), 9),
                            ("BACKGROUND", (a + 2, 0), (a + 2, 0), colors.black),
                            ("VALIGN", (0, 0), (-1, 0), "BOTTOM"),
                            ("TEXTCOLOR", (a + 2, 0), (a + 2, 0), colors.white),
                        ]
                    )
                )

        GRID_STYLE = TableStyle(
            [
                ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.black),
                ("BOX", (0, 0), (-1, -1), 0.25, colors.black),
                ("LINEABOVE", (0, 1), (-1, -1), 0.25, colors.black),
            ]
        )

        t.setStyle(GRID_STYLE)
        elements = []
        elements.append(Paragraph(titulo))
        elements.append(t)
        # canvas.rotate(90)
        canvas.build(elements)

    # Crear ZIP file
    zipObj = ZipFile("triquiResult.zip", "w")

    # Add multiple files to the zip
    for w in range(1, cantvendedores + 1):
        zipObj.write(f"Zona {zonas[w-1]}.pdf")

    # close the Zip File
    zipObj.close()

    ZipfileDotZip = "triquiResult.zip"

    with open(ZipfileDotZip, "rb") as f:
        bytes = f.read()
        b64 = base64.b64encode(bytes).decode()
        href = f"<a href=\"data:file/zip;base64,{b64}\" download='{ZipfileDotZip}'>\
            Descarga Formatos\
        </a>"
    st.sidebar.markdown(href, unsafe_allow_html=True)
