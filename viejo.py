

from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pickle
from pprint import pprint


# def main():
#     import gspread
#     from gspread_formatting import *
#     from oauth2client.service_account import ServiceAccountCredentials
#     from df2gspread import df2gspread as d2g
#     from pandas.tseries.offsets import MonthEnd
#     from googleapiclient import discovery
#     from httplib2 import Http
#     from oauth2client import file, client, tools
#     from googleapiclient.discovery import build
#     from gspread_pandas import Spread, Client, conf
#     import os.path, os
#     # iniciar la configuración y credenciales
#     scope = ["https://spreadsheets.google.com/feeds", " https://www.googleapis.com/auth/drive"]
#     credentials = ServiceAccountCredentials.from_json_keyfile_name("credential.json", scope)
#     gc = gspread.authorize(credentials)
#     # cargar book
#     spreadsheet_name = "inversion"
#     inversion_book = gc.open(spreadsheet_name)
#     # el id del archivo
#     spreadsheet_id = inversion_book.id
#     print("spreadsheet id: {}".format(spreadsheet_id))
#     # cargar sheet
#     sheet_name = inversion_book.worksheet("movimientos")
#     # guardar en un dataframe
#     df = pd.DataFrame(sheet_name.get_all_records())
#     organizarDatos(df)
#     # crear nuevo dataframe para exportar
#     nuevo = iterarMeses(df)
#     # crear hoja y cargar
#     new_sheet_name = "Organizado"
#     # cargar dataframe
#     scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
#              "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
#     credentials = ServiceAccountCredentials.from_json_keyfile_name("client_secret1.json", scope)
#     client = gspread.authorize(credentials)
#     spreadsheet_name = "inversion"
#     sheet = client.open(spreadsheet_name)
#     c = conf.get_config('[Your Path]', '[Your filename]')
#     spread = Spread(credentials, spreadsheet_name, config=c)
#     spread.df_to_sheet(df=df, sheet=new_sheet_name)
#
#     # dar formato a header
#     sheet = sheet_name.worksheet("Organizado")
#     # get id from worksheet
#     worksheet_id = leer_id(sheet_name, new_sheet_name)
#     print("Organizado id: {}".format(worksheet_id))
#
#     # EDIT SERVICE
#     service = build('sheets', 'v4', credentials=credentials)
#
#     # request = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=reqs)
#     # response = request.execute()
#     # print(response)

# delete spreadsheet
# id = '1umOvOOPLruUzFJQBeGX8jvRBYQFJbokGigPKVM1iwMo'
# gc.del_spreadsheet(id)
reformar_celdas = {'requests':
    [
        {
            "updateSheetProperties":
            {
                "properties":
                    {
                        "sheetId": worksheet_id,
                        "gridProperties":
                            {
                                'frozenRowCount': 1,
                                "rowCount": 500,
                                "columnCount": 7,
                            }
                    },
                "fields":
                    {
                        "gridProperties(rowCount, columnCount)",
                        'gridProperties.frozenRowCount'
                    }
            }
        },
        {
            "repeatCell":
            {
                "range": #F2:G500 -
                {
                    'startRowIndex': 2,
                    'endRowIndex': 500,
                    'startColumnIndex': 5,
                    'endColumnIndex': 7,
                },
                'cell':
                {
                    'userEnteredFormat':
                    {
                        'numberFormat':
                            {
                                'type': 'CURRENCY',
                                'pattern': '"$"#,##0.00',
                            },
                    },
                },
            }
        }
    ]
}

reqs = {
    'requests': [
    # cambiar tamaño de columnas c - g
    {
      "updateDimensionProperties": {
        "range": {
          "sheetId": worksheet_id,
          "dimension": "COLUMNS",
          "startIndex": 2,
          "endIndex": 10
        },
        "properties": {
          "pixelSize": 120
        },
        "fields": "pixelSize"
      }
    },
    # cambiar tamaño de columna b
    {
        "updateDimensionProperties": {
            "range": {
                "sheetId": worksheet_id,
                "dimension": "COLUMNS",
                "startIndex": 1,
                "endIndex": 2
            },
            "properties": {
                "pixelSize": 200
            },
            "fields": "pixelSize"
        }
    },
    # cambiar tamaño de columna a
    {
        "updateDimensionProperties": {
            "range": {
                "sheetId": worksheet_id,
                "dimension": "COLUMNS",
                "startIndex": 0,
                "endIndex": 1
            },
            "properties": {
                "pixelSize": 120
            },
            "fields": "pixelSize"
        }
    },
    # cambiar tamaño de fila 0
    {
        "updateDimensionProperties": {
            "range": {
                "sheetId": worksheet_id,
                "dimension": "ROWS",
                "startIndex": 0,
                "endIndex": 1
            },
            "properties": {
                "pixelSize": 40
            },
            "fields": "pixelSize"
        }
    },
    # frozen row 1
    {
        'updateSheetProperties': {
        'properties': {"sheetId": worksheet_id, 'gridProperties': {'frozenRowCount': 1}},
        'fields': 'gridProperties.frozenRowCount',
    }},
    # NUMBER TYPE
    {
        'repeatCell':
        {
            'range':
            {
                "sheetId": worksheet_id,
                'startRowIndex': 1,
                'endRowIndex': 500,
                'startColumnIndex': 2,
                'endColumnIndex': 10,
            },
            'cell':
                {
                'userEnteredFormat':
                    {
                        'numberFormat':
                        {
                            'type': 'NUMBER',
                            'pattern': "#,###"
                        },
                    },
                },
            'fields': 'userEnteredFormat.numberFormat',
        },

    },
    # FORMATO DE FECHA
    {
        "repeatCell":
        {
            "range": {
              "sheetId": worksheet_id,
              "startRowIndex": 1,
              "endRowIndex": 500,
              "startColumnIndex": 0,
              "endColumnIndex": 2
            },
            "cell": {
              "userEnteredFormat": {
                "numberFormat": {
                  "type": "DATE",
                  "pattern": "yyyy mmmm dddd"
                }
              }
            },
            "fields": "userEnteredFormat.numberFormat"
        }
    },
    # FORMATO DE CELDAS INTERNARS
    {
        "repeatCell":
        {
            "range": {
              "sheetId": worksheet_id,
              "startRowIndex": 1,
              "endRowIndex": 500,
              "startColumnIndex": 0,
              "endColumnIndex": 10
            },
            "cell": {
                "userEnteredFormat":
                {
                    "textFormat":
                    {
                        "fontFamily": "verdana",
                        "fontSize": 12,
                    }
                }
            },
            "fields": "userEnteredFormat.textFormat"
        }
    },
    # FORMATO DE CELDAS HEADER
    {
        "repeatCell":
            {
                "range": {
                    "sheetId": worksheet_id,
                    "startRowIndex": 0,
                    "endRowIndex": 1,
                },
                "cell": {
                    "userEnteredFormat":
                        {
                            "horizontalAlignment": "CENTER",
                            "verticalAlignment": "MIDDLE",
                            "textFormat":
                            {
                                "fontFamily": "verdana",
                                "fontSize": 14,
                                "bold": True,
                                "foregroundColor":
                                {
                                    "red": 1.0,
                                    "green": 1.0,
                                    "blue": 1.0
                                },
                            },
                            "backgroundColor":
                            {
                                "red": 0.5,
                                "green": 0.5,
                                "blue": 0.5
                            },
                        }
                },
                "fields": "userEnteredFormat.(backgroundColor,textFormat,horizontalAlignment, verticalAlignment)"
            }
    },
]}

reqs = {
    'requests':
        [
        # NUMBER TYPE
        {
            'repeatCell':
                {
                    'range':
                        {
                            "sheetId": worksheet_id,
                            'startRowIndex': 1,
                            'endRowIndex': 500,
                            'startColumnIndex': 7,
                            'endColumnIndex': 8,
                        },
                    'cell':
                        {
                            'userEnteredFormat':
                                {
                                    'numberFormat':
                                        {
                                            'type': 'CURRENCY',
                                            'pattern': '"$"#,##0.00'
                                        },
                                },
                        },
                    'fields': 'userEnteredFormat.numberFormat',
                },

        },
    ]
}

# number format
reqs = {
        'requests':
            [
                # NUMBER TYPE
                {
                    'repeatCell':
                        {
                            'range':
                                {
                                    "sheetId": worksheet_id,
                                    'startRowIndex': 1,
                                    'endRowIndex': 500,
                                    'startColumnIndex': 7,
                                    'endColumnIndex': 8,
                                },
                            'cell':
                                {
                                    'userEnteredFormat':
                                        {
                                            'numberFormat':
                                                {
                                                    'type': 'NUMBER',
                                                    'pattern': "####.#"
                                                },
                                        },
                                },
                            'fields': 'userEnteredFormat.numberFormat',
                        },

                },
            ]
    }


def organizeData(dataframe):
    # AÑO | MES | MOVIMIENTO | ABONOS | RETIROS | Fecha inicial | Fecha final | Dias | Dias del mes
    fecha = pd.to_datetime(dataframe.Fecha, dayfirst=True)
    dataframe["año"] = fecha.dt.year
    dataframe["mes"] = fecha.dt.month
    dataframe["dia"] = fecha.dt.day
    dataframe["dias_al_mes"] = fecha.dt.days_in_month
    dataframe["dias_efectivos"] = dataframe["dias_al_mes"] - dataframe["dia"]
    # para calcular el inicio y fin de mes
    dataframe["inicio_de_mes"] = fecha.dt.strftime("01/%m/%Y")
    dataframe["fin_de_mes"] = (fecha + MonthEnd(0)).dt.strftime("%d/%m/%Y")
    dataframe["montos"] = dataframe["Monto"].str.replace('[\$,]', '', regex=True).astype(float)
    dataframe["saldo_promedio"] = dataframe["montos"] * dataframe["dias_efectivos"]

    # contiene Deposito
    deposito = dataframe["Descripción"].str.contains("Dep|Abono").map({True: "Deposito"}).dropna()
    rendimiento = dataframe["Descripción"].str.contains("Rendi|intereses").map({True: "Rendimiento"}).dropna()
    pago = dataframe["Descripción"].str.contains("inversion|capital").map({True: "Pagos"}).dropna()
    retiro = dataframe["Descripción"].str.contains("Retiro").map({True: "Retiro"}).dropna()
    inversion = dataframe["Descripción"].str.contains("Inver").map({True: "Inversion"}).dropna()
    # print(deposito, rendimiento, retiro, pago)

    serie = [deposito, rendimiento, pago, retiro, inversion]
    tipo = pd.concat(serie).sort_index()
    dataframe["tipo"] = tipo

    dataframe["abonos"] = dataframe.query("tipo != 'Retiro' and tipo != 'Inversion'")["Monto"]
    dataframe["retiros"] = dataframe.query("tipo == 'Retiro'")["Monto"]
    dataframe["inversion"] = dataframe.query("tipo == 'Inversion'")["Monto"]
    dataframe.fillna(0, inplace=True)




def main():
    """
    Creates a sheet with 7 columns, and creates some formats
    :return: NONE
    """
    # create spread objet con id
    spread = create_spread()

    print("id: {}".format(spreadsheet_id))

    # create dataframe form loaded spreadsheet
    df = spread.sheet_to_df(index=0)
    print("Data frame from sheet:")
    print(df.head())

    # structure new dataframe
    organizeData(df)
    # create a new dataframe to export into destiny sheet
    new_df = createNewDataFrame(df)
    # create new sheet is does not exist
    spread.df_to_sheet(
        df=new_df,
        index=False,
        sheet=DESTINY_SHEET_NAME,
        start="A1",
        replace=True
    )
    # get sheet objet to recover id in format function
    destiny_sheet = spread.find_sheet(DESTINY_SHEET_NAME)

    # formatting
    header_format = CellFormat(
        backgroundColor=Color(0.7, 0.7, 0.7),
        textFormat=TextFormat(
            fontFamily="verdana",
            fontSize=14,
            bold=True,
            foregroundColor=Color(1, 1, 1)
        ),
        wrapStrategy="CLIP",
        horizontalAlignment='CENTER',
        verticalAlignment="MIDDLE"
    )
    format_cell_range(worksheet=destiny_sheet, name="1", cell_format=header_format)

    content_format = CellFormat(
        textFormat=TextFormat(
            fontFamily="verdana",
            fontSize=12,
            bold=False,
            foregroundColor=Color(0, 0, 0)
        ),
    )
    format_cell_range(worksheet=destiny_sheet, name="A2:G100", cell_format=content_format)

    date_format = CellFormat(
        numberFormat=NumberFormat
            (
            type="DATE",
            pattern="dddd, dd mmmm yyyy"
        ),
    )
    format_cell_range(worksheet=destiny_sheet, name="A2:A100", cell_format=date_format)

    currency_format = CellFormat(
        numberFormat=NumberFormat
            (
            type="CURRENCY",
            pattern="$#,###,###,##0.00"
        ),
    )
    format_cell_range(worksheet=destiny_sheet, name="C2:G100", cell_format=currency_format)

    # updateDimension(destiny_sheet, 'cols', 260, "A:B")
    # updateDimension(destiny_sheet, 'cols', 160, "C:G")
    # updateDimension(destiny_sheet, 'rows', 35, "A1:G1")
    # updateDimension(destiny_sheet, 'rows', 25, "A2:G100")

    # PAINT ALL ROWS THAT CONTAINS SALDO FINAL IN B COLUMN
    # formula = '=$B1="Saldo final"'
    # rango_formula = "A1:G100"
    # bg_color = (0.8, 0.9, 1)    # RGB
    # fg_color = (0, 0, 0)  # RGB
    # conditionalFormattingFormula(destiny_sheet, formula, rango_formula, bg_color, fg_color)

    # formula = '=$B1="Saldo inicial"'
    # bg_color = (0.88, 0.93, 0.85)  # RGB
    # conditionalFormattingFormula(destiny_sheet, formula, rango_formula, bg_color, fg_color)



def createNewDataFrame(dataframe):
    # crear un nuevo dataframe
    nuevas_columnas = ["saldo inicial", "saldo final"]
    nuevo_df = pd.DataFrame(index=range(0), columns=nuevas_columnas)

    # buscar el nombre del mes en el diccionario
    dict_of_df = {}
    nombre_meses = {
        1: "Enero",
        2: "Febero",
        3: "Marzo",
        4: "Abril",
        5: "Mayo",
        6: "Junio",
        7: "Julio",
        8: "Agosto",
        9: "Septiembre",
        10: "Octubre",
        11: "Noviembre",
        12: "Diciembre"
    }
    # empezar con los meses
    dataframe["mes"] = dataframe["mes"].map(nombre_meses)
    meses = dataframe["mes"]
    set_meses = uniqueValues(meses)
    index = 0
    saldo_inicial_mensual = 0
    # iterar por los meses unicos
    for mes in set_meses:
        # hacer un query para sub iterar solo en los meses iterados
        query = "mes == '{}'".format(mes)
        busqueda = dataframe.query(query)
        filtro = busqueda[["Fecha", 'Descripción', "montos", "abonos", "retiros", "inversion"]]
        # print(filtro)
        largo = len(busqueda.mes)
        for subindex in range(largo):
            # calcular el indice del dataframe global
            superindex = subindex + index
            # print(superindex)
            # encontrar el valor de cada celda iterada en el superindex
            date = busqueda.Fecha[superindex]
            saldo = busqueda.montos[superindex]

            if subindex == 0 and superindex != 0:
                saldo_inicial_mensual = suma
                # print("saldo mensual inicial: {:.2f}".format(saldo_inicial_mensual))
                # crear dataframe temporal
                d = {"Descripción": ["Saldo inicial"], "saldo inicial": [saldo_inicial_mensual], "Fecha": [date]}
                temp = pd.DataFrame(data=d)
                nuevo_df = nuevo_df.append(temp)

        # sumar para calcular saldo final de mes
        suma = sum(busqueda.montos) + saldo_inicial_mensual
        # print("saldo mensual final: {:.2f}".format(suma))
        # print()

        # cambiar el valor de index para calcular el superindex
        index += largo

        # AGREGAR EL QUERY AL NUEVO DATA FRAME Y AGREGAR SALDO FINAL
        nuevo_df = nuevo_df.append(filtro)
        # crear dataframe temporal
        d = {"Descripción": ["Saldo final"], "saldo final": [suma], "Fecha": [date]}
        temp = pd.DataFrame(data=d)
        # print(temp)
        nuevo_df = nuevo_df.append(temp)

    # print(nuevo_df)
    nuevo_df.reset_index(drop=True, inplace=True)
    nuevo_df.fillna(0, inplace=True)
    # reorganizar columnas
    cols_organizadas = [
        "Fecha",
        "Descripción",
        "abonos",
        "retiros",
        "inversion",
        "saldo inicial",
        "saldo final"
    ]
    nuevo_df = nuevo_df[cols_organizadas]
    nuevo_df["abonos"] = nuevo_df["abonos"].str.replace('[\$,]', '', regex=True)
    nuevo_df["retiros"] = nuevo_df["retiros"].str.replace('[\$,]', '', regex=True)
    nuevo_df["inversion"] = nuevo_df["inversion"].str.replace('[\$,]', '', regex=True)
    nuevo_df["saldo inicial"] = nuevo_df["saldo inicial"].astype(float)
    nuevo_df["saldo final"] = nuevo_df["saldo final"].astype(float)
    nuevo_df.fillna(0, inplace=True)
    # print(nuevo_df.abonos)

    nuevo_df["abonos"] = pd.to_numeric(nuevo_df["abonos"])
    nuevo_df["retiros"] = pd.to_numeric(nuevo_df["retiros"])
    nuevo_df["inversion"] = pd.to_numeric(nuevo_df["inversion"])
    # print(nuevo_df.abonos)

    # print(nuevo_df)
    return nuevo_df

    # # iterar por meses y contar cuantos rows tienen el mismo mes
    # mes_anterior = ""
    # conteo = 0
    # key = 0
    # for index, mes in enumerate(meses):
    #     if mes != mes_anterior:
    #         if index != 0:
    #             key += 1
    #             name = "new_df_" + str(key)
    #             print("{} crear dataframe de: {} rows".format(name, conteo))
    #             dict_of_df.update({name: pd.DataFrame(index=range(0, conteo), columns=[])})
    #         conteo = 1
    #         mes_anterior = mes
    #         print()
    #         print("nuevo mes encontrado, {}".format(mes))
    #         print("mes: {} conteo: {} index: {}".format(mes, conteo, index))
    #     else:
    #         conteo += 1
    #         print("mes: {} conteo: {} index: {}".format(mes, conteo, index))
    #         if index == len(meses)-1:
    #             key += 1
    #             name = "new_df_" + str(key)
    #             print("{} crear dataframe de: {} rows".format(name, conteo))
    #             dict_of_df.update({name: pd.DataFrame(index=range(0, conteo), columns=[])})
    #
    # # acceder al dataframe desde el diccionario
    # dict_of_df.get("new_df_1")["Mes"] = meses
    # print(dict_of_df.get("new_df_1")["Mes"])


def conditionalFormattingConstains(worksheet, text, rango):
    global service, spreadsheet_id
    startcolumn, startrow, endcolumn, endrow = get_index_from_range(rango)
    print(startcolumn,endcolumn, startrow, endrow)
    req = {
        "requests":
            [
                {
                    "addConditionalFormatRule":
                    {
                        "rule":
                        {
                            "ranges":
                            [{
                                "sheetId": worksheet.id,
                                "startRowIndex": startrow,
                                "endRowIndex": endrow,
                                "startColumnIndex": startcolumn,
                                "endColumnIndex": endcolumn,
                            }],
                            "booleanRule":
                            {
                                "condition":
                                    {
                                        "type": "TEXT_CONTAINS",
                                        "values":
                                            [{
                                                "userEnteredValue": text
                                            }]
                                    },
                                "format":
                                {
                                    "textFormat":
                                    {
                                        "bold": True,
                                        "italic": True,
                                        "foregroundColor":
                                        {
                                            "red": 0.8,
                                            "green": 0.2,
                                            "blue": 0.0
                                        }
                                    },
                                    "backgroundColor":
                                    {
                                        "red": 0.8,
                                        "green": 0.2,
                                        "blue": 0.0
                                    },

                                }
                            },
                        }
                    }
                }
            ]
    }
    request = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=req)
    response = request.execute()
    print(response)


def crear_compartir(gc):
    # crear y compartir sheet
    inversion = gc.create('Defecto')
    email = "cgibranvalle@gmail.com"
    inversion.share(email, perm_type='user', role='writer')


def leer_con_url(gc):
    # leer la gsheet
    url = "https://docs.google.com/spreadsheets/d/1v5hu0l8pkuWieq-SP_jaHj1sY4CGnTWlQpf0H59MFQk/edit#gid=212539596"
    test = gc.open_by_url(url)
    print(test)


def crearHeader(gc):
    # abrir el archivo
    inversion_sheet = gc.open("inversion")
    # crear un nuevo sheet
    # lista_movimientos = inversion_sheet.add_worksheet(title="movimientos", rows=100, cols="3")
    # si existe no crea, solo carga
    lista_movimientos = inversion_sheet.worksheet("movimientos")
    # crear los headers
    lista_movimientos.update("A1", "Fecha")
    lista_movimientos.update("B1", "Descripción")
    lista_movimientos.update("C1", "Monto")
    # dar formato
    formato_columna = CellFormat(
        backgroundColor=Color(0.7, 0.7, 0.7),
        textFormat=TextFormat(bold=True, foregroundColor=Color(1, 1, 1), fontFamily="Verdana", fontSize=14),
        horizontalAlignment='CENTER'
    )
    format_cell_range(lista_movimientos, 'A1:C1', formato_columna)


def formatoHeader(rango, worksheet):
    formato_header = CellFormat(
        backgroundColor=Color(0.7, 0.7, 0.7),
        textFormat=TextFormat(bold=True, foregroundColor=Color(1, 1, 1), fontFamily="Verdana", fontSize=14),
        horizontalAlignment='CENTER'
    )
    format_cell_range(worksheet, rango, formato_header)



def leer_id(spreadsheet_name, sheet):
    # leer el id del spreadsheet
    id = spreadsheet_name.id
    # cargar el sheet
    sheet = spreadsheet_name.worksheet(sheet)
    sheet_id = sheet.id
    return sheet_id


def organizarDatos(dataframe):
    # AÑO | MES | MOVIMIENTO | ABONOS | RETIROS | Fecha inicial | Fecha final | Dias | Dias del mes
    fecha = pd.to_datetime(dataframe.Fecha, dayfirst=True)
    dataframe["año"] = fecha.dt.year
    dataframe["mes"] = fecha.dt.month
    dataframe["dia"] = fecha.dt.day
    dataframe["dias_al_mes"] = fecha.dt.days_in_month
    dataframe["dias_efectivos"] = dataframe["dias_al_mes"] - dataframe["dia"]
    # para calcular el inicio y fin de mes
    dataframe["inicio_de_mes"] = fecha.dt.strftime("01/%m/%Y")
    dataframe["fin_de_mes"] = (fecha + MonthEnd(0)).dt.strftime("%d/%m/%Y")
    dataframe["montos"] = dataframe["Monto"].str.replace('[\$,]', '', regex=True).astype(float)
    dataframe["saldo_promedio"] = dataframe["montos"] * dataframe["dias_efectivos"]

    # contiene Deposito
    deposito = dataframe["Descripción"].str.contains("Dep|Abono").map({True: "Deposito"}).dropna()
    rendimiento = dataframe["Descripción"].str.contains("Rendi|intereses").map({True: "Rendimiento"}).dropna()
    pago = dataframe["Descripción"].str.contains("inversion|capital").map({True: "Pagos"}).dropna()
    retiro = dataframe["Descripción"].str.contains("Retiro").map({True: "Retiro"}).dropna()
    inversion = dataframe["Descripción"].str.contains("Inver").map({True: "Inversion"}).dropna()
    # print(deposito, rendimiento, retiro, pago)

    serie = [deposito, rendimiento, pago, retiro, inversion]
    tipo = pd.concat(serie).sort_index()
    dataframe["tipo"] = tipo

    dataframe["abonos"] = dataframe.query("tipo != 'Retiro' and tipo != 'Inversion'")["Monto"]
    dataframe["retiros"] = dataframe.query("tipo == 'Retiro'")["Monto"]
    dataframe["inversion"] = dataframe.query("tipo == 'Inversion'")["Monto"]
    dataframe.fillna(0, inplace=True)


def cargarDataFrame(dataframe, id, worksheet, credentials):
    cell_of_start_df = 'A1'
    d2g.upload(dataframe, id, worksheet,
               credentials=credentials,
               col_names=True,
               row_names=False,
               start_cell=cell_of_start_df,
               clean=True)


def crearFila(dataframe, superindex, tipo):
    if tipo == "saldo":
        movimiento = "Saldo inicial"
    elif tipo == "abono":
        movimiento = "Abono"
    elif tipo == "retiro":
        movimiento = "Retiro"
    elif tipo == "promedio saldo":
        movimiento = "Saldo promedio saldo inicial"
    elif tipo == "promedio depositos":
        movimiento = "Saldo promedio depositos"
    elif tipo == "promedio retiros":
        movimiento = "Saldo promedio retiros"
    elif tipo == "promedio diario":
        movimiento = "Saldo promedio diario"
    elif tipo == "tasa":
        movimiento = "Tasa de interes anual"
    elif tipo == "rendimiento":
        movimiento = "Rendimiento"
    elif tipo == "subtotal":
        movimiento = "Subtotal"
    elif tipo == "final":
        movimiento = "Saldo mensual"

    month = dataframe.mes[superindex]
    abonos = dataframe.Abonos[superindex]
    retiros = dataframe.Retiros[superindex]
    initial_date = dataframe.Fecha[superindex]
    final_date = dataframe["fin_de_mes"][superindex]
    dias = dataframe["dias_efectivos"][superindex]
    dias_del_mes = dataframe["dias_al_mes"][superindex]
    lista = [month, movimiento, abonos, retiros, initial_date, final_date, dias, dias_del_mes]
    # print(lista)
    return lista


def iterarMeses(dataframe):
    # crear un nuevo dataframe
    nuevas_columnas = ["saldo inicial", "saldo final"]
    nuevo_df = pd.DataFrame(index=range(0), columns=nuevas_columnas)

    # buscar el nombre del mes en el diccionario
    dict_of_df = {}
    nombre_meses = {
        1: "Enero",
        2: "Febero",
        3: "Marzo",
        4: "Abril",
        5: "Mayo",
        6: "Junio",
        7: "Julio",
        8: "Agosto",
        9: "Septiembre",
        10: "Octubre",
        11: "Noviembre",
        12: "Diciembre"
    }
    # empezar con los meses
    dataframe["mes"] = dataframe["mes"].map(nombre_meses)
    meses = dataframe["mes"]
    set_meses = valoresUnicos(meses)
    index = 0
    fecha_inicial = ""
    saldo_inicial = 0
    saldo_inicial_mensual = 0
    # iterar por los meses unicos
    for mes in set_meses:
        # hacer un query para sub iterar solo en los meses iterados
        query = "mes == '{}'".format(mes)
        busqueda = dataframe.query(query)
        filtro = busqueda[["Fecha", 'Descripción', "montos", "abonos", "retiros", "inversion"]]
        # print(filtro)
        largo = len(busqueda.mes)
        for subindex in range(largo):
            # calcular el indice del dataframe global
            superindex = subindex + index
            # print(superindex)
            # encontrar el valor de cada celda iterada en el superindex
            date = busqueda.Fecha[superindex]
            saldo = busqueda.montos[superindex]

            if subindex == 0 and superindex != 0:
                saldo_inicial_mensual = suma
                # print("saldo mensual inicial: {:.2f}".format(saldo_inicial_mensual))
                # crear dataframe temporal
                d = {"saldo inicial": [saldo_inicial_mensual], "Fecha": [date]}
                temp = pd.DataFrame(data=d)
                nuevo_df = nuevo_df.append(temp)

        # sumar para calcular saldo final de mes
        suma = sum(busqueda.montos) + saldo_inicial_mensual
        # print("saldo mensual final: {:.2f}".format(suma))
        # print()

        # cambiar el valor de index para calcular el superindex
        index += largo

        # AGREGAR EL QUERY AL NUEVO DATA FRAME Y AGREGAR SALDO FINAL
        nuevo_df = nuevo_df.append(filtro)
        # crear dataframe temporal
        d = {"saldo final": [suma], "Fecha": [date]}
        temp = pd.DataFrame(data=d)
        # print(temp)
        nuevo_df = nuevo_df.append(temp)

    # print(nuevo_df)
    nuevo_df.reset_index(drop=True, inplace=True)
    nuevo_df.fillna(0, inplace=True)
    # reorganizar columnas
    cols_organizadas = [
        "Fecha",
        "Descripción",
        "abonos",
        "retiros",
        "inversion",
        "saldo inicial",
        "saldo final"
    ]
    nuevo_df = nuevo_df[cols_organizadas]
    nuevo_df["abonos"] = nuevo_df["abonos"].str.replace('[\$,]', '', regex=True)
    nuevo_df["retiros"] = nuevo_df["retiros"].str.replace('[\$,]', '', regex=True)
    nuevo_df["inversion"] = nuevo_df["inversion"].str.replace('[\$,]', '', regex=True)
    nuevo_df["saldo inicial"] = nuevo_df["saldo inicial"].astype(float)
    nuevo_df["saldo final"] = nuevo_df["saldo final"].astype(float)
    nuevo_df.fillna(0, inplace=True)
    print(nuevo_df.abonos)

    nuevo_df["abonos"] = pd.to_numeric(nuevo_df["abonos"])
    nuevo_df["retiros"] = pd.to_numeric(nuevo_df["retiros"])
    nuevo_df["inversion"] = pd.to_numeric(nuevo_df["inversion"])
    print(nuevo_df.abonos)

    # print(nuevo_df)
    return nuevo_df

    # # iterar por meses y contar cuantos rows tienen el mismo mes
    # mes_anterior = ""
    # conteo = 0
    # key = 0
    # for index, mes in enumerate(meses):
    #     if mes != mes_anterior:
    #         if index != 0:
    #             key += 1
    #             name = "new_df_" + str(key)
    #             print("{} crear dataframe de: {} rows".format(name, conteo))
    #             dict_of_df.update({name: pd.DataFrame(index=range(0, conteo), columns=[])})
    #         conteo = 1
    #         mes_anterior = mes
    #         print()
    #         print("nuevo mes encontrado, {}".format(mes))
    #         print("mes: {} conteo: {} index: {}".format(mes, conteo, index))
    #     else:
    #         conteo += 1
    #         print("mes: {} conteo: {} index: {}".format(mes, conteo, index))
    #         if index == len(meses)-1:
    #             key += 1
    #             name = "new_df_" + str(key)
    #             print("{} crear dataframe de: {} rows".format(name, conteo))
    #             dict_of_df.update({name: pd.DataFrame(index=range(0, conteo), columns=[])})
    #
    # # acceder al dataframe desde el diccionario
    # dict_of_df.get("new_df_1")["Mes"] = meses
    # print(dict_of_df.get("new_df_1")["Mes"])


def valoresUnicos(list):
    valores_unicos = []
    largo = len(list)
    valor_anterior = list[largo - 1]
    for dato in list:
        if valor_anterior != dato:
            valores_unicos.append(dato)
            valor_anterior = dato
    return valores_unicos


def generarCredenciales():
    SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
    store = file.Storage('storage.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)
    SHEETS = discovery.build('sheets', 'v4', http=creds.authorize(Http()))
    return SHEETS

# RUN MAIN
#main()
nuevo_main()