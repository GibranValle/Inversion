import os
from oauth2client.service_account import ServiceAccountCredentials
from gspread_pandas import Spread, conf, Client
import gspread
import pandas as pd
from gspread_formatting import *
from pandas.tseries.offsets import MonthEnd
from googleapiclient.discovery import build

# CONSTANTS
SPREADSHEET_NAME = "inversion"
DESTINY_SHEET_NAME = "Organizado"
ORIGIN_SHEET_NAME = "movimientos"
SPREADSHEET_ID = ""

def create_spread():
    # create credentials from local file
    scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
             "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
    credential = ServiceAccountCredentials.from_json_keyfile_name("credential.json", scope)

    # create path
    path = os.getcwd()
    filename = "credential.json"
    path_file = conf.get_config(path, filename)
    # create spread objets from credentail, open origin spread and sheet.
    spread = Spread(creds=credential, spread=SPREADSHEET_NAME, config=path_file, sheet=ORIGIN_SHEET_NAME)

    # create client mode for recovery id
    client = Client(creds=credential, scope=scope)
    # recover list of dictionaries
    list_dicts = client.list_spreadsheet_files()
    # iterate all spreadsheets created by credencial user
    for dict in list_dicts:
        print(dict)
        name = dict.get("name")
        if name == SPREADSHEET_NAME:
            spreadsheet_id = dict.get("id")
            print("el id del sp es: {}".format(spreadsheet_id))
    SPREADSHEET_ID = spreadsheet_id

    # create SERVICE to create json request
    service = build('sheets', 'v4', credentials=credential)
    return spread, service


def main():
    # create spread objet con id
    spread, service = create_spread()

    # create dataframe form loaded spreadsheet
    df = spread.sheet_to_df(index=0)
    print("Data frame from sheet:")
    print(df.head())

    # structure new dataframe
    organizeData(df)
    # create a new dataframe to export into destiny sheet
    new_df = iterarMeses(df)
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

    #formatting
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
            pattern="$#,###,###,###.00"
        ),
    )
    format_cell_range(worksheet=destiny_sheet, name="C2:G100", cell_format=currency_format)

    updateDimension(destiny_sheet, 'cols', 260, "A:B", service)
    updateDimension(destiny_sheet, 'cols', 160, "C:G", service)


def updateDimension(worksheet, element, pixels, rango, service):
    # TEST PATTERN
    # rango = "A1"
    # a, b, c, d = get_index_from_range(rango)
    # print("{} sci: {} sri:{} eci:{} eri:{}".format(rango, a, b, c, d))
    #
    # rango = "A1:A2"
    # a, b, c, d = get_index_from_range(rango)
    # print("{} sci: {} sri:{} eci:{} eri:{}".format(rango, a, b, c, d))
    #
    # rango = "A:A"
    # a, b, c, d = get_index_from_range(rango)
    # print("{} sci: {} sri:{} eci:{} eri:{}".format(rango, a, b, c, d))
    #
    # rango = "A:B100"
    # a, b, c, d = get_index_from_range(rango)
    # print("{} sci: {} sri:{} eci:{} eri:{}".format(rango, a, b, c, d))
    #
    # rango = "1:3"
    # a, b, c, d = get_index_from_range(rango)
    # print("{} sci: {} sri:{} eci:{} eri:{}".format(rango, a, b, c, d))
    #
    # rango = "1"
    # a, b, c, d = get_index_from_range(rango)
    # print("{} sci: {} sri:{} eci:{} eri:{}".format(rango, a, b, c, d))

    startcolumn, startrow, endcolumn, endrow = get_index_from_range(rango)

    if element == 'cols':
        dimension = "COLUMNS"
        start = startcolumn
        end = endcolumn
    elif element == 'rows':
        dimension = "ROWS"
        start = startrow
        end = endrow

    req = {
        "requests":
        [
            {
                "updateDimensionProperties":
                {
                    "range":
                    {
                        "sheetId": worksheet.id,
                        "dimension": dimension,
                        "startIndex": start,
                        "endIndex": end
                    },
                    "properties":
                    {
                        "pixelSize": pixels
                    },
                    "fields": "pixelSize"
                }
            }
        ]
    }
    request = service.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body=req)
    response = request.execute()
    print(response)


def get_index_from_range(ranges):
    startColumnIndex, startRowIndex, endColumnIndex, endRowIndex = (0, 0, 0, 0)
    values = ranges.split(":")
    # print()
    # print("values: {}".format(values))
    last_column_index = 0
    last_row_index = 0
    list_size = len(values)
    #print(list_size)

    for index, value in enumerate(values):
        value_size = len(value)
        #print("value size: {}".format(value_size))
        if list_size == 1:  #single value
            startRowIndex = ord(value[:1]) - 65
            startRowIndex += 16 if startRowIndex < 0 else startRowIndex
            endRowIndex = startRowIndex + 1
            #print("startRowIndex: {}".format(startRowIndex))
            if value_size == 1:
                startColumnIndex = 0
                # TODO ENCONTAR EL FINAL DEL INDEX
                endColumnIndex = 1000
                # print("endRowIndex: {}".format(endRowIndex))
                # print("startColumnIndex: {}".format(startColumnIndex))
                # print("endColumnIndex: {}".format(endColumnIndex))
                # print("ENTIRE ROW {}".format(startRowIndex))
            elif value_size == 2:
                startColumnIndex = ord(value[:1]) - 65
                endColumnIndex = startColumnIndex + 1
                # print("endRowIndex: {}".format(endRowIndex))
                # print("startColumnIndex: {}".format(startColumnIndex))
                # print("endColumnIndex: {}".format(endColumnIndex))
                # print("column {} row {}".format(value[:1], value[1:]))

        else: # 2 list value
            if value_size <= 1:
                if index == 0:
                    number = ord(value[:1]) - 65
                    if number < 0:
                        startColumnIndex = 0
                        startRowIndex = number + 16
                        last_row_index = startRowIndex
                        # print("ENTIRE ROW")
                        # print("from column {}".format(startColumnIndex))
                        # print("from row {}".format(startRowIndex))
                    else:
                        startColumnIndex = number
                        last_column_index = startColumnIndex
                        startRowIndex = 0
                        # print("ENTIRE COLUMN")
                        # print("from column {} undefined row".format(startColumnIndex))
                        # print("from row {}".format(startRowIndex))
                elif index == 1:
                    number = ord(value[:1]) - 65
                    if number < 0:
                        # TODO FIND END COLUMNS
                        endColumnIndex = 1000
                        endRowIndex = number + 16
                        endRowIndex += 1
                        # print("to column {}".format(endColumnIndex))
                        # print("to row {}".format(endRowIndex))
                    else:
                        endColumnIndex = ord(value[:1]) - 65
                        endColumnIndex += 1
                        # TODO find last row
                        endRowIndex = 1000
                        # print("to column {}".format(endColumnIndex))
                        # print("to row {}".format(endRowIndex))

            else:       # value_size = 2 list_value = 2
                column = ord(value[:1]) - 65
                row = int(value[1:]) - 1
                if index == 0:
                    startColumnIndex = column
                    last_column_index = startColumnIndex
                    startRowIndex = row
                    last_row_index = startRowIndex
                    # print("from column {}".format(value))
                    # print("startColumnIndex: {}".format(startColumnIndex))
                    # print("startRowIndex: {}".format(startRowIndex))
                elif index == 1:
                    endColumnIndex = column
                    if endColumnIndex == last_column_index:
                        endColumnIndex += 1
                    endRowIndex = row
                    endRowIndex += 1
                    # print("to column {}".format(value))
                    # print("endColumnIndex: {}".format(endColumnIndex))
                    # print("endRowIndex: {}".format(endRowIndex))
    return startColumnIndex, startRowIndex, endColumnIndex, endRowIndex


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
    #print(nuevo_df.abonos)

    nuevo_df["abonos"] = pd.to_numeric(nuevo_df["abonos"])
    nuevo_df["retiros"] = pd.to_numeric(nuevo_df["retiros"])
    nuevo_df["inversion"] = pd.to_numeric(nuevo_df["inversion"])
    #print(nuevo_df.abonos)

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


# RUN MAIN
main()
