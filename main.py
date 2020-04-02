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
CALCULATOR_SHEET = "calculator"
AZUL = (0.88, 0.95, 1)    # RGB azulito
NEGRO = (0, 0, 0)
VERDE = (0.88, 0.93, 0.85)
ROJO = (1, 0.8, 0.8)
BLANCO = (1, 1, 1)
GRIS = (0.9, 0.9, 0.9)
# GLOBAL VARIABLES
service = ""
spreadsheet_id = ""


def create_math_sheet():
    """
    Creates a sheet for calculations of return
    :return: None
    """
    # create spread objet con id
    spread = create_spread()
    print("id: {}".format(spreadsheet_id))

    # create dataframe form loaded spreadsheet
    origin_df = spread.sheet_to_df(index=0)
    print("Data frame from sheet:")
    print(origin_df.head())

    # structure new dataframe
    organizeMathData(origin_df)
    # export this dataframe to analyze
    createNewSheet(spread, origin_df, "TEST")

    new_df = createMathDataFrame(origin_df)
    print(new_df)
    createNewSheet(spread, new_df, CALCULATOR_SHEET)

    # formatting
    # get sheet objet to recover id in format function
    destiny_sheet = spread.find_sheet(CALCULATOR_SHEET)
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

    percent_format = CellFormat(
        numberFormat=NumberFormat
        (
            type="PERCENT",
            pattern="#0.0#%"
        ),
    )
    format_cell_range(worksheet=destiny_sheet, name="C", cell_format=percent_format)

    currency_format = CellFormat(
        numberFormat=NumberFormat
        (
            type="CURRENCY",
            pattern="$#,###,###,##0.00"
        ),
    )
    format_cell_range(worksheet=destiny_sheet, name="D:E", cell_format=currency_format)

    updateDimension(destiny_sheet, 'cols', 140, "A:G")
    updateDimension(destiny_sheet, 'cols', 350, "B:B")
    updateDimension(destiny_sheet, 'rows', 35, "A1:G1")
    updateDimension(destiny_sheet, 'rows', 25, "A2:G100")

    #PAINT ALL ROWS THAT CONTAINS SALDO FINAL IN B COLUMN
    formula = '=SEARCH("ingreso",$B1)'
    priority = 2
    rango_formula = "A1:G100"
    bg_color = VERDE
    fg_color = NEGRO
    conditionalFormattingFormula(destiny_sheet, formula, rango_formula, bg_color, fg_color, priority)

    formula = '=SEARCH("egreso",$B1)'
    bg_color = ROJO
    conditionalFormattingFormula(destiny_sheet, formula, rango_formula, bg_color, fg_color, priority)

    formula = '=SEARCH("inicial",$B1)'
    bg_color = AZUL
    conditionalFormattingFormula(destiny_sheet, formula, rango_formula, bg_color, fg_color, priority)

    formula = '=SEARCH("rendimiento",$B1)'
    priority = 0
    bg_color = VERDE
    rango_formula = "A1:D100"
    conditionalFormattingFormula(destiny_sheet, formula, rango_formula, bg_color, fg_color, priority)

    formula = '=$B1="Saldo promedio mensual de ingresos"'
    bg_color = VERDE
    conditionalFormattingFormula(destiny_sheet, formula, rango_formula, bg_color, fg_color, priority)

    formula = '=$B1="Saldo promedio mensual de saldo inicial"'
    bg_color = AZUL
    conditionalFormattingFormula(destiny_sheet, formula, rango_formula, bg_color, fg_color, priority)

    formula = '=$B1="Saldo promedio mensual de egresos"'
    bg_color = ROJO
    conditionalFormattingFormula(destiny_sheet, formula, rango_formula, bg_color, fg_color, priority)

    formula = '=SEARCH("tasa",$B1)'
    bg_color = GRIS
    conditionalFormattingFormula(destiny_sheet, formula, rango_formula, bg_color, fg_color, priority)

    formula = '=SEARCH("total",$B1)'
    conditionalFormattingFormula(destiny_sheet, formula, rango_formula, bg_color, fg_color, priority)

    formula = '=$B1="Saldo promedio mensual"'
    conditionalFormattingFormula(destiny_sheet, formula, rango_formula, bg_color, fg_color, priority)

    formula = '=$B1="Total"'
    conditionalFormattingFormulaBold(destiny_sheet, formula, rango_formula, bg_color, fg_color)



def create_spread():
    # global variables for other methods
    global spreadsheet_id, service

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

    # start client mode and create SERVICE to send json request, before that get ID
    client = gspread.authorize(credential)
    inversion_book = client.open(SPREADSHEET_NAME)
    spreadsheet_id = inversion_book.id
    service = build('sheets', 'v4', credentials=credential)
    return spread


def updateDimension(worksheet, element, pixels, rango):
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
    # global variables from create spread
    global service, spreadsheet_id

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
    request = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=req)
    response = request.execute()
    print(response)


def conditionalFormattingFormula(worksheet, formula, rango, background, foreground, priority):
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
                                        "type": "CUSTOM_FORMULA",
                                        "values":
                                            [{
                                                "userEnteredValue": formula
                                            }]
                                    },
                                "format":
                                {
                                    "textFormat":
                                    {
                                        "bold": False,
                                        "italic": False,
                                        "foregroundColor":
                                        {
                                            "red": foreground[0],
                                            "green": foreground[1],
                                            "blue": foreground[2]
                                        }
                                    },
                                    "backgroundColor":
                                    {
                                        "red": background[0],
                                        "green": background[1],
                                        "blue": background[2]
                                    },

                                }
                            },
                        },
                        "index": priority
                    }
                }
            ]
    }
    request = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=req)
    response = request.execute()
    print(response)


def conditionalFormattingFormulaBold(worksheet, formula, rango, background, foreground):
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
                                        "type": "CUSTOM_FORMULA",
                                        "values":
                                            [{
                                                "userEnteredValue": formula
                                            }]
                                    },
                                "format":
                                {
                                    "textFormat":
                                    {
                                        "bold": True,
                                        "italic": False,
                                        "foregroundColor":
                                        {
                                            "red": foreground[0],
                                            "green": foreground[1],
                                            "blue": foreground[2]
                                        }
                                    },
                                    "backgroundColor":
                                    {
                                        "red": background[0],
                                        "green": background[1],
                                        "blue": background[2]
                                    },

                                }
                            },
                        },
                        "index": 0
                    }
                }
            ]
    }
    request = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=req)
    response = request.execute()
    print(response)


def organizeMathData(dataframe):
    """
    FINAL REVISION: 27/03/2020 WORKING CORRECTLY
    Edits the dataframe with columns
    AÑO | MES | MOVIMIENTO | TASA DE INTERES | INGRESOS | EGRESOS | FECHA INICIAL | FECHA FINAL | DIAS DEL MES | DIAS CONTADOS

    :param dataframe:
    :return: none, however the dataframe is edited
    """
    fecha = pd.to_datetime(dataframe.Fecha, dayfirst=True)
    dataframe["año"] = fecha.dt.year
    dataframe["mes"] = fecha.dt.month
    dataframe["dia"] = fecha.dt.day
    dataframe["dias al mes"] = fecha.dt.days_in_month
    dataframe["dias efectivos"] = dataframe["dias al mes"] - dataframe["dia"] + 1
    dataframe["tasa de interes"] = dataframe["Tasa de interes"].str.replace('[\%,]', '', regex=True).astype(float)/100
    dataframe["inicio de mes"] = fecha.dt.strftime("01/%m/%Y")
    dataframe["fin de mes"] = (fecha + MonthEnd(0)).dt.strftime("%d/%m/%Y")
    dataframe["montos"] = dataframe["Monto"].str.replace('[\$,]', '', regex=True).astype(float)
    dataframe["saldo promedio"] = dataframe["montos"] * dataframe["dias efectivos"] / dataframe["dias al mes"]

    ingresos = dataframe["Descripción"].str.contains("Dep|Abono|Pago").map({True: "Ingreso"}).dropna()
    egresos = dataframe["Descripción"].str.contains("Ret|Inver").map({True: "Egreso"}).dropna()
    #rendimiento = dataframe["Descripción"].str.contains("Rendi").map({True: "Rendimiento"}).dropna()
    #serie = [ingresos, egresos, rendimiento]
    serie = [ingresos, egresos]
    movimiento = pd.concat(serie).sort_index()
    dataframe["movimiento"] = movimiento
    dataframe.dropna(axis=0, inplace=True)

    # query to fill this columns by tipe of transaction
    dataframe["ingresos"] = dataframe.query("movimiento == 'Ingreso'")["Monto"]
    dataframe["egresos"] = dataframe.query("movimiento == 'Egreso'")["Monto"]

    # float type
    dataframe["ingresos"] = dataframe["ingresos"].str.replace('[\$,]', '', regex=True).astype(float)
    dataframe["egresos"] = dataframe["egresos"].str.replace('[\$,]', '', regex=True).astype(float)

    dataframe.fillna(0, inplace=True)

    # delete colmns
    drop_cols = ["Descripción", "Monto", "dia", "montos"]
    dataframe.drop(drop_cols, axis=1, inplace=True)
    dataframe.reset_index(drop=True, inplace=True)


def createMathDataFrame(dataframe):
    # create empty dataframe
    new_df = pd.DataFrame(index=range(0))
    months_names = {
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
    dataframe["mes"] = dataframe["mes"].map(months_names)
    months = dataframe["mes"]
    set_months = uniqueValues(months)
    index = 0
    # columns of interes
    # MOVIMIENTO | TASA DE INTERES | INGRESOS | EGRESOS | FECHA INICIAL | FECHA FINAL | DIAS DEL MES | DIAS CONTADOS
    # iterate in order to create new columns for new rows: (formulas)
    # saldo inicial (recuperado del saldo final del mes anterior)
    # saldo promedio mensual de saldo inicial (saldo inicial * dias efectivos)
    # saldo promedio mensual de ingresos (sum(ingresos * dias efectivos))
    # saldo promedio mensual de egresos (sum(egresos * dias efectivos))
    # saldo promedio diario (sum(spmsi, spmi, spme))
    # rendimiento (spd * ti / 12)
    # subtotal (sum(si, ingresos, egresos)
    # total (subtotal + rendimiento)
    for month in set_months:
        # query to get values from this month only
        query = "mes == '{}'".format(month)
        search = dataframe.query(query)
        size = len(search.mes)      # range of sub iteration
        # counter for movimiento column
        income_counter = 0
        outcome_counter = 0
        # variables for final rows
        initial_balance = 0
        initial_balance_avg = 0
        income_sum = 0
        outcome_sum = 0
        income_sum_avg = 0
        outcome_sum_avg = 0

        for subindex in range(size):
            # stimate the index value from iteration and sub iteration
            superindex = subindex + index

            # find single values from superindex row
            # values for created rows
            date = dataframe["Fecha"][superindex]
            initial_date = dataframe["inicio de mes"][superindex]
            final_date = dataframe["fin de mes"][superindex]
            rate = dataframe["tasa de interes"][superindex]
            average = dataframe["saldo promedio"][superindex]
            income = dataframe["ingresos"][superindex]
            outcome = dataframe["egresos"][superindex]
            days = dataframe["dias efectivos"][superindex]
            all_days = dataframe["dias al mes"][superindex]
            transaction = dataframe["movimiento"][superindex]

            # print("rate: {}".format(rate))
            # print("dias efectivos: {}".format(days))

            name = transaction
            if transaction == "Ingreso":
                outcome = ""
                income_counter += 1
                income_sum += income
                income_sum_avg += average
                name = "Ingreso " + str(income_counter)
                #print("ingreso {}: {}, saldo: {}". format(income_counter, income_sum, income_sum_avg))
            elif transaction == "Egreso":
                income = ""
                outcome_counter += 1
                outcome_sum += outcome
                outcome_sum_avg += average
                name = "Egreso " + str(outcome_counter)
                #print("egreso {}: {}, saldo: {}".format(outcome_counter, outcome_sum, outcome_sum_avg))

            if subindex == 0 and superindex != 0:       # create saldo inicial row
                # create dictionary and temporal dataframe to append to returned dataframe
                d = {
                    "mes": [month],
                    "movimiento": ["Saldo inicial"],
                    "ingresos": [total],
                    "fecha inicial": [initial_date],
                    "fecha final": [final_date],
                    #"dias": [days],
                    #"dias al mes": [all_days]
                }
                temp = pd.DataFrame(data=d)
                new_df = new_df.append(temp)
                # sum saldo inicial
                initial_balance = total
                initial_balance_avg = total*days/all_days

            # create other long rows, ingresos y egresos
            # THIS IS THE INITIAL ORDER OF THE COLUMNS
            d = {
                "mes": [month],
                "movimiento": [name],
                "tasa": [""],
                "ingresos": [income],
                "egresos": [outcome],
                "fecha inicial": [date],
                "fecha final": [final_date],
                #"dias": [days],
                #"dias al mes": [all_days]
            }
            temp = pd.DataFrame(data=d)
            new_df = new_df.append(temp)
            #print("\nFINAL:",new_df)

        # sum
        index += size

        #calculate monthly final values
        monthly_average_balance = initial_balance_avg + income_sum_avg + outcome_sum_avg
        rendimiento = monthly_average_balance * rate / 12
        subtotal = initial_balance + income_sum + outcome_sum
        total = subtotal + rendimiento

        # create non repeted rows
        # saldo promedios (see above text)
        d = {
            "mes": [month],
            "movimiento": ["Saldo promedio mensual de saldo inicial"],
            "ingresos": [initial_balance_avg],

        }
        temp = pd.DataFrame(data=d)
        new_df = new_df.append(temp)

        d = {
            "mes": [month],
            "movimiento": ["Saldo promedio mensual de ingresos"],
            "ingresos": [income_sum_avg],
        }
        temp = pd.DataFrame(data=d)
        new_df = new_df.append(temp)

        d = {
            "mes": [month],
            "movimiento": ["Saldo promedio mensual de egresos"],
            "egresos": [outcome_sum_avg],
        }
        temp = pd.DataFrame(data=d)
        new_df = new_df.append(temp)

        d = {
            "mes": [month],
            "movimiento": ["Saldo promedio mensual"],
            "ingresos": [monthly_average_balance],

        }
        temp = pd.DataFrame(data=d)
        new_df = new_df.append(temp)

        d = {
            "mes": [month],
            "movimiento": ["Tasa de interes"],
            "tasa": [rate],
        }
        temp = pd.DataFrame(data=d)
        new_df = new_df.append(temp)

        d = {
            "mes": [month],
            "movimiento": ["Rendimiento"],
            "ingresos": [rendimiento],
        }
        temp = pd.DataFrame(data=d)
        new_df = new_df.append(temp)

        d = {
            "mes": [month],
            "movimiento": ["Subtotal"],
            "ingresos": [subtotal],
        }
        temp = pd.DataFrame(data=d)
        new_df = new_df.append(temp)

        d = {
            "mes": [month],
            "movimiento": ["Total"],
            "ingresos": [total],
        }
        temp = pd.DataFrame(data=d)
        new_df = new_df.append(temp)
        #print("rendimiento:{} subtotal:{} total:{}".format(rendimiento, subtotal, total))

    # format exit dataframe
    new_df.reset_index(drop=True, inplace=True)
    new_df.fillna("", inplace=True)
    return new_df


def createNewSheet(spread, dataframe, sheetname):
    spread.df_to_sheet(
        df=dataframe,
        index=False,
        sheet=sheetname,
        start="A1",
        replace=True
    )


def get_index_from_range(ranges):
    startColumnIndex, startRowIndex, endColumnIndex, endRowIndex = (0, 0, 0, 0)
    values = ranges.split(":")
    # print()
    # print("values: {}".format(values))
    last_column_index = 0
    last_row_index = 0
    list_size = len(values)
    # print(list_size)

    for index, value in enumerate(values):
        value_size = len(value)
        # print("value size: {}".format(value_size))
        if list_size == 1:  # single value [A1]
            startRowIndex = ord(value[:1]) - 65
            startRowIndex += 16 if startRowIndex < 0 else startRowIndex
            endRowIndex = startRowIndex + 1
            # print("startRowIndex: {}".format(startRowIndex))
            if value_size == 1:     #[A], [1]
                startColumnIndex = 0
                # TODO ENCONTAR EL FINAL DEL INDEX
                endColumnIndex = 1000
                # print("endRowIndex: {}".format(endRowIndex))
                # print("startColumnIndex: {}".format(startColumnIndex))
                # print("endColumnIndex: {}".format(endColumnIndex))
                # print("ENTIRE ROW {}".format(startRowIndex))
            elif value_size == 2:  #[A2]
                startColumnIndex = ord(value[:1]) - 65
                endColumnIndex = startColumnIndex + 1
                # print("endRowIndex: {}".format(endRowIndex))
                # print("startColumnIndex: {}".format(startColumnIndex))
                # print("endColumnIndex: {}".format(endColumnIndex))
                # print("column {} row {}".format(value[:1], value[1:]))

        else:  # 2 list value [A1, A2]
            if value_size <= 1:     # [A, A], [A, B100], [A10, B]
                if index == 0:  # FIRST VALUE
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
                elif index == 1: # END VALUE
                    number = ord(value[:1]) - 65
                    if number < 0:      # [#, 1] end row
                        # TODO FIND END COLUMNS
                        endColumnIndex = 1000
                        endRowIndex = number + 17
                        # print("to column {}".format(endColumnIndex))
                        # print("to row {}".format(endRowIndex))
                    else:   # [#, A] end column
                        endColumnIndex = ord(value[:1]) - 65
                        endColumnIndex += 1
                        # TODO find last row
                        endRowIndex = 1000
                        # print("to column {}".format(endColumnIndex))
                        # print("to row {}".format(endRowIndex))

            else:  # value_size = 2 list_value = 2 [A2, B2]
                column = ord(value[:1]) - 65
                row = int(value[1:]) - 1
                if index == 0:      # [A1, #]
                    startColumnIndex = column
                    last_column_index = startColumnIndex
                    startRowIndex = row
                    last_row_index = startRowIndex
                    # print("from column {}".format(value))
                    # print("startColumnIndex: {}".format(startColumnIndex))
                    # print("startRowIndex: {}".format(startRowIndex))
                elif index == 1:        # [#, B2]
                    endColumnIndex = column
                    endColumnIndex += 1
                    endRowIndex = row
                    endRowIndex += 1
                    # print("to column {}".format(value))
                    # print("endColumnIndex: {}".format(endColumnIndex))
                    # print("endRowIndex: {}".format(endRowIndex))
    return startColumnIndex, startRowIndex, endColumnIndex, endRowIndex


def uniqueValues(list):
    valores_unicos = []
    largo = len(list)
    valor_anterior = list[largo - 1]
    for dato in list:
        if valor_anterior != dato:
            valores_unicos.append(dato)
            valor_anterior = dato
    return valores_unicos

# START CODE
create_math_sheet()
