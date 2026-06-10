from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

def generar_cruce(archivo_ingresos,archivo_egresos,archivo_salida):
    

    wb_ingresos = load_workbook(archivo_ingresos,data_only=False)
    wb_egresos = load_workbook(archivo_egresos,data_only=False)
    ws_ingresos = wb_ingresos.active
    ws_egresos = wb_egresos.active

    wb_nuevo = Workbook()
    ws1 = wb_nuevo.active
    ws1.title = "INGRESOS"

    for fila in ws_ingresos.iter_rows():
        for celda in fila:
            ws1[celda.coordinate] = celda.value

    ws2 = wb_nuevo.create_sheet("EGRESOS")

    for fila in ws_egresos.iter_rows():
        for celda in fila:
            ws2[celda.coordinate] = celda.value

    ws3 = wb_nuevo.create_sheet("CRUCE INFORMACION")

    encabezados = ["FUENTE DE RECURSO","NOMBRE DE FUENTE DE RECURSO","PRESUPUESTO DEFINITIVO","PRESUPUESTO DEFINITIVO EJECUTADO","DIFERENCIA","Recaudo","Registros","Saldos no Comprometidos","Constitución","Reservas","Pago","Cuentas x Pagar","Recurso en Bancos"]

    gris = "7F7F7F"
    azul_oscuro = "385D8A"
    morado = "7030A0"
    azul = "0070C0"
    naranja = "ED7D31"
    verde = "92D050"

    colores = [gris,gris,azul_oscuro,azul_oscuro,azul_oscuro,morado,morado,morado,azul,azul,naranja,naranja,verde]

    for col, (titulo, color) in enumerate(zip(encabezados, colores), start=1):
        celda = ws3.cell(row=1, column=col)
        celda.value = titulo
        celda.fill = PatternFill("solid", fgColor=color)
        celda.font = Font(color="FFFFFF", bold=True)
        celda.alignment = Alignment(horizontal="center", vertical="center")


    ws3.row_dimensions[1].height = 30

    encabezados_egresos = {}

    for col in range(1, ws2.max_column + 1):
        valor = ws2.cell(row=2, column=col).value
        if valor is not None:
            encabezados_egresos[str(valor).strip().upper()] = col

    col_fuente = encabezados_egresos["FUENTE DE RECURSO"]
    col_nombre = encabezados_egresos["NOMBRE DE FUENTE DE RECURSO"]

    fuentes_vistas = set()
    fila_destino = 2

    for fila in range(3, ws2.max_row + 1):
        fuente = ws2.cell(row=fila, column=col_fuente).value
        nombre = ws2.cell(row=fila, column=col_nombre).value
        if fuente is None or str(fuente).strip() == "":
            continue
        try:
            if float(fuente) == 0:
                continue
        except:
            pass

        clave = str(fuente).strip()
        if clave in fuentes_vistas:
            continue

        fuentes_vistas.add(clave)
        ws3.cell(row=fila_destino, column=1, value=fuente)
        ws3.cell(row=fila_destino, column=2, value=nombre)
        fila_destino += 1

    ultima_fila = ws3.max_row

    for fila in range(2, ultima_fila + 1):
        ws3[f"C{fila}"] = f'=SUMIF(INGRESOS!$R:$R,A{fila},INGRESOS!$L:$L)'
        ws3[f"D{fila}"] = f'=SUMIF(EGRESOS!$P:$P,A{fila},EGRESOS!$Y:$Y)'
        ws3[f"E{fila}"] = f'=C{fila}-D{fila}'
        ws3[f"F{fila}"] = f'=SUMIF(INGRESOS!$R:$R,A{fila},INGRESOS!$N:$N)'
        ws3[f"G{fila}"] = f'=SUMIF(EGRESOS!$P:$P,A{fila},EGRESOS!$AA:$AA)'
        ws3[f"H{fila}"] = f'=F{fila}-G{fila}'
        ws3[f"I{fila}"] = f'=SUMIF(EGRESOS!$P:$P,A{fila},EGRESOS!$AB:$AB)'
        ws3[f"J{fila}"] = f'=G{fila}-I{fila}'
        ws3[f"K{fila}"] = f'=SUMIF(EGRESOS!$P:$P,A{fila},EGRESOS!$AC:$AC)'
        ws3[f"L{fila}"] = f'=I{fila}-K{fila}'
        ws3[f"M{fila}"] = f'=H{fila}+J{fila}+L{fila}'

    ws4 = wb_nuevo.create_sheet("INFORME RAPIDO")

    ws4["A1"] = "CONCEPTO"
    ws4["B1"] = "VALOR"
    ws4["A2"] = "Recaudo hoja INGRESOS"
    ws4["B2"] = "=INGRESOS!N2"
    ws4["A3"] = "Suma recaudo hoja CRUCE INFORMACION"
    ws4["B3"] = "=SUM('CRUCE INFORMACION'!F:F)"
    ws4["A4"] = "Diferencia"
    ws4["B4"] = "=B2-B3"
    ws4["A5"] = "Resultado"
    ws4["B5"] = (
        '=IF(ABS(B4)<0.01,'
        '"Los datos coinciden",'
        '"Los datos no coinciden, faltan productos")'
    )


    ws4["A8"] = "VALIDACION PRESUPUESTO DEFINITIVO"
    ws4["A9"] = "Presupuesto Definitivo INGRESOS"
    ws4["B8"] = "VALOR"
    ws4["B9"] = "=INGRESOS!L2"
    ws4["A10"] = "Suma Presupuesto Definitivo CRUCE"
    ws4["B10"] = "=SUM('CRUCE INFORMACION'!C:C)"
    ws4["A11"] = "Diferencia"
    ws4["B11"] = "=B9-B10"
    ws4["A12"] = "Resultado"
    ws4["B12"] = (
        '=IF(ABS(B11)<0.01,'
        '"Los datos coinciden",'
        '"Los datos no coinciden, falta un producto")'
    )

    for celda in ["B2", "B3", "B4", "B9" , "B10" , "B11" ]:
        ws4[celda].number_format = '$#,##0.00'

    thin = Side(style="thin", color="000000")

    for fila in range(1, 6):
        for col in range(1, 3):
            ws4.cell(fila, col).border = Border(
                left=thin,
                right=thin,
                top=thin,
                bottom=thin
            )

    ws4.column_dimensions["A"].width = 40
    ws4.column_dimensions["B"].width = 25

    for fila in range(1, ws3.max_row + 1):
        for col in range(1, ws3.max_column + 1):
            celda = ws3.cell(row=fila, column=col)
            celda.border = Border(left=thin,right=thin,top=thin,bottom=thin)

    for col in range(3, 14):  # C=3, M=13
        letra_col = get_column_letter(col)
        for fila in range(2, ws3.max_row + 1):
            ws3[f"{letra_col}{fila}"].number_format = '$#,##0.00'
    for col in ws3.columns:
        max_length = 0
        columna = get_column_letter(col[0].column)
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ws3.column_dimensions[columna].width = min(max_length + 3, 40)

    ws3.conditional_formatting.add(
        f"C2:M{ws3.max_row}",
        CellIsRule(
            operator='lessThan',
            formula=['0'],
            fill=PatternFill(start_color='FFFF00',
                            end_color='FFFF00',
                            fill_type='solid')
        )
    )
    for fila in range(8, 13):
        for col in range(1, 3):
            ws4.cell(fila, col).border = Border(left=thin,right=thin,top=thin,bottom=thin)
    wb_nuevo.save(archivo_salida)

    return archivo_salida