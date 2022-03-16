import pandas as pd
import dateutil
import openpyxl

def get_data(fecha_ini, fecha_fin):
    df_out=pd.DataFrame()
    for fecha in pd.date_range(start=fecha_ini, end=fecha_fin):
        año = fecha.strftime("%Y")
        mes = fecha.strftime("%m")
        dia = fecha.strftime("%d")
        aaaammdd = año + mes + dia
        aammdd = año[2:4] + mes + dia
        ruta = "Z:/2017/CEN/Prgdia/" + año + "/" + aaaammdd
        archivo = 'PS' + aammdd + '.xlsx'
        hoja = 'Generación horaria'
        try:
            wb = openpyxl.load_workbook(ruta + '/' + archivo, read_only=True)
        except:
            continue

        ws = wb.active
        for row in ws.iter_rows(min_row=1, max_col=1, max_row=20):
            for cell in row:
                if cell.value == "Central Nombre":
                    cabecera = cell.row - 1

        df = pd.read_excel(ruta + '/' + archivo, sheet_name=hoja, header=cabecera)
        df.drop(df.tail(2).index, inplace=True)
        df = df.iloc[: , :169]
        df2 = df.melt(id_vars=['Central Nombre'], var_name='hora', value_name='gen')
        # se corrigen nombre hora
        df2['dia'] = 1
        for i in range(1, 25):
            for j in range(1, 7):
                df2.loc[df2['hora'] == str(i) + '.' + str(j), 'dia'] = j + 1
                df2.loc[df2['hora'] == str(i) + '.' + str(j), 'hora'] = i

        df2['fecha_p'] = fecha.date()
        df2['fecha_a'] = df2.apply(
            lambda row: row.fecha_p + dateutil.relativedelta.relativedelta(days=row.dia - 1),
            axis=1)
        df_out = df_out.append(df2, ignore_index=True)
        print(fecha.date())

    return df_out

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    data = get_data('2017-01-06', '2017-12-31')
    data.to_csv('output_2017.csv',index=False)
