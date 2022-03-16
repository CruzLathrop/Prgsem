import pandas as pd
import dateutil

def get_data(fecha_ini, fecha_fin):
    df_out=pd.DataFrame()
    for fecha in pd.date_range(start=fecha_ini, end=fecha_fin):
        año = fecha.strftime("%Y")
        mes = fecha.strftime("%m")
        dia = fecha.strftime("%d")
        aaaammdd = año + mes + dia
        aammdd = año[2:4] + mes + dia
        ruta = "Z:/2022/CO2/Reportes Monitoreo/data CV/" + año + "/" + aaaammdd
        archivo = 'PO' + aammdd + '.xlsx'
        archivo2 = 'PO' + aammdd + '_ori.xlsx'
        hoja = 'politicas diarias'
        ##

        try:
            try:
                df = pd.read_excel(ruta + '/' + archivo2, sheet_name=hoja, header=4)
            except:
                df = pd.read_excel(ruta + '/' + archivo, sheet_name=hoja, header=4)
        except:
            continue

        df['fecha_p'] = fecha.date()
        df = df.dropna(subset=['CENTRALES'])
        df = df.loc[df["CENTRALES"] != 'CENTRALES' ]
        df = df.drop(['Unnamed: 3', 'Nº.1','Unnamed: 7', 'Nº.2'],1)
        df2 = df.reset_index(drop=True)

        try:
            num_cent = df2[df2['Nº'] == 1].index.values[1]
        except:
            pass

        for i in range(1,8):
            df2.loc[(df2.index <= (num_cent)*i) & (df2.index >= (num_cent)*(i-1)), 'fecha_a'] = fecha.date() + dateutil.relativedelta.relativedelta(days=(i-1))

        #pd.set_option('display.max_columns', None)
        #pd.set_option('display.max_rows', None)
        #print(df_out)
        df_out = df_out.append(df2, ignore_index=True)
        print(fecha.date())
    return df_out

# Press the green button in the gutter to run the script.
#if __name__ == '__main__':
data = get_data('2014-10-03', '2015-03-14')
data.to_csv('output_parte2.csv',index=False)
