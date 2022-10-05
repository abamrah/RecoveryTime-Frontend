import pandas as pd
import numpy as np
import streamlit as st
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb
import os

def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    format1 = workbook.add_format({'num_format': '0.00'}) 
    worksheet.set_column('A:A', None, format1)  
    writer.save()
    processed_data = output.getvalue()
    return processed_data


def recoverytime(df):
       df['Actual Start Date'] = df['Actual start'].dt.strftime('%Y-%m-%d')
       df['Actual End Date'] = df['Act.finish date'].dt.strftime('%Y-%m-%d')

       df['Actual_Start_Date_time'] = pd.to_datetime(df['Actual Start Date'] + ' ' + df['Act. start time'])
       df['Actual_Finish_Date_time'] = pd.to_datetime(df['Actual End Date'] + ' ' + df['Actual finish'])

       df.drop(['Actual Start Date', 'Actual End Date'], axis = 1, inplace = True)

       dfnew = df[df.groupby('Order').Actual_Finish_Date_time.transform('max') == df['Actual_Finish_Date_time']]
       dfnew2 = df[df.groupby('Order').Actual_Start_Date_time.transform('min') == df['Actual_Start_Date_time']]

       dfnew.drop(['Oper./Act.', 'Order Type', 'Actual start', 'Act. start time',
              'Act.finish date', 'Actual finish', 'Control key', 'Equipment',
              'Description', 'Work Center', 'Plant', 'Actual work',
              'Actual_Start_Date_time'], axis = 1, inplace = True)
       dfnew.drop_duplicates(inplace = True)

       dfnew2.drop(['Oper./Act.', 'Order Type', 'Actual start', 'Act. start time',
              'Act.finish date', 'Actual finish', 'Control key', 'Equipment',
              'Description', 'Work Center', 'Plant', 'Actual work',
              'Actual_Finish_Date_time'], axis = 1, inplace = True)
       dfnew2.drop_duplicates(inplace = True)

       df.drop(['Actual_Start_Date_time', 'Actual_Finish_Date_time'], axis = 1, inplace = True)

       dff = pd.merge(df, dfnew2, on ='Order', how ='left')
       dff = pd.merge(dff, dfnew, on ='Order', how ='left')

       dff['Recovery_Time_from_Maintenace'] = np.subtract(dff['Actual_Finish_Date_time'], dff['Actual_Start_Date_time'])
       dff['Recovery_Time_from_Maintenace'] = dff['Recovery_Time_from_Maintenace'].dt.total_seconds()/3600

       return dff



### Streamlit and GUI 

header = st.container()
setup = st.container()
outputmodel = st.container()

with header:
    st.title('Apotex Recovery Time Extractor')

with setup:
    st.header('Please Upload the csv file. It only accepts csv at the moment')
    data_file = st.file_uploader("Upload csv or excel")
    if data_file is not None:
       st.write(type(data_file))
       try: 
              df = pd.read_excel(data_file)
       
       except:
              df = pd.read_csv(data_file)

       st.write(df.head())
    else:
       st.warning('you need to upload a csv or excel file.')

    df['Actual start'] = pd.to_datetime(df['Actual start'])
    df['Act.finish date'] = pd.to_datetime(df['Act.finish date'])


with outputmodel:
         st.header('Begin Updating')
         dfff = recoverytime(df)     
         st.header('Finished updating')
         st.dataframe(dfff)

         df_xlsx = to_excel(dfff)

         st.header('Please Input the New filename')
         file_name_df = st.text_input('Enter file name', 'enter here')

         st.header('File will be saved as ' + file_name_df + '.xlsx')
        
         st.download_button(label='📥 Download Current Processed file',
                                data=df_xlsx,
                                file_name= file_name_df + '.xlsx')





