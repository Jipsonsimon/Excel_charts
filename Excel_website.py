from itertools import groupby
from unicodedata import numeric
import streamlit as st  # pip install streamlit
import pandas as pd  # pip install pandas
import plotly.express as px  # pip install plotly-express
import base64  # Standard Python Module
from io import StringIO, BytesIO  # Standard Python Module

def generate_excel_download_link(dft):
    # Credit Excel: https://discuss.streamlit.io/t/how-to-add-a-download-excel-csv-function-to-a-button/4474/5
    towrite = BytesIO()
    dft.to_excel(towrite, encoding="utf-8", index=False, header=True)  # write to BytesIO buffer
    towrite.seek(0)  # reset pointer
    b64 = base64.b64encode(towrite.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="data_download.xlsx">Download Excel File</a>'
    return st.markdown(href, unsafe_allow_html=True)

def generate_html_download_link(fig):
    # Credit Plotly: https://discuss.streamlit.io/t/download-plotly-plot-as-html/4426/2
    towrite = StringIO()
    fig.write_html(towrite, include_plotlyjs="cdn")
    towrite = BytesIO(towrite.getvalue().encode())
    b64 = base64.b64encode(towrite.read()).decode()
    href = f'<a href="data:text/html;charset=utf-8;base64, {b64}" download="plot.html">Download Plot</a>'
    return st.markdown(href, unsafe_allow_html=True)

st.set_page_config(page_title='Excel Plotter')
st.title('Excel Plotter 📈')
st.subheader('Feed me with your Excel file')

uploaded_file = st.file_uploader('Choose a XLSX file', type='xlsx')


if uploaded_file:
    st.markdown('---')
    dft = pd.read_excel(uploaded_file, engine='openpyxl')
    st.dataframe(dft)

    dfs= pd.read_excel(uploaded_file,  index_col=0)
    
    numeric_columns = list(dfs.select_dtypes(['float','int']).columns)
    chart_select = st.sidebar.selectbox(
        label="Select the chart type",
        options=['Barchart','Piechart']
    )
    if chart_select == 'Barchart' :
        st.sidebar.subheader("Barchart Settings")
        
        try:
            x_values = st.sidebar.selectbox('X axis', options=numeric_columns)
            y_values = st.sidebar.selectbox('Y axis', options=numeric_columns)
        except Exception as e:
            print(e)
        output_columns = []
        df_grouped = dft.groupby(by=[x_values,y_values], as_index=False)[output_columns].sum()
        st.dataframe(df_grouped)

        # -- PLOT DATAFRAME
        fig = px.bar(
            df_grouped,
            x=x_values,
            y=y_values,
            color=x_values,
            color_continuous_scale=['red', 'yellow', 'green'],
            template='plotly_white',
            title=f'<b>X axis & Y axis by {x_values,y_values}</b>'
        )
        st.plotly_chart(fig)
    
    if chart_select =='Piechart':
        st.sidebar.subheader("Piechart Settings")
        try:
            x_values = st.sidebar.selectbox('X axis', options=numeric_columns)
            y_values = st.sidebar.selectbox('Y axis', options=numeric_columns)
        except Exception as e:
            print(e)
        output_columns = []
        df_grouped = dft.groupby(by=[x_values,y_values], as_index=False)[output_columns].sum()
        st.dataframe(df_grouped) 
        fig = px.pie(df_grouped,
            values=x_values,
            names=y_values)
        fig.update_traces(
            textposition = 'inside',
            textinfo = 'percent+label'
        )
        st.plotly_chart(fig)


    # -- DOWNLOAD SECTION
    st.subheader('Downloads:')
    generate_excel_download_link(df_grouped)
    generate_html_download_link(fig)
