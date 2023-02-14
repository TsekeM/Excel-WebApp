import pandas as pd
import streamlit as st
import plotly.express as px
from PIL import Image


st.set_page_config(page_title='Muluma E-Learning dashboard')
st.header('Data visuals E_Learning dashboard')
st.subheader('Making sense of the data')

# --- LOAD DATAFRAME
excel_file = 'Survey_Results.xlsx'
sheet_name = 'DATA'

df = pd.read_excel(excel_file,
                   sheet_name=sheet_name,
                   usecols='B:D',
                   header=3)

df_participants = pd.read_excel(excel_file,
                                sheet_name=sheet_name,
                                usecols='F:G',
                                header=3)
df_participants.dropna(inplace=True)

# --- STREAMLIT SELECTION
department = df['Department'].unique().tolist()
ages = df['Age'].unique().tolist()

age_selection = st.slider('Age:',
                          min_value=min(ages),
                          max_value=max(ages),
                          value=(min(ages), max(ages)))

department_selection = st.multiselect('Department:',
                                      department,
                                      default=department)

# --- FILTER DATAFRAME BASED ON SELECTION
mask = (df['Age'].between(*age_selection)
        ) & (df['Department'].isin(department_selection))
number_of_result = df[mask].shape[0]
st.markdown(f'*Available Results: {number_of_result}*')

# --- GROUP DATAFRAME AFTER SELECTION
df_grouped = df[mask].groupby(by=['Rating']).count()[['Age']]
df_grouped = df_grouped.rename(columns={'Age': 'Votes'})
df_grouped = df_grouped.reset_index()

# --- PLOT BAR CHART
bar_chart = px.bar(df_grouped,
                   x='Rating',
                   y='Votes',
                   text='Votes',
                   color_discrete_sequence=['#F63366']*len(df_grouped),
                   template='plotly_white')
st.plotly_chart(bar_chart)

# --- DISPLAY IMAGE & DATAFRAME
col1, col2 = st.columns(2)
image = Image.open('images/survey.jpg')
col1.image(image,
           caption='Designed by slidesgo / Freepik',
           use_column_width=True)
col2.dataframe(df[mask])

# --- PLOT PIE CHART
pie_chart = px.pie(df_participants,
                   title='Total No. of Participants',
                   values='Participants',
                   names='Departments')

st.plotly_chart(pie_chart)

st.sidebar.image(
    'https://attachments.office.net/owa/Tseke%40muluma.co.za/service.svc/s/GetAttachmentThumbnail?id=AAMkAGY3OTJlNWNhLTllMjEtNGY3MS1iNjZkLTMxOWQwODQ2MzkyOQBGAAAAAACC0cOiyYrGTKmRC2zvU5oQBwBWrFuQWi4VQJTBYB9rh5vZAAAAAAEMAABWrFuQWi4VQJTBYB9rh5vZAAAvWiAiAAABEgAQAB0Fl2m2vc1IgLSvXTYSIQ4%3D&thumbnailType=2&token=eyJhbGciOiJSUzI1NiIsImtpZCI6IkQ4OThGN0RDMjk2ODQ1MDk1RUUwREZGQ0MzODBBOTM5NjUwNDNFNjQiLCJ0eXAiOiJKV1QiLCJ4NXQiOiIySmozM0Nsb1JRbGU0Tl84dzRDcE9XVUVQbVEifQ.eyJvcmlnaW4iOiJodHRwczovL291dGxvb2sub2ZmaWNlLmNvbSIsInVjIjoiOWFmZmE3YjBjZDVhNDQ5MjgxN2Q5ZGQxYmVkYzRjYTkiLCJzaWduaW5fc3RhdGUiOiJbXCJrbXNpXCJdIiwidmVyIjoiRXhjaGFuZ2UuQ2FsbGJhY2suVjEiLCJhcHBjdHhzZW5kZXIiOiJPd2FEb3dubG9hZEBhNTAxZmEyYy0wMDJmLTQ5MTItODUzMy1iYzg4OTVlYzEzYzkiLCJpc3NyaW5nIjoiV1ciLCJhcHBjdHgiOiJ7XCJtc2V4Y2hwcm90XCI6XCJvd2FcIixcInB1aWRcIjpcIjExNTM4MDExMjMxOTM5ODI2NzFcIixcInNjb3BlXCI6XCJPd2FEb3dubG9hZFwiLFwib2lkXCI6XCI5ODk3OTI0ZS0zNjkzLTRkYjUtYmEzNy1mYjU4MDVlZDViYjVcIixcInByaW1hcnlzaWRcIjpcIlMtMS01LTIxLTI5OTA4ODQ2NDEtMTM1MTQzMTUwNy0zMzUzMzg1ODg3LTE3NjI5MzI1XCJ9IiwibmJmIjoxNjY4NTk3MTM4LCJleHAiOjE2Njg1OTc3MzgsImlzcyI6IjAwMDAwMDAyLTAwMDAtMGZmMS1jZTAwLTAwMDAwMDAwMDAwMEBhNTAxZmEyYy0wMDJmLTQ5MTItODUzMy1iYzg4OTVlYzEzYzkiLCJhdWQiOiIwMDAwMDAwMi0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvYXR0YWNobWVudHMub2ZmaWNlLm5ldEBhNTAxZmEyYy0wMDJmLTQ5MTItODUzMy1iYzg4OTVlYzEzYzkiLCJoYXBwIjoib3dhIn0.LZpqdNTAhLoIBgtyp4zX9Ayq9IaOuZGqVPOPsjujfC8s5qiqOkU6jZlVcdWQbGyxhWiVrexTvMxYdrSEI3M0T-dgyQKnN5EMDsTdtLw__h-_SWIpsL2c3GPFzpM6aZ0FjflAr7y4juyyY-PyQZCPu4zYiBBe9F89RMqiIe2g3oBX2YmVSLm6KeqNa6NaSs97bKp1Ob_rGhZ8uHFmNcHHL_YVUoc8OUiZQeJwiDNsXjLxC9B7iBnjlIG3wK4Uj5B3LZiFV2yxa7pkCRa4YiS3OYJYo98OpmgpYu9jO7SgmMam8eaXDsq5b7cecLQBzspehukbIUAPbIfNvpCTCptefA&X-OWA-CANARY=CTEZMSwS6UeCvR69xOC1wCDmqqTDx9oYJrnDQY5ftaTzTyE2YzsT0p9EwYUk00I62Rf0GkiJ23U.&owa=outlook.office.com&scriptVer=20221104009.07&animation=true', width=200)
st.sidebar.header('Muluma Dashboard `version 1`')
#st.sidebar.header("Please Apply Filter Here:")
st.sidebar.markdown('''
---
Made with ❤️ by Tseke Maila''')

# ---- HIDE STREAMLIT STYLE ----
hide_st_style = """
            <style>
            # MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)
