import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.figure_factory as ff
import os
import warnings
from io import BytesIO
warnings.filterwarnings('ignore')

# ------------- TERMINAL COMMAND ------------
# py -m streamlit run app.py

# -------------- SETTINGS --------------------
def download_excel_file(df, fileName):
    # Create a BytesIO buffer to store the Excel file
    buffer = BytesIO()
    # Save the Excel data to the buffer
    writer = pd.ExcelWriter(buffer, engine='xlsxwriter')
    df.to_excel(writer, index=False)
    # Close the writer to ensure data is written to the buffer
    writer.close()

    # Create a download button with the buffer data
    st.download_button(label='Download Data', 
    data=buffer, # Get the binary data from the buffer
    file_name=fileName, 
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    help='Click here to download the data as a XLSX file')

excelFile = 'Superstore.xlsx'

st.set_page_config(page_title='Superstore', page_icon=':convenience_store:', layout='wide')

st.title(':bar_chart: SuperStore Sales Dashboard')
st.markdown('<style>div.block-container{padding-top:2rem;}</style>', unsafe_allow_html=True)

# ----------------- PREPARE THE DATASET ---------------------
dir_path = os.path.dirname(os.path.realpath(__name__))
os.chdir(dir_path) # Use chdir() to change the directory
df = pd.read_excel(excelFile)

# Round the "Sales" and "Profit" columns to two decimal places
df['Sales'] = df['Sales'].round(2)
df['Profit'] = df['Profit'].round(2)

col1, col2 = st.columns(2)
df['Order Date'] = pd.to_datetime(df['Order Date'])

# # ----------------- MIN AND MAX DATE ----------------------
startDate = pd.to_datetime(df['Order Date']).min()
endDate = pd.to_datetime(df['Order Date']).max()

with col1:
    date1 = pd.to_datetime(st.date_input('Start Date', startDate))

with col2:
    date2 = pd.to_datetime(st.date_input('End Date', endDate))

df = df[(df['Order Date'] >= date1) & (df['Order Date'] <= date2)]

# -------------------- SIDEBAR -----------------------------
st.sidebar.header("Choose your filter: ")

# Region filter
region = st.sidebar.multiselect('Pick your Region', df['Region'].unique())
if not region:
    df2 = df.copy()
else:
    df2 = df[df['Region'].isin(region)] # The isin() method checks if the Dataframe contains the specified value(s).

# State filter => it shows only the States related to the selected Regions
state = st.sidebar.multiselect('Pick the State', df2['State'].unique())
if not state:
    df3 = df2.copy()
else:
    df3 = df2[df2['State'].isin(state)]

# City filter
city = st.sidebar.multiselect('Pick the City', df3['City'].unique())

# Filter the data based on Region, State and City
if not region and not state and not city:
    filtered_df = df
elif not state and not city:
    filtered_df = df[df["Region"].isin(region)]
elif not region and not city:
    filtered_df = df[df["State"].isin(state)]
elif state and city:
    filtered_df = df3[df["State"].isin(state) & df3["City"].isin(city)]
elif region and city:
    filtered_df = df3[df["Region"].isin(region) & df3["City"].isin(city)]
elif region and state:
    filtered_df = df3[df["Region"].isin(region) & df3["State"].isin(state)]
elif city:
    filtered_df = df3[df3["City"].isin(city)]
else:
    filtered_df = df3[df3["Region"].isin(region) & df3["State"].isin(state) & df3["City"].isin(city)]

category_df = filtered_df.groupby(by=["Category"], as_index = False)["Sales"].sum()

# ------- CATEGORY WISE SALES [BAR CHART] -----------
with col1: 
    st.subheader('Category wise Sales') # vendite in termini di categoria
    fig = px.bar(category_df, x='Category', y='Sales', text=['${:,.2f}'.format(x) for x in category_df['Sales']], template='seaborn') # for each value in the 'Sales' column => ${:,.2f} is a string format specifier: 
    # - ${} => literal dollar sign that will appear at the beginning of each formatted string.
    # - :, => used to format numbers with commas as thousands separators. For example, it will format 1000 as "1,000".
    # - .2f => the number should be formatted as a floating-point number with exactly two decimal places.
    st.plotly_chart(fig, use_container_width=True, height=200)

# ------- REGION WISE SALES [PIE CHART] -----------
with col2:
    st.subheader('Region wise Sales')
    fig = px.pie(filtered_df, values='Sales', names='Region', hole=0.25)
    fig.update_traces(text=filtered_df['Region'], textposition='outside')
    st.plotly_chart(fig, use_container_width=True)

# --------------- EXPANDER WIDGETS WITH CHARTS DATA -----------------
cl1, cl2 = st.columns(2)
with cl1:
    with st.expander('View Data by Category', expanded=True):
        st.write(category_df.style.background_gradient(cmap='Blues'))
        download_excel_file(category_df, 'Category.xlsx')
with cl2:
    with st.expander('View Regional Data', expanded=True):
        region = filtered_df.groupby(by=["Region"], as_index = False)["Sales"].sum()

        st.write(region.style.background_gradient(cmap='Purples'))
        download_excel_file(region, 'Region.xlsx')

# ------- TIME SERIES ANALYSIS [LINE CHART] -----------       
filtered_df['month_year'] = filtered_df['Order Date'].dt.to_period('M')
st.subheader('Time Series Analysis')

# create the line chart
linechart = pd.DataFrame(filtered_df.groupby(filtered_df['month_year'].dt.strftime('%Y : %b'))['Sales'].sum()).reset_index()
fig2 = px.line(linechart, x='month_year', y='Sales', labels={'Sales':'Amount'}, height=500, width=1000, template='gridon')
st.plotly_chart(fig2, use_container_width=True)

with st.expander('View Time Series Data'):
        st.write(linechart.T.style.background_gradient(cmap='Blues'))
        download_excel_file(linechart, 'TimeSeries.xlsx')

# ---------- HIERARCHICAL VIEW OF SALES [TREEMAP CHART] ----------------
# Create a treemap based on Region, Category, sub-Category
st.subheader('Hierarchical view of Sales')
fig3 = px.treemap(filtered_df, 
                  path=['Region','Category','Sub-Category'], 
                  values='Sales', 
                  hover_data=['Sales'], 
                  color='Sub-Category')

fig3.update_layout(width=800,height=650)
st.plotly_chart(fig3, use_container_width=True)

# ------- SEGMENT WISE SALES + CATEGORY WISE SALES [PIE CHARTS] -----------
chart1, chart2 = st.columns(2)

with chart1:
    st.subheader('Segment wise Sales')
    fig = px.pie(filtered_df, values='Sales', names='Segment', template='plotly_dark')
    fig.update_traces(text=filtered_df['Segment'], textposition='inside')
    st.plotly_chart(fig, use_container_width=True)

with chart2:
    st.subheader('Category wise Sales')
    fig = px.pie(filtered_df, values='Sales', names='Category', template='gridon')
    fig.update_traces(text=filtered_df['Category'], textposition='inside')
    st.plotly_chart(fig, use_container_width=True)

st.subheader(':point_right: Month wise Sub-Category Sales Summary')
with st.expander('Summary Tables'):
    df_sample = df[0:5][['Region','State','City','Category','Sales','Profit', 'Quantity']]
    colorscale = [[0, '#0C134F'],[.5, '#AED2FF'],[1, '#E4F1FF']]
    fig = ff.create_table(df_sample, colorscale=colorscale)
    st.plotly_chart(fig, use_container_width=True)

    st.markdown('Month wise Sub-Category Table')
    filtered_df['month'] = filtered_df['Order Date'].dt.month_name()
    sub_category_Year = pd.pivot_table(data = filtered_df, values='Sales', index=['Sub-Category'], columns='month')
    st.write(sub_category_Year.style.background_gradient(cmap='Blues'))

# --------- HIDE STREAMLIT STYLE -----------
hide_st_style = """
                <style>
                    MainMenu {visibility: hidden;}
                    footer {visibility: hidden;}
                    header {visibility: hidden;}
                </style>
                """
st.markdown(hide_st_style, unsafe_allow_html=True)