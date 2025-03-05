import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from statsmodels.tsa.arima.model import ARIMA
from prophet import Prophet
import io

def load_data(file):
    data = pd.read_csv(file, parse_dates=['Date'], index_col='Date')
    return data

def forecast_sales(data, periods, model_type='ARIMA'):
    data = data.resample('M').sum()
    if model_type == 'ARIMA':
        model = ARIMA(data, order=(5,1,0))
        model_fit = model.fit()
        forecast = model_fit.forecast(steps=periods)
    else:
        df = data.reset_index().rename(columns={'Date': 'ds', 'Sales': 'y'})
        model = Prophet()
        model.fit(df)
        future = model.make_future_dataframe(periods=periods, freq='M')
        forecast = model.predict(future)[['ds', 'yhat']].tail(periods)
        forecast.set_index('ds', inplace=True)
    return forecast

def main():
    st.title('E-commerce Sales Forecasting App')
    st.write("Upload your sales data to forecast future sales trends.")
    
    uploaded_file = st.file_uploader("Upload CSV file", type=["csv"])
    if uploaded_file is not None:
        data = load_data(uploaded_file)
        st.write("### Uploaded Data Preview:")
        st.write(data.head())
        
        periods = st.selectbox("Select Forecast Period:", [6, 12])
        model_type = st.radio("Select Model:", ['ARIMA', 'Prophet'])
        
        if st.button("Generate Forecast"):
            forecast = forecast_sales(data, periods, model_type)
            st.write("### Forecasted Sales")
            st.line_chart(forecast)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                forecast.to_excel(writer, sheet_name='Forecast')
            output.seek(0)
            
            st.download_button(label="Download Forecast Report",
                               data=output,
                               file_name="forecast_report.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    main()
