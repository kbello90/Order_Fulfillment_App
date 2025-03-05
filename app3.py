import streamlit as st
import pandas as pd
import plotly.express as px
import io

def load_data(file):
    df = pd.read_csv(file)
    return df

def main():
    st.title("SmartWay Logistics Management")
    st.write("Track shipments, analyze costs, and optimize logistics operations.")
    
    uploaded_file = st.file_uploader("Upload CSV file (Shipments & Orders)", type=["csv"])
    if uploaded_file is not None:
        df = load_data(uploaded_file)
        st.write("### Uploaded Data Preview:")
        st.write(df.head())
        
        # Shipment status analysis
        if 'Status' in df.columns:
            status_counts = df['Status'].value_counts()
            fig_status = px.pie(status_counts, names=status_counts.index, values=status_counts.values, title="Shipment Status Distribution")
            st.plotly_chart(fig_status)
        
        # Cost analysis
        if 'Cost' in df.columns and 'Distance' in df.columns:
            df['Cost per Mile'] = df['Cost'] / df['Distance']
            st.write("### Cost Analysis")
            st.bar_chart(df.groupby('Route')['Cost per Mile'].mean())
        
        # Export updated data
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Updated_Data', index=False)
        output.seek(0)
        
        st.download_button(label="Download Updated Data",
                           data=output,
                           file_name="smartway_data.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    main()
