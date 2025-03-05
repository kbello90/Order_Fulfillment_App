import streamlit as st
import pandas as pd
import plotly.express as px
import io
from fpdf import FPDF

def load_data(file):
    df = pd.read_excel(file) if file.name.endswith('.xlsx') else pd.read_csv(file)
    df['TKM'] = (df['Actual Weight (kgs)'] * df['Transport Distance (km)']) / 1000
    df['TON-MILE'] = df['TKM'] * 0.6213
    df = df.round({'TKM': 2, 'TON-MILE': 2})
    return df

def generate_pdf(pivot_table):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, txt="Smartway Calculator Report", ln=True, align='C')
    pdf.ln(10)
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(40, 10, 'Carrier')
    pdf.cell(40, 10, 'TON-MILE')
    pdf.cell(40, 10, 'TKM')
    pdf.ln(10)
    pdf.set_font("Arial", size=12)
    
    for _, row in pivot_table.iterrows():
        pdf.cell(40, 10, row['Carrier'])
        pdf.cell(40, 10, str(row['TON-MILE']))
        pdf.cell(40, 10, str(row['TKM']))
        pdf.ln(10)
    
    pdf_output = io.BytesIO()
    pdf.output(pdf_output, dest='S')
    pdf_output.seek(0)
    return pdf_output.getvalue()

def main():
    st.title("SmartWay Logistics Management")
    st.write("Upload shipment data to analyze logistics efficiency.")
    
    uploaded_file = st.file_uploader("Upload CSV or Excel file", type=["csv", "xlsx"])
    if uploaded_file is not None:
        df = load_data(uploaded_file)
        st.write("### Data Preview:")
        st.dataframe(df.head())
        
        carrier_list = ['All'] + df['Carrier'].unique().tolist()
        selected_carrier = st.selectbox("Filter by Carrier:", carrier_list)
        filtered_df = df if selected_carrier == 'All' else df[df['Carrier'] == selected_carrier]
        
        pivot_table = filtered_df.groupby('Carrier').agg({'TON-MILE': 'sum', 'TKM': 'sum'}).reset_index()
        st.write("### Aggregated Data")
        st.dataframe(pivot_table)
        
        fig = px.bar(pivot_table, x='Carrier', y=['TON-MILE', 'TKM'], title='TON-MILE & TKM by Carrier')
        st.plotly_chart(fig)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            pivot_table.to_excel(writer, sheet_name='Aggregated Data', index=False)
        output.seek(0)
        
        st.download_button(label="Download Excel Report", data=output, file_name="smartway_report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        pdf_output = generate_pdf(pivot_table)
        st.download_button(label="Download PDF Report", data=pdf_output, file_name="smartway_report.pdf", mime="application/pdf")

if __name__ == "__main__":
    main()

