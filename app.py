import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# Load Data
file_path = "order_fulfillment_dashboard_updated.xlsx"
xls = pd.ExcelFile(file_path)
fact_sales_df = pd.read_excel(xls, sheet_name="FactSales")

# Ensure Date column is in datetime format
fact_sales_df["Date"] = pd.to_datetime(fact_sales_df["Date"], errors='coerce')

# Drop rows with missing dates to avoid KeyError
fact_sales_df = fact_sales_df.dropna(subset=["Date"])

# Define color palette
colors = ["#00A6FB", "#0582CA", "#006494", "#003554", "#051923"]

# Extract Month from Date
fact_sales_df["Month"] = fact_sales_df["Date"].dt.to_period("M").astype(str)

# Sidebar Navigation
st.sidebar.title("📊 Order Fulfillment Dashboard by Karen Bello")
st.sidebar.markdown("### Navigation")
page = st.sidebar.radio("Select a Page:", ["Overview", "Sales Analysis", "Order Fulfillment", "Monthly Metrics", "KPI Analysis", "Moving Average", "Forecasting", "Download Reports"])

# Custom CSS for better UI
st.markdown("""
    <style>
    .stButton>button {
        background-color: #00A6FB;
        color: white;
        border-radius: 5px;
        padding: 10px 20px;
        font-size: 16px;
    }
    .stButton>button:hover {
        background-color: #0582CA;
    }
    .stMetric {
        background-color: #f0f2f6;
        padding: 20px;
        border-radius: 10px;
        text-align: center;
    }
    .stMetric h3 {
        font-size: 24px;
        color: #051923;
    }
    .stMetric p {
        font-size: 18px;
        color: #006494;
    }
    </style>
    """, unsafe_allow_html=True)

if page == "Overview":
    st.title("📈 Dashboard Overview")
    
    # Metrics
    total_sales = fact_sales_df["TotalSales"].sum()
    total_orders = fact_sales_df.shape[0]
    avg_processing_time = fact_sales_df["ProcessingTime"].mean()
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown(f"<div class='stMetric'><h3>💰 Total Sales</h3><p>${total_sales:,.2f}</p></div>", unsafe_allow_html=True)
    with col2:
        st.markdown(f"<div class='stMetric'><h3>📦 Total Orders</h3><p>{total_orders}</p></div>", unsafe_allow_html=True)
    with col3:
        st.markdown(f"<div class='stMetric'><h3>⏳ Avg Processing Time</h3><p>{avg_processing_time:.1f} days</p></div>", unsafe_allow_html=True)
    
    # Order Status Distribution
    st.markdown("### Order Status Distribution")
    order_status_counts = fact_sales_df["OrderStatus"].value_counts()
    fig = px.pie(order_status_counts, values=order_status_counts.values, names=order_status_counts.index,
                 title="Order Status Distribution", color_discrete_sequence=colors)
    st.plotly_chart(fig)
    
    # Monthly Order and Completion Analysis
    st.markdown("### Monthly Order and Completion Analysis")
    monthly_orders = fact_sales_df.groupby("Month").agg({"SalesID": "count", "OrderStatus": lambda x: (x == "Completed").sum()}).reset_index()
    monthly_orders.columns = ["Month", "Total Orders", "Completed Orders"]
    
    fig_bar = px.bar(monthly_orders, x="Month", y=["Completed Orders", "Total Orders"],
                     title="Monthly Total Orders and Completed Orders", 
                     color_discrete_sequence=[colors[3], colors[0]],
                     labels={"value": "Count", "variable": "Order Type"})
    st.plotly_chart(fig_bar)

elif page == "Sales Analysis":
    st.title("📊 Sales Trends")
    
    # Sales Trend Over Time
    st.markdown("### Sales Trend Over Time")
    sales_trend = fact_sales_df.groupby("Date")["TotalSales"].sum().reset_index()
    fig = px.line(sales_trend, x="Date", y="TotalSales", title="Sales Trend Over Time", color_discrete_sequence=[colors[1]])
    st.plotly_chart(fig)

elif page == "Order Fulfillment":
    st.title("📦 Order Fulfillment Analysis")
    
    # Processing Time Distribution
    st.markdown("### Processing Time Distribution")
    processing_time_dist = px.histogram(fact_sales_df, x="ProcessingTime", nbins=10, title="Processing Time Distribution", color_discrete_sequence=[colors[2]])
    st.plotly_chart(processing_time_dist)

elif page == "Monthly Metrics":
    st.title("📊 Monthly Metrics")
    
    # Monthly Metrics Table
    st.markdown("### Monthly Metrics Table")
    monthly_metrics = fact_sales_df.groupby("Month").agg({
        "TotalSales": "sum", 
        "ShippingCost": "sum", 
        "QuantitySold": "sum"}).reset_index()
    
    st.dataframe(monthly_metrics)
    
    # Monthly Total Sales
    st.markdown("### Monthly Total Sales")
    fig_sales = px.line(monthly_metrics, x="Month", y="TotalSales", title="Monthly Total Sales", color_discrete_sequence=[colors[0]])
    st.plotly_chart(fig_sales)
    
    # Monthly Shipping Cost
    st.markdown("### Monthly Shipping Cost")
    fig_shipping = px.line(monthly_metrics, x="Month", y="ShippingCost", title="Monthly Shipping Cost", color_discrete_sequence=[colors[1]])
    st.plotly_chart(fig_shipping)
    
    # Monthly Quantity Sold
    st.markdown("### Monthly Quantity Sold")
    fig_quantity = px.line(monthly_metrics, x="Month", y="QuantitySold", title="Monthly Quantity Sold", color_discrete_sequence=[colors[2]])
    st.plotly_chart(fig_quantity)

elif page == "KPI Analysis":
    st.title("📊 KPI Analysis")
    
    # KPIs
    total_orders = fact_sales_df.shape[0]
    total_sales = fact_sales_df["TotalSales"].sum()
    avg_processing_time = fact_sales_df["ProcessingTime"].mean()
    cancellation_rate = (fact_sales_df["OrderStatus"].value_counts().get("Cancelled", 0) / total_orders) * 100
    
    kpi_col1, kpi_col2 = st.columns(2)
    with kpi_col1:
        st.markdown(f"<div class='stMetric'><h3>⏳ Avg Processing Time</h3><p>{avg_processing_time:.1f} days</p></div>", unsafe_allow_html=True)
    with kpi_col2:
        st.markdown(f"<div class='stMetric'><h3>❌ Cancellation Rate</h3><p>{cancellation_rate:.1f} %</p></div>", unsafe_allow_html=True)
    
    # Processing Time by Warehouse
    st.markdown("### Processing Time by Warehouse")
    warehouse_processing = fact_sales_df.groupby("WarehouseID")["ProcessingTime"].mean().reset_index()
    if not warehouse_processing.empty:
        fig_processing = px.bar(warehouse_processing, x="WarehouseID", y="ProcessingTime", 
                                title="Avg Processing Time by Warehouse", color_discrete_sequence=[colors[1]])
        st.plotly_chart(fig_processing)
    else:
        st.write("No data available for Processing Time by Warehouse.")
    
    # Cancellation Rate by Warehouse
    st.markdown("### Cancellation Rate by Warehouse")
    warehouse_cancellations = fact_sales_df.groupby("WarehouseID")["OrderStatus"].apply(lambda x: (x == "Cancelled").mean() * 100).reset_index()
    warehouse_cancellations.columns = ["WarehouseID", "Cancellation Rate"]
    if not warehouse_cancellations.empty:
        fig_cancellations = px.bar(warehouse_cancellations, x="WarehouseID", y="Cancellation Rate", 
                                   title="Avg Cancellation Rate by Warehouse", color_discrete_sequence=[colors[2]])
        st.plotly_chart(fig_cancellations)
    else:
        st.write("No data available for Cancellation Rate by Warehouse.")

elif page == "Moving Average":
    st.title("📊 Moving Average for Next 3 Months")
    
    # Moving Average Calculation
    sales_trend = fact_sales_df.groupby("Date")["TotalSales"].sum().reset_index()
    sales_trend["Moving_Avg_3M"] = sales_trend["TotalSales"].rolling(window=3, min_periods=1).mean()
    
    # Display the data in a matrix (table) format
    st.markdown("### Moving Average Data")
    st.dataframe(sales_trend)
    
    # Moving Average Chart
    st.markdown("### Moving Average (3 Months) vs. Total Sales")
    fig = px.line(sales_trend, x="Date", y=["TotalSales", "Moving_Avg_3M"], 
                  title="Moving Average (3 Months) vs. Total Sales", 
                  labels={"value": "Sales", "variable": "Metric"}, 
                  color_discrete_sequence=[colors[1], colors[3]])
    st.plotly_chart(fig, use_container_width=True)
    
    # Download Button for the Moving Average Data
    st.markdown("### Download Moving Average Data")
    csv_data = sales_trend.to_csv(index=False).encode('utf-8')
    st.download_button(label="Download CSV", data=csv_data, file_name="moving_average_data.csv", mime="text/csv")

elif page == "Forecasting":
    st.title("🔮 Sales Forecast")
    st.write("Coming Soon: Predicting future sales using AI models.")

elif page == "Download Reports":
    st.title("📥 Download Reports")
    csv_data = fact_sales_df.to_csv(index=False).encode('utf-8')
    st.download_button(label="Download CSV", data=csv_data, file_name="sales_report.csv", mime="text/csv")

st.sidebar.write("Powered by Data Science 🚀")