import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import re

# Set page configuration
st.set_page_config(
    page_title="Warehouse Picking Analysis",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# App title
st.title("📦 Warehouse Order Picking Analysis")
st.markdown("Upload your warehouse order picking data to gain insights into performance metrics and trends.")

# Sidebar for file upload
with st.sidebar:
    st.header("Upload Data")
    uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        st.success("File uploaded successfully!")
        
        # Additional filters
        st.header("Filters")
        show_filters = st.checkbox("Apply Filters", value=False)
        
        if show_filters:
            # These will be implemented after loading the data
            pass

# Helper functions
def load_data(uploaded_file):
    """Load data from uploaded Excel file with flexible column handling"""
    try:
        df = pd.read_excel(uploaded_file)
        
        # Debug: Show available columns
        st.write("Available columns in your file:", df.columns.tolist())
        
        # Clean and normalize column names (collapse all non-alphanumerics to single underscore)
        df.columns = (
            df.columns
            .str.strip()
            .str.lower()
            .str.replace(r'[^a-z0-9]+', '_', regex=True)
            .str.replace(r'_+', '_', regex=True)
            .str.strip('_')
        )
        
        # Keep only the expected columns from the screenshot (case/space-insensitive)
        expected = {
            'warehouse_order': 'warehouse_order',
            'queue': 'queue',
            'whse_order_status': 'whse_order_status',
            'actual_quantity': 'actual_quantity',
            'number_of_wts': 'number_of_wts',
            'wo_activity_area': 'wo_activity_area',
            'processor': 'processor',
            'start_date': 'start_date',
            'start_time': 'start_time',
            'confirmation_time': 'confirmation_time'
        }
        # Map close variants by ignoring underscores
        normalized_set = {c.replace('_', ''): c for c in df.columns}
        selected_cols = {}
        for want in expected.keys():
            key = want.replace('_', '')
            if key in normalized_set:
                selected_cols[normalized_set[key]] = expected[want]
            else:
                # try relaxed match: startswith first 5 characters
                candidate = next((c for k, c in normalized_set.items() if key in k), None)
                if candidate is not None:
                    selected_cols[candidate] = expected[want]
        df = df[list(selected_cols.keys())].rename(columns=selected_cols).copy()
        
        # Handle duplicate start_time columns
        if 'start_time_1' in df.columns and 'start_time_2' in df.columns:
            # Use the first start_time column and drop the duplicate
            df['start_time'] = df['start_time_1']
            df = df.drop(['start_time_1', 'start_time_2'], axis=1)
        
        # Convert numeric columns - handle comma formatting
        if 'actual_quantity' in df.columns:
            df['actual_quantity'] = pd.to_numeric(
                df['actual_quantity'].astype(str).str.replace(',', ''), 
                errors='coerce'
            )
        
        if 'number_of_wts' in df.columns:
            df['number_of_wts'] = pd.to_numeric(
                df['number_of_wts'].astype(str).str.replace(',', ''), 
                errors='coerce'
            )
        
        # Convert date column
        if 'start_date' in df.columns:
            df['start_date'] = pd.to_datetime(df['start_date'], errors='coerce')
        
        # Convert time columns with flexible parsing including Finnish AM/PM ("ap." = AM, "ip." = PM)
        def parse_time_string(value: str):
            if pd.isna(value):
                return pd.NaT
            s = str(value).strip().lower()
            # Normalize separators and remove extra spaces
            s = re.sub(r"\s+", " ", s)
            # Extract clock portion HH:MM[:SS]
            match = re.search(r"(\d{1,2}:\d{2}(?::\d{2})?)", s)
            if not match:
                return pd.NaT
            clock = match.group(1)
            # Ensure seconds exist
            if len(clock.split(":")) == 2:
                clock = f"{clock}:00"
            # Detect Finnish AM/PM markers
            is_am = any(tok in s for tok in ["ap", "a.p", "ap."])
            is_pm = any(tok in s for tok in ["ip", "i.p", "ip."])
            try:
                if is_am or is_pm:
                    # 12-hour input
                    t = datetime.strptime(clock, "%H:%M:%S").time()
                    hour = t.hour
                    if hour == 0:
                        hour = 12
                    if is_pm and hour != 12:
                        hour = (hour + 12) % 24
                    if is_am and hour == 12:
                        hour = 0
                    return datetime.strptime(f"{hour:02d}:{t.minute:02d}:{t.second:02d}", "%H:%M:%S").time()
                # 24-hour input
                return datetime.strptime(clock, "%H:%M:%S").time()
            except Exception:
                return pd.NaT

        if 'start_time' in df.columns:
            df['start_time'] = df['start_time'].apply(parse_time_string)
        
        if 'confirmation_time' in df.columns:
            df['confirmation_time'] = df['confirmation_time'].apply(parse_time_string)
        
        # Create datetime columns for easier time calculations
        if all(col in df.columns for col in ['start_date', 'start_time']):
            df['start_datetime'] = df.apply(
                lambda row: datetime.combine(row['start_date'].date(), row['start_time']) 
                if pd.notna(row['start_date']) and pd.notna(row['start_time']) else pd.NaT, 
                axis=1
            )
        
        if all(col in df.columns for col in ['start_date', 'confirmation_time']):
            def build_end_dt(row):
                if pd.isna(row['start_date']) or pd.isna(row['confirmation_time']):
                    return pd.NaT
                start_dt = row.get('start_datetime', pd.NaT)
                end_dt = datetime.combine(row['start_date'].date(), row['confirmation_time'])
                # If end time is before start time, assume it finished after midnight
                if pd.notna(start_dt) and end_dt < start_dt:
                    end_dt = end_dt + timedelta(days=1)
                return end_dt
            df['end_datetime'] = df.apply(build_end_dt, axis=1)
        
        # Calculate processing time in minutes if we have both start and end times
        if all(col in df.columns for col in ['start_datetime', 'end_datetime']):
            df['processing_time_(min)'] = df.apply(
                lambda row: (row['end_datetime'] - row['start_datetime']).total_seconds() / 60 
                if pd.notna(row['end_datetime']) and pd.notna(row['start_datetime']) else np.nan, 
                axis=1
            )

            # Keep only realistic durations (0 to 180 minutes) to avoid placeholder or batching artifacts
            df['processing_time_valid_(min)'] = df['processing_time_(min)'].where(
                (df['processing_time_(min)'] >= 0) & (df['processing_time_(min)'] <= 180)
            )

        # Completed hour (rounded to hour) for hourly KPIs
        if 'end_datetime' in df.columns:
            df['completed_hour'] = pd.to_datetime(df['end_datetime']).dt.floor('H')
            df['completed_date'] = pd.to_datetime(df['end_datetime']).dt.date
        
        st.success("✅ Data loaded successfully!")
        return df
        
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")
        st.write("File structure:", df.columns.tolist() if 'df' in locals() else "No data loaded")
        return None

def calculate_kpis(df):
    """Calculate key performance indicators with error handling"""
    kpis = {}
    
    try:
        # Basic metrics
        kpis['total_orders'] = len(df)
        
        # Handle completion rate
        if 'confirmation_time' in df.columns:
            kpis['completed_orders'] = len(df[df['confirmation_time'].notna()])
            kpis['completion_rate'] = kpis['completed_orders'] / kpis['total_orders'] * 100
        else:
            kpis['completed_orders'] = 0
            kpis['completion_rate'] = 0
        
        # Time metrics
        if 'processing_time_valid_(min)' in df.columns:
            kpis['avg_processing_time'] = df['processing_time_valid_(min)'].mean()
            kpis['median_processing_time'] = df['processing_time_valid_(min)'].median()
        else:
            kpis['avg_processing_time'] = 0
            kpis['median_processing_time'] = 0
        
        # Efficiency metrics
        if 'actual_quantity' in df.columns:
            kpis['total_items_picked'] = df['actual_quantity'].sum()
            kpis['avg_items_per_order'] = df['actual_quantity'].mean()
        else:
            kpis['total_items_picked'] = 0
            kpis['avg_items_per_order'] = 0

        # Orders completed in an hour (peak hourly rate)
        if 'completed_hour' in df.columns:
            per_hour = df.dropna(subset=['completed_hour']).groupby('completed_hour').size()
            kpis['orders_per_hour_peak'] = int(per_hour.max()) if not per_hour.empty else 0
        else:
            kpis['orders_per_hour_peak'] = 0
        
        # No items_per_minute KPI to avoid dependency on missing fields
        
        # Worker metrics
        if 'processor' in df.columns:
            kpis['unique_workers'] = df['processor'].nunique()
            kpis['busiest_worker'] = df['processor'].value_counts().idxmax() if kpis['unique_workers'] > 0 else "N/A"
        else:
            kpis['unique_workers'] = 0
            kpis['busiest_worker'] = "N/A"
        
        # Area metrics
        if 'wo_activity_area' in df.columns:
            kpis['unique_areas'] = df['wo_activity_area'].nunique()
        else:
            kpis['unique_areas'] = 0
        
        return kpis
        
    except Exception as e:
        st.error(f"Error calculating KPIs: {str(e)}")
        # Return default values
        return {
            'total_orders': len(df),
            'completed_orders': 0,
            'completion_rate': 0,
            'avg_processing_time': 0,
            'median_processing_time': 0,
            'total_items_picked': 0,
            'avg_items_per_order': 0,
            'avg_items_per_minute': 0,
            'unique_workers': 0,
            'busiest_worker': "N/A",
            'unique_areas': 0
        }

def create_visualizations(df, kpis):
    """Create visualizations for the data"""
    visualizations = {}
    
    try:
        # Orders by status
        if 'whse_order_status' in df.columns:
            status_counts = df['whse_order_status'].value_counts()
            fig_status = px.pie(values=status_counts.values, names=status_counts.index, 
                                title='Orders by Status')
            visualizations['status_pie'] = fig_status
        
        # Processing time distribution
        if 'processing_time_valid_(min)' in df.columns:
            fig_time_dist = px.histogram(df, x='processing_time_valid_(min)', 
                                        title='Distribution of Processing Time (Minutes)',
                                        nbins=50)
            visualizations['time_dist'] = fig_time_dist
        
        # Orders by activity area
        if 'wo_activity_area' in df.columns:
            area_counts = df['wo_activity_area'].value_counts()
            fig_area = px.bar(x=area_counts.index, y=area_counts.values, 
                             title='Orders by Activity Area',
                             labels={'x': 'Activity Area', 'y': 'Number of Orders'})
            visualizations['area_bar'] = fig_area
        
        # Orders over time by start hour (reference)
        if 'start_datetime' in df.columns:
            df['StartHour'] = df['start_datetime'].dt.hour
            hourly_orders = df.groupby('StartHour').size()
            fig_hourly = px.line(x=hourly_orders.index, y=hourly_orders.values,
                                title='Orders Started by Hour of Day',
                                labels={'x': 'Hour of Day', 'y': 'Number of Orders'})
            visualizations['hourly_orders'] = fig_hourly

        # Completed orders per hour and quantities per hour (primary KPI)
        if 'completed_hour' in df.columns:
            completed_hourly = df.dropna(subset=['completed_hour']).groupby('completed_hour').agg({
                'warehouse_order': 'count',
                'actual_quantity': 'sum'
            }).rename(columns={'warehouse_order': 'orders_completed', 'actual_quantity': 'quantity_completed'})

            if not completed_hourly.empty:
                fig_completed = px.bar(
                    completed_hourly,
                    x=completed_hourly.index,
                    y='orders_completed',
                    title='Orders Completed per Hour',
                    labels={'x': 'Hour', 'orders_completed': 'Orders Completed'}
                )
                visualizations['completed_per_hour'] = fig_completed

                fig_qty = px.bar(
                    completed_hourly,
                    x=completed_hourly.index,
                    y='quantity_completed',
                    title='Quantity Completed per Hour',
                    labels={'x': 'Hour', 'quantity_completed': 'Quantity Completed'}
                )
                visualizations['quantity_per_hour'] = fig_qty
                visualizations['completed_hourly_table'] = completed_hourly

        # Completed orders per day and quantities per day
        if 'completed_date' in df.columns:
            completed_daily = df.dropna(subset=['completed_date']).groupby('completed_date').agg({
                'warehouse_order': 'count',
                'actual_quantity': 'sum'
            }).rename(columns={'warehouse_order': 'orders_completed', 'actual_quantity': 'quantity_completed'})

            if not completed_daily.empty:
                fig_daily_orders = px.bar(
                    completed_daily,
                    x=completed_daily.index,
                    y='orders_completed',
                    title='Orders Completed per Day',
                    labels={'x': 'Date', 'orders_completed': 'Orders Completed'}
                )
                visualizations['completed_per_day'] = fig_daily_orders

                fig_daily_qty = px.bar(
                    completed_daily,
                    x=completed_daily.index,
                    y='quantity_completed',
                    title='Quantity Completed per Day',
                    labels={'x': 'Date', 'quantity_completed': 'Quantity Completed'}
                )
                visualizations['quantity_per_day'] = fig_daily_qty
                visualizations['completed_daily_table'] = completed_daily
        
        # Top performers
        if 'processor' in df.columns and kpis['unique_workers'] > 0:
            worker_orders = df.groupby('processor').agg({
                'warehouse_order': 'count',
                'actual_quantity': 'sum'
            }).rename(columns={'warehouse_order': 'order_count', 'actual_quantity': 'total_quantity'})

            fig_workers = px.bar(worker_orders.nlargest(10, 'order_count'), 
                                x=worker_orders.nlargest(10, 'order_count').index, 
                                y='order_count',
                                title='Top 10 Workers by Number of Orders')
            visualizations['top_workers'] = fig_workers
        
        return visualizations
        
    except Exception as e:
        st.error(f"Error creating visualizations: {str(e)}")
        return visualizations

# Main app logic
if uploaded_file is not None:
    # Load data
    with st.spinner('Loading data...'):
        df = load_data(uploaded_file)
    
    if df is not None:
        # Calculate KPIs
        kpis = calculate_kpis(df)
        
        # Display KPIs
        st.header("Key Performance Indicators")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Total Orders", kpis['total_orders'])
            st.metric("Completed Orders", kpis['completed_orders'])
        
        with col2:
            st.metric("Completion Rate", f"{kpis['completion_rate']:.1f}%")
            st.metric("Total Items Picked", f"{kpis['total_items_picked']:,.0f}")
        
        with col3:
            st.metric("Avg Processing Time", f"{max(kpis['avg_processing_time'], 0):.1f} min")
            st.metric("Avg Items per Order", f"{kpis['avg_items_per_order']:.1f}")
        
        with col4:
            st.metric("Unique Workers", kpis['unique_workers'])
            st.metric("Orders Completed in a Peak Hour", kpis.get('orders_per_hour_peak', 0))
        
        # Create and display visualizations
        st.header("Data Visualizations")
        
        with st.spinner('Generating visualizations...'):
            visuals = create_visualizations(df, kpis)
            
            # Display charts in tabs
            tab_names = ["Order Status", "Processing Time", "Activity Areas", "Hourly Starts", "Hourly Completions", "Hourly Quantities", "Daily Completions", "Daily Quantities"]
            if 'top_workers' in visuals:
                tab_names.extend(["Top Workers"])
            
            tabs = st.tabs(tab_names)
            
            tab_index = 0
            
            with tabs[tab_index]:
                if 'status_pie' in visuals:
                    st.plotly_chart(visuals['status_pie'], use_container_width=True)
                else:
                    st.info("No status data available")
            tab_index += 1
            
            with tabs[tab_index]:
                if 'time_dist' in visuals:
                    st.plotly_chart(visuals['time_dist'], use_container_width=True)
                else:
                    st.info("No processing time data available")
            tab_index += 1
            
            with tabs[tab_index]:
                if 'area_bar' in visuals:
                    st.plotly_chart(visuals['area_bar'], use_container_width=True)
                else:
                    st.info("No activity area data available")
            tab_index += 1
            
            with tabs[tab_index]:
                if 'hourly_orders' in visuals:
                    st.plotly_chart(visuals['hourly_orders'], use_container_width=True)
                else:
                    st.info("No hourly start distribution data available")
            tab_index += 1

            with tabs[tab_index]:
                if 'completed_per_hour' in visuals:
                    st.plotly_chart(visuals['completed_per_hour'], use_container_width=True)
                    if 'completed_hourly_table' in visuals:
                        st.dataframe(visuals['completed_hourly_table'])
                else:
                    st.info("No hourly completion data available")
            tab_index += 1

            with tabs[tab_index]:
                if 'quantity_per_hour' in visuals:
                    st.plotly_chart(visuals['quantity_per_hour'], use_container_width=True)
                else:
                    st.info("No hourly quantity data available")
            tab_index += 1

            with tabs[tab_index]:
                if 'completed_per_day' in visuals:
                    st.plotly_chart(visuals['completed_per_day'], use_container_width=True)
                    if 'completed_daily_table' in visuals:
                        st.dataframe(visuals['completed_daily_table'])
                else:
                    st.info("No daily completion data available")
            tab_index += 1

            with tabs[tab_index]:
                if 'quantity_per_day' in visuals:
                    st.plotly_chart(visuals['quantity_per_day'], use_container_width=True)
                else:
                    st.info("No daily quantity data available")
            tab_index += 1
            
            if 'top_workers' in visuals:
                with tabs[tab_index]:
                    st.plotly_chart(visuals['top_workers'], use_container_width=True)
                tab_index += 1
            
            # No worker efficiency tab (removed dependency on items_per_minute)
        
        # Raw data view
        st.header("Raw Data")
        
        if st.checkbox("Show raw data"):
            st.dataframe(df)
            
            # Download button for processed data
            csv = df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="Download processed data as CSV",
                data=csv,
                file_name="processed_warehouse_data.csv",
                mime="text/csv"
            )
        
        # Advanced analysis section
        st.header("Advanced Analysis")
        
        # Time analysis by activity area
        if 'wo_activity_area' in df.columns and 'processing_time_(min)' in df.columns:
            st.subheader("Processing Time by Activity Area")
            area_time = df.groupby('wo_activity_area')['processing_time_(min)'].agg(['mean', 'median', 'count']).round(1)
            st.dataframe(area_time)
        
        # Worker performance analysis (no items_per_minute dependency)
        if 'processor' in df.columns and kpis['unique_workers'] > 0:
            st.subheader("Worker Performance Analysis")
            worker_stats = df.groupby('processor').agg({
                'warehouse_order': 'count',
                'actual_quantity': 'sum',
                'processing_time_valid_(min)': 'mean'
            }).round(2).rename(columns={
                'warehouse_order': 'Orders Processed',
                'actual_quantity': 'Total Items Picked',
                'processing_time_valid_(min)': 'Avg Processing Time (min)'
            })
            st.dataframe(worker_stats.sort_values('Orders Processed', ascending=False))
        
    else:
        st.error("Failed to load data. Please check the file format.")
else:
    # Show instructions when no file is uploaded
    st.info("👈 Please upload an Excel file to begin analysis.")
    
    # Example of expected format
    st.subheader("Expected Data Format")
    st.markdown("""
    Your Excel file should contain the following columns:
    - **Warehouse Order**: Order identifier
    - **Queue**: Queue information (e.g., MP01-PICK, MP3-PICK)
    - **Whse Order Status**: Order status (e.g., C for completed)
    - **Actual quantity**: Number of items picked
    - **Number of WTs**: Number of warehouse tasks
    - **WO Activity Area**: Activity area (e.g., MP01, MP23)
    - **Processor**: Worker identifier
    - **Start Date**: Start date of the task
    - **Start Time**: Start time of the task
    - **Confirmation Time**: Completion time of the task
    """)
    
    # Add sample data download
    sample_data = pd.DataFrame({
        'Warehouse Order': [2012723539, 2012723532],
        'Queue': ['MP01-PICK', 'MP01-PICK'],
        'Whse Order Status': ['C', 'C'],
        'Actual quantity': [85.0, 84.0],
        'Number of WTs': [66.0, 52.0],
        'WO Activity Area': ['MP01', 'MP01'],
        'Processor': ['FIOLLMUL', 'FIJOOVIR'],
        'Start Date': ['2025-08-15', '2025-08-15'],
        'Start Time': ['01:18:35', '01:25:15'],
        'Confirmation Time': ['01:45:12', '01:54:06']
    })
    
    csv = sample_data.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="Download sample data format",
        data=csv,
        file_name="sample_warehouse_data.csv",
        mime="text/csv"
    )

# Footer
st.markdown("---")
st.markdown("### 📊 Warehouse Analytics Tool v1.0")
