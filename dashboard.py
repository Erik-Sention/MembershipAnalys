import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from membership_analyzer import MembershipAnalyzer
import os
import hashlib
import time
from datetime import datetime, timedelta

# Configure Streamlit page
st.set_page_config(
    page_title="Membership Analysis Dashboard",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #2c3e50;
        text-align: center;
        margin-bottom: 2rem;
        font-weight: 600;
        border-bottom: 2px solid #34495e;
        padding-bottom: 1rem;
    }
    .metric-container {
        background-color: #f8f9fa;
        padding: 1.5rem;
        border-radius: 8px;
        margin: 0.5rem 0;
        border: 1px solid #dee2e6;
    }
    .insight-box {
        background-color: #ffffff;
        border: 1px solid #d1ecf1;
        border-left: 4px solid #17a2b8;
        padding: 1.2rem;
        margin: 1rem 0;
        border-radius: 4px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .login-container {
        background-color: #ffffff;
        border: 1px solid #dee2e6;
        border-radius: 8px;
        padding: 2rem;
        margin: 2rem auto;
        max-width: 400px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .professional-text {
        color: #495057;
        font-size: 1rem;
        line-height: 1.6;
    }
</style>
""", unsafe_allow_html=True)

def check_password():
    """Returns True if the user has entered the correct password."""
    
    def hash_password(password):
        """Hash password for basic security"""
        return hashlib.sha256(password.encode()).hexdigest()
    
    # Security settings from Streamlit secrets
    try:
        CORRECT_USERNAME = st.secrets["auth"]["username"]
        CORRECT_PASSWORD = st.secrets["auth"]["password"]
        CORRECT_PASSWORD_HASH = hash_password(CORRECT_PASSWORD)
    except KeyError:
        st.error("Authentication credentials not found in secrets. Please configure your secrets properly.")
        st.stop()
    
    SESSION_TIMEOUT_MINUTES = 30  # Auto-logout after 30 minutes of inactivity
    
    def check_session_timeout():
        """Check if session has timed out due to inactivity"""
        if "last_activity" in st.session_state:
            time_since_activity = datetime.now() - st.session_state["last_activity"]
            if time_since_activity > timedelta(minutes=SESSION_TIMEOUT_MINUTES):
                # Session has timed out
                st.session_state["authenticated"] = False
                st.session_state["session_expired"] = True
                return True
        return False
    
    def update_activity():
        """Update last activity timestamp"""
        st.session_state["last_activity"] = datetime.now()
    
    def login_form():
        """Login form"""
        st.markdown("---")
        
        # Show session expired message if applicable
        if st.session_state.get("session_expired", False):
            st.warning("Session expired due to inactivity. Please login again.")
            st.session_state["session_expired"] = False
        
        st.markdown("### Authentication Required")
        st.markdown('<p class="professional-text">Please enter your credentials to access the Membership Analysis Dashboard.</p>', unsafe_allow_html=True)
        
        with st.form("login_form"):
            username = st.text_input("Username", placeholder="Enter username")
            password = st.text_input("Password", type="password", placeholder="Enter password")
            submit_button = st.form_submit_button("Login")
            
            if submit_button:
                if username == CORRECT_USERNAME and hash_password(password) == CORRECT_PASSWORD_HASH:
                    st.session_state["authenticated"] = True
                    st.session_state["session_expired"] = False
                    update_activity()  # Set initial activity timestamp
                    st.success("Login successful! Redirecting...")
                    st.rerun()
                else:
                    st.error("Invalid username or password. Please try again.")
                    st.session_state["authenticated"] = False
    
    # Check if already authenticated
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False
    
    # Check for session timeout if user is authenticated
    if st.session_state["authenticated"]:
        if check_session_timeout():
            # Session expired, show login again
            st.session_state["authenticated"] = False
        else:
            # Update activity on each interaction
            update_activity()
    
    if not st.session_state["authenticated"]:
        # Show login page
        st.markdown('<h1 class="main-header">Membership Analysis Dashboard</h1>', unsafe_allow_html=True)
        
        # Add professional branding
        st.markdown("""
        <div class="login-container">
            <div style="text-align: center; margin-bottom: 2rem;">
                <h3 style="color: #2c3e50; margin-bottom: 0.5rem;">Secure Access Portal</h3>
                <p class="professional-text">Enterprise Membership Analytics Platform</p>
            </div>
        """, unsafe_allow_html=True)
        
        login_form()
        
        st.markdown("</div>", unsafe_allow_html=True)
        
        # Add footer info
        st.markdown("---")
        st.markdown("""
        <div style="text-align: center; color: #6c757d; font-size: 0.9rem; margin-top: 2rem;">
            <p>Advanced Membership Analytics Platform</p>
            <p>Secured with enterprise-grade authentication</p>
            <p>Auto-logout after {SESSION_TIMEOUT_MINUTES} minutes of inactivity</p>
        </div>
        """.format(SESSION_TIMEOUT_MINUTES=SESSION_TIMEOUT_MINUTES), unsafe_allow_html=True)
        
        return False
    
    return True

def logout():
    """Logout function"""
    st.session_state["authenticated"] = False
    st.rerun()

def main():
    # Check authentication first
    if not check_password():
        return
    
    # Add logout button and session info to sidebar
    with st.sidebar:
        st.markdown("---")
        
        # Show session status
        if "last_activity" in st.session_state:
            time_since_activity = datetime.now() - st.session_state["last_activity"]
            remaining_minutes = 30 - int(time_since_activity.total_seconds() / 60)
            
            if remaining_minutes > 0:
                if remaining_minutes <= 5:
                    st.warning(f"Session expires in {remaining_minutes} minutes")
                else:
                    st.info(f"Session active ({remaining_minutes} min remaining)")
            
        if st.button("Logout", help="Logout and return to login screen"):
            logout()
    
    st.markdown('<h1 class="main-header">Membership Analysis Dashboard</h1>', unsafe_allow_html=True)
    
    # Sidebar for file upload and sheet selection
    st.sidebar.header("Data Input")
    
    # Check for auto-load file in root directory
    auto_load_files = [
        'membership_data.xlsx',
        'memberships.xlsx', 
        'data.xlsx',
        'Masterdokument Memberships Antal.xlsx'  # Your actual filename
    ]
    
    auto_file = None
    for filename in auto_load_files:
        if os.path.exists(filename):
            auto_file = filename
            break
    
    if auto_file:
        st.sidebar.success(f"Data loaded: {auto_file}")
        use_auto_file = st.sidebar.button("Reload Data", help="Refresh data from the file")
        
        uploaded_file = auto_file
        temp_file_path = auto_file  # Use file directly, no need to save temp
    else:
        st.sidebar.error("**No data file found**")
        st.sidebar.info("**Please add your Excel file to the root directory with one of these names:**")
        st.sidebar.write("• `membership_data.xlsx`")
        st.sidebar.write("• `memberships.xlsx`") 
        st.sidebar.write("• `data.xlsx`")
        st.sidebar.write("• `Masterdokument Memberships Antal.xlsx`")
        uploaded_file = None
    
    if uploaded_file is not None:
        try:
            # Initialize analyzer
            analyzer = MembershipAnalyzer(temp_file_path)
            
            # Load and process data
            with st.spinner("Loading data..."):
                if analyzer.load_data():
                    analyzer.clean_and_standardize_data()
                    
                    # Sheet selection
                    available_sheets = analyzer.get_available_sheets()
                    if available_sheets:
                        selected_sheet = st.sidebar.selectbox(
                            "Select Sheet to Analyze",
                            available_sheets,
                            help="Choose which sheet/location to analyze"
                        )
                        
                                            # Analysis mode selection
                    st.sidebar.header("Analysis Mode")
                    analysis_mode = st.sidebar.radio(
                        "Choose Analysis Type",
                        ["Single Location Analysis", "Two Location Comparison", "Multi-Location Comparison"],
                        help="Analyze one location, compare two locations side-by-side, or compare all locations"
                    )
                    
                    if analysis_mode == "Single Location Analysis":
                        # Row filtering options
                        st.sidebar.subheader("Filter Membership Types")
                        
                        # Get available membership types for the selected sheet
                        if selected_sheet in analyzer.data:
                            df = analyzer.data[selected_sheet]
                            membership_col = None
                            for col in df.columns:
                                if 'membership' in col.lower() or col.strip() == 'Membership':
                                    membership_col = col
                                    break
                            if not membership_col and len(df.columns) > 0:
                                membership_col = df.columns[0]
                            
                            if membership_col:
                                all_memberships = df[membership_col].astype(str).str.strip().tolist()
                                all_memberships = [m for m in all_memberships if m and m.lower() not in ['nan', 'none', '']]
                                
                                # Show all membership types with checkboxes for better visibility
                                st.sidebar.write("**Available membership types:**")
                                
                                # Create a container for better layout
                                with st.sidebar.container():
                                    # Use session state to track selections
                                    if 'selected_memberships' not in st.session_state:
                                        st.session_state.selected_memberships = all_memberships.copy()
                                    
                                    # Select/Deselect all buttons
                                    col1, col2 = st.columns(2)
                                    with col1:
                                        if st.button("Select All", key="select_all_single"):
                                            st.session_state.selected_memberships = all_memberships.copy()
                                    with col2:
                                        if st.button("Clear All", key="clear_all_single"):
                                            st.session_state.selected_memberships = []
                                    
                                    # Individual checkboxes for each membership type
                                    selected_memberships = []
                                    for i, membership in enumerate(all_memberships):
                                        is_selected = membership in st.session_state.selected_memberships
                                        
                                        if st.checkbox(
                                            membership, 
                                            value=is_selected, 
                                            key=f"single_membership_{i}_{membership[:20]}"
                                        ):
                                            selected_memberships.append(membership)
                                    
                                    # Update session state
                                    st.session_state.selected_memberships = selected_memberships
                                    
                                    # Show count
                                    st.write(f"Selected: {len(selected_memberships)} of {len(all_memberships)}")
                                
                                # Filter the data
                                if selected_memberships:
                                    analyzer.data[selected_sheet] = df[df[membership_col].isin(selected_memberships)].copy()
                        
                        # Analysis options for single location
                        st.sidebar.subheader("Display Options")
                        show_trends = st.sidebar.checkbox("Monthly Trends", True)
                        show_pie = st.sidebar.checkbox("Membership Distribution", True)
                        show_heatmap = st.sidebar.checkbox("Activity Heatmap", True)
                        show_top_performers = st.sidebar.checkbox("Top Performers", True)
                        show_insights = st.sidebar.checkbox("Automated Insights", True)
                        
                        # Main dashboard content
                        if selected_sheet:
                            display_dashboard(analyzer, selected_sheet, {
                                'trends': show_trends,
                                'pie': show_pie,
                                'heatmap': show_heatmap,
                                'top_performers': show_top_performers,
                                'insights': show_insights
                            })
                    
                    elif analysis_mode == "Two Location Comparison":
                        # Two location comparison options
                        st.sidebar.subheader("Select Locations to Compare")
                        location_1 = st.sidebar.selectbox(
                            "First Location",
                            available_sheets,
                            help="Choose the first location to compare"
                        )
                        
                        location_2 = st.sidebar.selectbox(
                            "Second Location",
                            [sheet for sheet in available_sheets if sheet != location_1],
                            help="Choose the second location to compare"
                        )
                        
                        # Row filtering for two location comparison
                        st.sidebar.subheader("Filter Membership Types")
                        
                        # Get all unique membership types from both locations
                        all_memberships_set = set()
                        for loc in [location_1, location_2]:
                            if loc and loc in analyzer.data:
                                df = analyzer.data[loc]
                                membership_col = None
                                for col in df.columns:
                                    if 'membership' in col.lower() or col.strip() == 'Membership':
                                        membership_col = col
                                        break
                                if not membership_col and len(df.columns) > 0:
                                    membership_col = df.columns[0]
                                
                                if membership_col:
                                    memberships = df[membership_col].astype(str).str.strip().tolist()
                                    all_memberships_set.update([m for m in memberships if m and m.lower() not in ['nan', 'none', '']])
                        
                        all_memberships_list = sorted(list(all_memberships_set))
                        
                        if all_memberships_list:
                            st.sidebar.write("**Available membership types:**")
                            
                            with st.sidebar.container():
                                # Use session state for two location comparison
                                if 'selected_memberships_comp' not in st.session_state:
                                    st.session_state.selected_memberships_comp = all_memberships_list.copy()
                                
                                # Select/Deselect all buttons
                                col1, col2 = st.columns(2)
                                with col1:
                                    if st.button("Select All", key="select_all_comp"):
                                        st.session_state.selected_memberships_comp = all_memberships_list.copy()
                                with col2:
                                    if st.button("Clear All", key="clear_all_comp"):
                                        st.session_state.selected_memberships_comp = []
                                
                                # Individual checkboxes
                                selected_memberships_comp = []
                                for i, membership in enumerate(all_memberships_list):
                                    is_selected = membership in st.session_state.selected_memberships_comp
                                    
                                    if st.checkbox(
                                        membership, 
                                        value=is_selected, 
                                        key=f"comp_membership_{i}_{membership[:20]}"
                                    ):
                                        selected_memberships_comp.append(membership)
                                
                                # Update session state
                                st.session_state.selected_memberships_comp = selected_memberships_comp
                                
                                # Show count
                                st.write(f"Selected: {len(selected_memberships_comp)} of {len(all_memberships_list)}")
                            
                            # Filter both locations
                            for loc in [location_1, location_2]:
                                if loc and loc in analyzer.data and selected_memberships_comp:
                                    df = analyzer.data[loc]
                                    membership_col = None
                                    for col in df.columns:
                                        if 'membership' in col.lower() or col.strip() == 'Membership':
                                            membership_col = col
                                            break
                                    if not membership_col and len(df.columns) > 0:
                                        membership_col = df.columns[0]
                                    
                                    if membership_col:
                                        analyzer.data[loc] = df[df[membership_col].isin(selected_memberships_comp)].copy()
                        
                        st.sidebar.subheader("Display Options")
                        show_trends_comp = st.sidebar.checkbox("Monthly Trends Comparison", True)
                        show_pie_comp = st.sidebar.checkbox("Membership Distribution Comparison", True)
                        show_heatmap_comp = st.sidebar.checkbox("Activity Heatmaps", True)
                        show_top_performers_comp = st.sidebar.checkbox("Top Performers Comparison", True)
                        show_insights_comp = st.sidebar.checkbox("Comparative Insights", True)
                        
                        # Display two location comparison
                        if location_1 and location_2:
                            display_two_location_comparison(analyzer, location_1, location_2, {
                                'trends': show_trends_comp,
                                'pie': show_pie_comp,
                                'heatmap': show_heatmap_comp,
                                'top_performers': show_top_performers_comp,
                                'insights': show_insights_comp
                            })
                    
                    else:  # Multi-Location Comparison
                        # Row filtering for multi-location comparison
                        st.sidebar.subheader("Filter Membership Types")
                        
                        # Get all unique membership types from all locations
                        all_memberships_set = set()
                        for sheet_name in available_sheets:
                            if sheet_name in analyzer.data:
                                df = analyzer.data[sheet_name]
                                membership_col = None
                                for col in df.columns:
                                    if 'membership' in col.lower() or col.strip() == 'Membership':
                                        membership_col = col
                                        break
                                if not membership_col and len(df.columns) > 0:
                                    membership_col = df.columns[0]
                                
                                if membership_col:
                                    memberships = df[membership_col].astype(str).str.strip().tolist()
                                    all_memberships_set.update([m for m in memberships if m and m.lower() not in ['nan', 'none', '']])
                        
                        all_memberships_list = sorted(list(all_memberships_set))
                        
                        if all_memberships_list:
                            st.sidebar.write("**Available membership types:**")
                            
                            with st.sidebar.container():
                                # Use session state for multi-location comparison
                                if 'selected_memberships_multi' not in st.session_state:
                                    st.session_state.selected_memberships_multi = all_memberships_list.copy()
                                
                                # Select/Deselect all buttons
                                col1, col2 = st.columns(2)
                                with col1:
                                    if st.button("Select All", key="select_all_multi"):
                                        st.session_state.selected_memberships_multi = all_memberships_list.copy()
                                with col2:
                                    if st.button("Clear All", key="clear_all_multi"):
                                        st.session_state.selected_memberships_multi = []
                                
                                # Individual checkboxes
                                selected_memberships_multi = []
                                for i, membership in enumerate(all_memberships_list):
                                    is_selected = membership in st.session_state.selected_memberships_multi
                                    
                                    if st.checkbox(
                                        membership, 
                                        value=is_selected, 
                                        key=f"multi_membership_{i}_{membership[:20]}"
                                    ):
                                        selected_memberships_multi.append(membership)
                                
                                # Update session state
                                st.session_state.selected_memberships_multi = selected_memberships_multi
                                
                                # Show count
                                st.write(f"Selected: {len(selected_memberships_multi)} of {len(all_memberships_list)}")
                            
                            # Filter all locations
                            if selected_memberships_multi:
                                for sheet_name in available_sheets:
                                    if sheet_name in analyzer.data:
                                        df = analyzer.data[sheet_name]
                                        membership_col = None
                                        for col in df.columns:
                                            if 'membership' in col.lower() or col.strip() == 'Membership':
                                                membership_col = col
                                                break
                                        if not membership_col and len(df.columns) > 0:
                                            membership_col = df.columns[0]
                                        
                                        if membership_col:
                                            analyzer.data[sheet_name] = df[df[membership_col].isin(selected_memberships_multi)].copy()
                        
                        # Comparison options
                        st.sidebar.subheader("Comparison Options")
                        show_totals_comparison = st.sidebar.checkbox("Total Memberships Comparison", True)
                        show_monthly_comparison = st.sidebar.checkbox("Monthly Performance Trends", True)
                        show_growth_comparison = st.sidebar.checkbox("Growth Rate Comparison", True)
                        show_distribution_comparison = st.sidebar.checkbox("Membership Distribution Comparison", True)
                        show_summary_table = st.sidebar.checkbox("Summary Statistics Table", True)
                        
                        # Display comparison dashboard
                        display_comparison_dashboard(analyzer, {
                            'totals': show_totals_comparison,
                            'monthly': show_monthly_comparison,
                            'growth': show_growth_comparison,
                            'distribution': show_distribution_comparison,
                            'table': show_summary_table
                        })
                else:
                    st.error("Failed to load the Excel file. Please check the file format.")
        
        except Exception as e:
            st.error(f"Error processing file: {str(e)}")
        
        finally:
            # Clean up temporary file (only if it's not the auto-loaded file)
            if 'temp_' in temp_file_path and os.path.exists(temp_file_path):
                os.remove(temp_file_path)
    
    else:
        # Show sample data format
        st.info("👆 Please upload your Excel file to begin analysis")
        
        st.subheader("📋 Expected Data Format")
        st.write("Your Excel file should contain sheets with names like:")
        st.write("- Stockholm 2024")
        st.write("- Stockholm 2025") 
        st.write("- Göteborg 2024")
        st.write("- Göteborg 2025")
        
        st.write("Each sheet should have columns for:")
        st.write("- Membership types (first column)")
        st.write("- Monthly data (Januar 2024, Februari 2024, etc.)")
        st.write("- Total column")
        st.write("- Percentage column")
        
        # Show sample data structure
        sample_data = {
            'Membership': [
                'Löpande Membership Standard',
                'Löpande Membership Premium',
                'Membership Aktivitus Iform 4 mån'
            ],
            'Januar 2024': [15, 11, 17],
            'Februari 2024': [25, 5, 6],
            'Mars 2024': [7, 10, 8],
            'Totalt per medlemskap 2024': [162, 87, 51],
            'Procent (%) av memberships': [33.6, 18.0, 10.6]
        }
        
        st.subheader("Sample Data Structure")
        st.dataframe(pd.DataFrame(sample_data))

def display_dashboard(analyzer, sheet_name, options):
    """Display the main dashboard content"""
    
    st.header(f"Analysis for: {sheet_name}")
    
    # Generate insights first
    if options['insights']:
        with st.spinner("Generating insights..."):
            insights = analyzer.generate_insights(sheet_name)
            
            if insights:
                st.subheader("Key Insights")
                for insight in insights:
                    st.markdown(f'<div class="insight-box">{insight}</div>', unsafe_allow_html=True)
    
    # Create columns for layout
    col1, col2 = st.columns(2)
    
    # Monthly trends chart
    if options['trends']:
        with col1:
            st.subheader("Monthly Membership Trends")
            try:
                fig_trends = analyzer.create_monthly_trend_chart(sheet_name)
                if fig_trends:
                    st.plotly_chart(fig_trends, use_container_width=True)
                else:
                    st.warning("Could not generate monthly trends chart")
            except Exception as e:
                st.error(f"Error creating trends chart: {str(e)}")
    
    # Pie chart
    if options['pie']:
        with col2:
            st.subheader("Membership Type Distribution")
            try:
                fig_pie = analyzer.create_membership_type_pie_chart(sheet_name)
                if fig_pie:
                    st.plotly_chart(fig_pie, use_container_width=True)
                else:
                    st.warning("Could not generate pie chart")
            except Exception as e:
                st.error(f"Error creating pie chart: {str(e)}")
    
    # Heatmap (full width)
    if options['heatmap']:
        st.subheader("Membership Activity Heatmap")
        try:
            fig_heatmap = analyzer.create_heatmap(sheet_name)
            if fig_heatmap:
                st.plotly_chart(fig_heatmap, use_container_width=True)
            else:
                st.warning("Could not generate heatmap")
        except Exception as e:
            st.error(f"Error creating heatmap: {str(e)}")
    
    # Top performers
    if options['top_performers']:
        st.subheader("Top Performing Membership Types")
        try:
            fig_top = analyzer.create_top_performers_chart(sheet_name)
            if fig_top:
                st.plotly_chart(fig_top, use_container_width=True)
            else:
                st.warning("Could not generate top performers chart")
        except Exception as e:
            st.error(f"Error creating top performers chart: {str(e)}")
    
    # Raw data view
    if st.checkbox("Show Raw Data"):
        st.subheader("Raw Data")
        if sheet_name in analyzer.data:
            st.dataframe(analyzer.data[sheet_name])
    
    # Export options
    st.subheader("Export Options")
    col_export1, col_export2 = st.columns(2)
    
    with col_export1:
        if st.button("Export Summary to Excel"):
            try:
                output_file = f"membership_summary_{sheet_name.replace(' ', '_')}.xlsx"
                if analyzer.export_summary(output_file):
                    st.success(f"Data exported to {output_file}")
                else:
                    st.error("Failed to export data")
            except Exception as e:
                st.error(f"Export error: {str(e)}")
    
    with col_export2:
        if st.button("Generate Report"):
            st.info("Report generation feature coming soon!")

def display_comparison_dashboard(analyzer, options):
    """Display the multi-location comparison dashboard"""
    
    st.header("🏢 Multi-Location Comparison")
    
    # Get comparison data
    comparison_df = analyzer.create_location_comparison_summary()
    
    if comparison_df.empty:
        st.error("No data available for comparison")
        return
    
    # Summary metrics
    st.subheader("Overview")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        total_locations = len(comparison_df)
        st.metric("Total Locations", total_locations)
    
    with col2:
        total_memberships = comparison_df['Total_Memberships'].sum()
        st.metric("Total Memberships", f"{total_memberships:,.0f}")
    
    with col3:
        best_location = comparison_df.loc[comparison_df['Total_Memberships'].idxmax(), 'Sheet_Name']
        best_count = comparison_df['Total_Memberships'].max()
        st.metric("Top Performer", best_location, f"{best_count:,.0f}")
    
    with col4:
        avg_growth = comparison_df['Growth_Rate'].mean()
        st.metric("Avg Growth Rate", f"{avg_growth:.1f}%")
    
    # Charts
    if options['totals'] or options['monthly'] or options['growth']:
        fig_totals, fig_monthly, fig_growth = analyzer.create_location_comparison_charts()
        
        if options['totals'] and fig_totals:
            st.subheader("Total Memberships Comparison")
            st.plotly_chart(fig_totals, use_container_width=True)
        
        if options['monthly'] and fig_monthly:
            st.subheader("Monthly Performance Trends")
            st.plotly_chart(fig_monthly, use_container_width=True)
        
        if options['growth'] and fig_growth:
            st.subheader("Growth Rate Comparison")
            st.plotly_chart(fig_growth, use_container_width=True)
    
    # Membership distribution comparison
    if options['distribution']:
        st.subheader("Membership Type Distribution Comparison")
        fig_distribution = analyzer.create_membership_distribution_comparison()
        if fig_distribution:
            st.plotly_chart(fig_distribution, use_container_width=True)
        else:
            st.warning("Could not generate membership distribution comparison")
    
    # Summary table
    if options['table']:
        st.subheader("Detailed Statistics")
        
        # Format the display table
        display_df = comparison_df[['Sheet_Name', 'Total_Memberships', 'Active_Membership_Types', 
                                  'Best_Month', 'Best_Month_Count', 'Avg_Monthly', 'Growth_Rate']].copy()
        
        display_df.columns = ['Location & Year', 'Total Members', 'Active Types', 
                             'Best Month', 'Best Month Count', 'Avg Monthly', 'Growth Rate (%)']
        
        # Format numbers
        display_df['Total Members'] = display_df['Total Members'].apply(lambda x: f"{x:,.0f}")
        display_df['Best Month Count'] = display_df['Best Month Count'].apply(lambda x: f"{x:,.0f}")
        display_df['Avg Monthly'] = display_df['Avg Monthly'].apply(lambda x: f"{x:.1f}")
        display_df['Growth Rate (%)'] = display_df['Growth Rate (%)'].apply(lambda x: f"{x:.1f}%")
        
        st.dataframe(display_df, use_container_width=True)
    
    # Key insights for comparison
    st.subheader("Comparison Insights")
    
    insights = []
    
    # Best and worst performers
    best_idx = comparison_df['Total_Memberships'].idxmax()
    worst_idx = comparison_df['Total_Memberships'].idxmin()
    
    best_location = comparison_df.loc[best_idx, 'Sheet_Name']
    best_total = comparison_df.loc[best_idx, 'Total_Memberships']
    worst_location = comparison_df.loc[worst_idx, 'Sheet_Name']
    worst_total = comparison_df.loc[worst_idx, 'Total_Memberships']
    
    insights.append(f"**Best performing location:** {best_location} with {best_total:,.0f} total memberships")
    
    if len(comparison_df) > 1:
        insights.append(f"**Lowest performing location:** {worst_location} with {worst_total:,.0f} total memberships")
        
        # Performance gap
        performance_gap = ((best_total - worst_total) / worst_total) * 100
        insights.append(f"**Performance gap:** {performance_gap:.1f}% difference between best and worst")
    
    # Growth insights
    growing_locations = comparison_df[comparison_df['Growth_Rate'] > 5]
    declining_locations = comparison_df[comparison_df['Growth_Rate'] < -5]
    
    if len(growing_locations) > 0:
        insights.append(f"**Growing locations:** {len(growing_locations)} location(s) showing strong growth (>5%)")
    
    if len(declining_locations) > 0:
        insights.append(f"**Declining locations:** {len(declining_locations)} location(s) showing decline (<-5%)")
    
    # Year-over-year comparison if available
    locations_2024 = comparison_df[comparison_df['Year'] == '2024']
    locations_2025 = comparison_df[comparison_df['Year'] == '2025']
    
    if len(locations_2024) > 0 and len(locations_2025) > 0:
        avg_2024 = locations_2024['Total_Memberships'].mean()
        avg_2025 = locations_2025['Total_Memberships'].mean()
        yoy_change = ((avg_2025 - avg_2024) / avg_2024) * 100
        insights.append(f"**Year-over-year change:** {yoy_change:.1f}% change from 2024 to 2025 average")
    
    for insight in insights:
        st.markdown(f'<div class="insight-box">{insight}</div>', unsafe_allow_html=True)
    
    # Export comparison
    st.subheader("Export Comparison")
    if st.button("Export Comparison Data"):
        try:
            output_file = "location_comparison_analysis.xlsx"
            if analyzer.export_summary(output_file):
                st.success(f"Comparison data exported to {output_file}")
            else:
                st.error("Failed to export comparison data")
        except Exception as e:
            st.error(f"Export error: {str(e)}")

def display_two_location_comparison(analyzer, location_1, location_2, options):
    """Display side-by-side comparison of two locations"""
    
    st.header(f"Comparison: {location_1} vs {location_2}")
    
    # Overview metrics comparison
    st.subheader("Quick Comparison")
    
    # Get insights for both locations
    insights_1 = analyzer.generate_insights(location_1)
    insights_2 = analyzer.generate_insights(location_2)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown(f"### 📍 {location_1}")
        for insight in insights_1[:3]:  # Show top 3 insights
            st.markdown(f"- {insight}")
    
    with col2:
        st.markdown(f"### 📍 {location_2}")
        for insight in insights_2[:3]:  # Show top 3 insights
            st.markdown(f"- {insight}")
    
    # Side-by-side charts
    if options['trends']:
        st.subheader("Monthly Trends Comparison")
        col1, col2 = st.columns(2)
        
        with col1:
            fig_trends_1 = analyzer.create_monthly_trend_chart(location_1)
            if fig_trends_1:
                st.plotly_chart(fig_trends_1, use_container_width=True)
        
        with col2:
            fig_trends_2 = analyzer.create_monthly_trend_chart(location_2)
            if fig_trends_2:
                st.plotly_chart(fig_trends_2, use_container_width=True)
    
    if options['pie']:
        st.subheader("Membership Distribution Comparison")
        
        # First location pie chart
        st.markdown(f"#### {location_1}")
        fig_pie_1 = analyzer.create_membership_type_pie_chart(location_1)
        if fig_pie_1:
            st.plotly_chart(fig_pie_1, use_container_width=True)
        
        # Second location pie chart
        st.markdown(f"#### {location_2}")
        fig_pie_2 = analyzer.create_membership_type_pie_chart(location_2)
        if fig_pie_2:
            st.plotly_chart(fig_pie_2, use_container_width=True)
    
    if options['heatmap']:
        st.subheader("Activity Heatmaps Comparison")
        
        # First location heatmap
        st.markdown(f"#### {location_1}")
        fig_heatmap_1 = analyzer.create_heatmap(location_1)
        if fig_heatmap_1:
            st.plotly_chart(fig_heatmap_1, use_container_width=True)
        
        # Second location heatmap  
        st.markdown(f"#### {location_2}")
        fig_heatmap_2 = analyzer.create_heatmap(location_2)
        if fig_heatmap_2:
            st.plotly_chart(fig_heatmap_2, use_container_width=True)
    
    if options['top_performers']:
        st.subheader("Top Performers Comparison")
        
        # First location top performers
        st.markdown(f"#### {location_1}")
        fig_top_1 = analyzer.create_top_performers_chart(location_1)
        if fig_top_1:
            st.plotly_chart(fig_top_1, use_container_width=True)
        
        # Second location top performers
        st.markdown(f"#### {location_2}")
        fig_top_2 = analyzer.create_top_performers_chart(location_2)
        if fig_top_2:
            st.plotly_chart(fig_top_2, use_container_width=True)
    
    if options['insights']:
        st.subheader("Detailed Insights Comparison")
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown(f"### {location_1} - Complete Analysis")
            for insight in insights_1:
                st.markdown(f'<div class="insight-box">{insight}</div>', unsafe_allow_html=True)
        
        with col2:
            st.markdown(f"### {location_2} - Complete Analysis")
            for insight in insights_2:
                st.markdown(f'<div class="insight-box">{insight}</div>', unsafe_allow_html=True)
    
    # Winner analysis
    st.subheader("Head-to-Head Analysis")
    
    try:
        # Extract total memberships from insights
        total_1 = None
        total_2 = None
        
        for insight in insights_1:
            if "Total memberships:" in insight:
                total_1 = int(insight.split(":")[1].replace(",", "").strip())
                break
        
        for insight in insights_2:
            if "Total memberships:" in insight:
                total_2 = int(insight.split(":")[1].replace(",", "").strip())
                break
        
        if total_1 and total_2:
            winner = location_1 if total_1 > total_2 else location_2
            winner_total = max(total_1, total_2)
            loser_total = min(total_1, total_2)
            difference = winner_total - loser_total
            percentage_diff = (difference / loser_total) * 100
            
            st.success(f"**Winner: {winner}** with {winner_total:,} total memberships")
            st.info(f"**Performance Gap:** {difference:,} memberships ({percentage_diff:.1f}% more)")
            
            # Performance breakdown
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric(location_1, f"{total_1:,}", 
                         delta=f"{total_1 - total_2:,}" if total_1 > total_2 else f"{total_1 - total_2:,}")
            
            with col2:
                st.metric("vs", "")
            
            with col3:
                st.metric(location_2, f"{total_2:,}", 
                         delta=f"{total_2 - total_1:,}" if total_2 > total_1 else f"{total_2 - total_1:,}")
        
    except Exception as e:
        st.warning("Could not perform head-to-head analysis")
    
    # Export comparison
    st.subheader("Export Two-Location Comparison")
    if st.button("Export Comparison Data", key="two_location_export"):
        try:
            output_file = f"{location_1.replace(' ', '_')}_vs_{location_2.replace(' ', '_')}_comparison.xlsx"
            if analyzer.export_summary(output_file):
                st.success(f"Comparison data exported to {output_file}")
            else:
                st.error("Failed to export comparison data")
        except Exception as e:
            st.error(f"Export error: {str(e)}")

if __name__ == "__main__":
    main()
