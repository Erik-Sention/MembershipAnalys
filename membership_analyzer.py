import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import warnings
warnings.filterwarnings('ignore')

class MembershipAnalyzer:
    def __init__(self, excel_file_path, data_type='memberships'):
        """
        Initialize the analyzer with an Excel file containing multiple sheets
        data_type: 'memberships' or 'tests'
        """
        self.excel_file_path = excel_file_path
        self.data = {}
        self.combined_data = None
        self.data_type = data_type
        
        # Define the correct membership types that should be shown
        self.valid_membership_types = {
            'Standard',
            'Standard TRI/OCR/MULTI', 
            'Premium',
            'Premium TRI/OCR/MULTI',
            'Supreme',
            'Supreme TRI/OCR/MULTI',
            'Iform 4 mån',
            'Iform Tillägg till MS 4 mån',
            'Iform Extra månad',
            'Iform Fortsättning',
            'BAS',
            'Avslut',
            'Konvertering från test till membership'
        }
        
        # Define valid test types for 2024
        self.valid_test_types_2024 = {
            'Avancerat test Löpning',
            'Avancerat test Cykel',
            'Avancerat test Skidor',
            'Avancerat test Triatlon/Multisport',
            'Avancerat test OCR',
            'VO2max fristående',
            'VO2max tilläggstjänst',
            'Kostregistrering och kostrådgivning',
            'Wingate Fristående',
            'Wingate tilläggstjänst',
            'Kroppss fett% tillägg',
            'Kroppss fett% fristående',
            'Blodanalys',
            'Hb endast',
            'Glucos endast',
            'Blodfetter',
            'Teknikanalys Löpning Tilläggtjänst',
            'Teknikanalys Skidor Tilläggtjänst',
            'Teknikanalys Löpning Fristående',
            'Teknikanalys Skidor Fristående',
            'Funktionsanalys Fristående',
            'Teknikanalys med funktionsanalys',
            'Sen avbokning',
            'Sommardubbel',
            'Hälsopaket - Privatkund',
            'Natriumanalys'
        }
        
        # Define valid test types for 2025
        self.valid_test_types_2025 = {
            'Tröskeltest',
            'Tröskeltest + VO2max',
            'Tröskeltest Triathlon',
            'Tröskeltest Triathlon + VO2max',
            'VO2max fristående',
            'VO2max tillägg',
            'Wingate fristående',
            'Wingatetest tillägg',
            'Styrketest tillägg',
            'Teknikanalys tillägg',
            'Teknikanalys',
            'Funktionsanalys',
            'Funktions- och löpteknikanalys',
            'Hälsopaket',
            'Sommardubbel',
            'Personlig Träning 1 - Betald yta',
            'Personlig Träning 1 - Gratis yta',
            'Personlig Träning 5',
            'Personlig Träning 10',
            'Personlig Träning 20',
            'PT-Klipp - Betald yta',
            'PT-Klipp - Gratis yta',
            'Konvertering från test till PT20 - Till kollega',
            'Sen avbokning',
            'Genomgång eller testdel utförd av någon annan - Minus 30 min tid',
            'Genomgång eller testdel utförd till någon annan - Plus 30 min tid',
            'Natriumanalys (Svettest)'
        }
        
    def load_data(self):
        """
        Load all sheets from the Excel file
        """
        try:
            # Read all sheets from Excel file
            all_sheets = pd.read_excel(self.excel_file_path, sheet_name=None, engine='openpyxl')
            
            for sheet_name, df in all_sheets.items():
                # Clean sheet name and store
                clean_name = sheet_name.strip()
                self.data[clean_name] = df
                print(f"Loaded sheet: {clean_name} with shape {df.shape}")
                
            return True
        except Exception as e:
            print(f"Error loading Excel file: {e}")
            return False
    
    def clean_and_standardize_data(self):
        """
        Clean and standardize the data format (works for both memberships and tests)
        """
        cleaned_data = {}
        
        for sheet_name, df in self.data.items():
            try:
                # Make a copy to avoid modifying original
                clean_df = df.copy()
                
                # Extract location and year from sheet name
                parts = sheet_name.split()
                if len(parts) >= 2:
                    location = parts[0]
                    year = parts[1] if parts[1].isdigit() else '2024'
                else:
                    location = sheet_name
                    year = '2024'
                
                # Remove problematic columns that shouldn't be included in calculations
                columns_to_remove = []
                for col in clean_df.columns:
                    col_lower = str(col).lower()
                    # Remove total, percent, and summary columns
                    if any(keyword in col_lower for keyword in [
                        'total', 'summa', 'procent', '%', 'per år', 'per månad'
                    ]):
                        columns_to_remove.append(col)
                
                # Drop the problematic columns
                clean_df = clean_df.drop(columns=columns_to_remove, errors='ignore')
                
                # Determine the data column name based on data type
                data_col_name = 'Test' if self.data_type == 'tests' else 'Membership'
                
                # Remove rows that are summary rows (like "Summa per Månad")
                if data_col_name in clean_df.columns or len(clean_df.columns) > 0:
                    first_col = data_col_name if data_col_name in clean_df.columns else clean_df.columns[0]
                    
                    # Filter out summary rows
                    mask = ~clean_df[first_col].astype(str).str.lower().str.contains(
                        'summa|total|sum|per månad|per år', 
                        na=False, 
                        regex=True
                    )
                    clean_df = clean_df[mask].copy()
                
                # Rename first column to appropriate name if needed
                if len(clean_df.columns) > 0:
                    first_col = clean_df.columns[0]
                    if first_col != data_col_name and not first_col in ['Location', 'Year', 'Sheet_Name']:
                        clean_df = clean_df.rename(columns={first_col: data_col_name})
                
                # Filter to only valid types based on data type
                if data_col_name in clean_df.columns:
                    before_filter = len(clean_df)
                    
                    if self.data_type == 'tests':
                        # For tests, don't filter - show all tests that exist in the data
                        # This allows identical tests to be summed naturally without pre-filtering
                        valid_types = None  # No filtering for tests
                    else:
                        # Only filter memberships to the predefined list
                        valid_types = self.valid_membership_types
                        clean_df = clean_df[clean_df[data_col_name].isin(valid_types)].copy()
                        after_filter = len(clean_df)
                        if before_filter != after_filter:
                            print(f"🔍 Filtered {sheet_name}: {before_filter} → {after_filter} rows (removed {before_filter - after_filter} invalid membership types)")
                
                # Add metadata columns AFTER filtering
                clean_df['Location'] = location
                clean_df['Year'] = year
                clean_df['Sheet_Name'] = sheet_name
                
                # Store cleaned data if we have actual data
                if len(clean_df) > 0 and data_col_name in clean_df.columns:
                    cleaned_data[sheet_name] = clean_df
                    print(f"✅ Cleaned {sheet_name}: {len(clean_df)} rows, {len(clean_df.columns)} columns")
                    
            except Exception as e:
                print(f"Error cleaning sheet {sheet_name}: {e}")
                continue
        
        self.data = cleaned_data
        return len(cleaned_data) > 0
    
    def combine_data_from_all_sheets(self):
        """
        Combine data from all sheets into a single DataFrame for analysis
        """
        if not self.data:
            print("No data loaded. Please call load_data() and clean_and_standardize_data() first.")
            return None
        
        combined_dfs = []
        data_col_name = 'Test' if self.data_type == 'tests' else 'Membership'
        
        for sheet_name, df in self.data.items():
            if data_col_name in df.columns and len(df) > 0:
                combined_dfs.append(df.copy())
        
        if combined_dfs:
            self.combined_data = pd.concat(combined_dfs, ignore_index=True)
            print(f"✅ Combined data from {len(combined_dfs)} sheets: {len(self.combined_data)} total rows")
            
            # Show unique types in combined data
            unique_types = self.combined_data[data_col_name].dropna().astype(str).unique()
            unique_types = sorted([t for t in unique_types if t != 'nan'])
            type_name = "test types" if self.data_type == 'tests' else "membership types"
            print(f"📋 Valid {type_name} found: {len(unique_types)}")
            for t in unique_types:
                print(f"   - {t}")
            
            return self.combined_data
        else:
            print("No valid data found to combine")
            return None
    
    def create_monthly_trend_chart(self, sheet_name=None):
        """
        Create a monthly trend chart showing membership acquisitions over time
        """
        if sheet_name and sheet_name in self.data:
            df = self.data[sheet_name]
        else:
            # Use first available sheet
            df = list(self.data.values())[0]
            sheet_name = list(self.data.keys())[0]
        
        # Extract monthly columns (handling both 2024 and 2025 formats)
        month_columns = [col for col in df.columns if any(month in str(col).lower() for month in 
                        ['januari', 'februari', 'mars', 'april', 'maj', 'juni', 
                         'juli', 'augusti', 'september', 'oktober', 'november', 'december'])]
        
        # Sort month columns chronologically
        month_order = ['januari', 'februari', 'mars', 'april', 'maj', 'juni', 
                      'juli', 'augusti', 'september', 'oktober', 'november', 'december']
        
        def get_month_order(col_name):
            col_lower = col_name.lower()
            for i, month in enumerate(month_order):
                if month in col_lower:
                    return i
            return 999  # Unknown month goes to end
        
        month_columns = sorted(month_columns, key=get_month_order)
        
        if not month_columns:
            # Fallback to numeric columns that might be months
            month_columns = [col for col in df.columns if df[col].dtype in ['int64', 'float64']][:12]
        
        # Calculate total memberships per month
        monthly_totals = []
        month_labels = []
        
        for col in month_columns:
            # Convert to numeric, handling any non-numeric values
            numeric_col = pd.to_numeric(df[col], errors='coerce').fillna(0)
            total = numeric_col.sum()
            monthly_totals.append(total)
            
            # Clean up month label for display
            clean_label = col.replace(' 2024', '').replace(' 2025', '').title()
            month_labels.append(clean_label)
        
        # Create the plot
        fig = go.Figure()
        
        fig.add_trace(go.Scatter(
            x=month_labels,
            y=monthly_totals,
            mode='lines+markers',
            name='Total Memberships',
            line=dict(color='#1f77b4', width=3),
            marker=dict(size=8),
            hovertemplate='<b>%{x}</b><br>Memberships: %{y}<extra></extra>'
        ))
        
        fig.update_layout(
            title=f'Monthly Membership Trends - {sheet_name}',
            xaxis_title='Month',
            yaxis_title='Number of Memberships',
            template='plotly_white',
            height=500,
            xaxis={'tickangle': 45}
        )
        
        return fig
    
    def create_membership_type_pie_chart(self, sheet_name=None):
        """
        Create a pie chart showing distribution of membership/test types
        """
        if sheet_name and sheet_name in self.data:
            df = self.data[sheet_name]
        else:
            df = list(self.data.values())[0]
            sheet_name = list(self.data.keys())[0]
        
        # Find data column based on data type
        data_col_name = 'Test' if self.data_type == 'tests' else 'Membership'
        data_col = None
        
        for col in df.columns:
            if data_col_name.lower() in str(col).lower() or str(col).strip() == data_col_name:
                data_col = col
                break
        
        if not data_col and len(df.columns) > 0:
            data_col = df.columns[0]
        
        # Get month columns for calculating totals
        month_columns = [col for col in df.columns if any(month in str(col).lower() for month in 
                        ['januari', 'februari', 'mars', 'april', 'maj', 'juni', 
                         'juli', 'augusti', 'september', 'oktober', 'november', 'december'])]
        
        if data_col and month_columns:
            df_copy = df.copy()
            
            # Calculate totals from monthly data
            total_values = []
            type_names = []
            
            for idx, row in df_copy.iterrows():
                # Sum up monthly values for each type
                monthly_sum = 0
                for col in month_columns:
                    value = pd.to_numeric(row[col], errors='coerce')
                    if not pd.isna(value):
                        monthly_sum += value
                
                if monthly_sum > 0:  # Only include types with actual data
                    total_values.append(monthly_sum)
                    type_names.append(str(row[data_col]).strip())
            
            if total_values:
                # Create DataFrame for pie chart
                pie_data = pd.DataFrame({
                    data_col_name: type_names,
                    'Total': total_values
                })
                
                fig = px.pie(
                    pie_data, 
                    values='Total', 
                    names=data_col_name,
                    title=f'{data_col_name} Type Distribution - {sheet_name}',
                    color_discrete_sequence=px.colors.qualitative.Set3
                )
                
                fig.update_traces(
                    textposition='inside', 
                    textinfo='percent+label',
                    hovertemplate='<b>%{label}</b><br>Count: %{value}<br>Percentage: %{percent}<extra></extra>'
                )
                fig.update_layout(
                    height=500,  # Standard height
                    width=600,   # Max width to prevent fullscreen distortion
                    showlegend=True,
                    margin=dict(t=80, b=80, l=20, r=20)  # More bottom margin for legend
                )
                
                return fig
        
        return None
    
    def create_heatmap(self, sheet_name=None):
        """
        Create a heatmap showing membership activity by type and month
        """
        if sheet_name and sheet_name in self.data:
            df = self.data[sheet_name]
        else:
            df = list(self.data.values())[0]
            sheet_name = list(self.data.keys())[0]
        
        # Get month columns
        month_columns = [col for col in df.columns if any(month in str(col).lower() for month in 
                        ['januari', 'februari', 'mars', 'april', 'maj', 'juni', 
                         'juli', 'augusti', 'september', 'oktober', 'november', 'december'])]
        
        # Sort month columns chronologically
        month_order = ['januari', 'februari', 'mars', 'april', 'maj', 'juni', 
                      'juli', 'augusti', 'september', 'oktober', 'november', 'december']
        
        def get_month_order(col_name):
            col_lower = col_name.lower()
            for i, month in enumerate(month_order):
                if month in col_lower:
                    return i
            return 999
        
        month_columns = sorted(month_columns, key=get_month_order)
        
        # Find membership column
        membership_col = None
        for col in df.columns:
            if 'membership' in str(col).lower() or str(col).strip() == 'Membership':
                membership_col = col
                break
        
        if not membership_col and len(df.columns) > 0:
            membership_col = df.columns[0]
        
        if not month_columns or not membership_col:
            return None
        
        # Prepare data for heatmap
        df_copy = df.copy()
        
        # Convert month columns to numeric
        for col in month_columns:
            df_copy[col] = pd.to_numeric(df_copy[col], errors='coerce').fillna(0)
        
        # Filter out rows with all zeros
        row_sums = df_copy[month_columns].sum(axis=1)
        df_filtered = df_copy[row_sums > 0].copy()
        
        if len(df_filtered) == 0:
            return None
        
        # Set up heatmap data
        heatmap_data = df_filtered.set_index(membership_col)[month_columns]
        
        # Clean up column names for display
        clean_columns = [col.replace(' 2024', '').replace(' 2025', '').title() for col in month_columns]
        heatmap_data.columns = clean_columns
        
        # Create heatmap
        fig = px.imshow(
            heatmap_data,
            title=f'Membership Activity Heatmap - {sheet_name}',
            labels=dict(x="Month", y="Membership Type", color="Count"),
            aspect="auto",
            color_continuous_scale='Blues'
        )
        
        fig.update_layout(
            height=max(600, len(heatmap_data) * 25),  # Dynamic height based on data
            xaxis={'tickangle': 45}
        )
        
        return fig
    
    def create_top_performers_chart(self, sheet_name=None, top_n=10):
        """
        Create a bar chart showing top performing membership types
        """
        if sheet_name and sheet_name in self.data:
            df = self.data[sheet_name]
        else:
            df = list(self.data.values())[0]
            sheet_name = list(self.data.keys())[0]
        
        # Find membership column
        membership_col = None
        for col in df.columns:
            if 'membership' in str(col).lower() or str(col).strip() == 'Membership':
                membership_col = col
                break
        
        if not membership_col and len(df.columns) > 0:
            membership_col = df.columns[0]
        
        # Get month columns for calculating totals
        month_columns = [col for col in df.columns if any(month in str(col).lower() for month in 
                        ['januari', 'februari', 'mars', 'april', 'maj', 'juni', 
                         'juli', 'augusti', 'september', 'oktober', 'november', 'december'])]
        
        if membership_col and month_columns:
            df_copy = df.copy()
            
            # Calculate totals from monthly data
            total_values = []
            membership_names = []
            
            for idx, row in df_copy.iterrows():
                # Sum up monthly values for each membership type
                monthly_sum = 0
                for col in month_columns:
                    value = pd.to_numeric(row[col], errors='coerce')
                    if not pd.isna(value):
                        monthly_sum += value
                
                if monthly_sum > 0:  # Only include memberships with actual data
                    total_values.append(monthly_sum)
                    membership_names.append(str(row[membership_col]).strip())
            
            if total_values:
                # Create DataFrame for top performers
                performers_data = pd.DataFrame({
                    'Membership': membership_names,
                    'Total': total_values
                })
                
                # Sort and get top N
                performers_data = performers_data.sort_values('Total', ascending=False)
                top_performers = performers_data.head(min(top_n, len(performers_data)))
                
                fig = px.bar(
                    top_performers,
                    x='Membership',
                    y='Total',
                    title=f'Top {min(top_n, len(performers_data))} Membership Types - {sheet_name}',
                    color='Total',
                    color_continuous_scale='viridis',
                    text='Total'
                )
                
                fig.update_traces(texttemplate='%{text}', textposition='outside')
                
                fig.update_layout(
                    xaxis_title='Membership Type',
                    yaxis_title='Total Count',
                    xaxis={'tickangle': 45},
                    height=600,  # Normal height
                    showlegend=False,
                    margin=dict(t=60, b=120, l=60, r=60),  # Reasonable margins
                    yaxis=dict(range=[0, max(top_performers['Total']) * 1.15])  # Modest extra space for text
                )
                
                return fig
        
        return None
    
    def generate_insights(self, sheet_name=None):
        """
        Generate automated insights from the data
        """
        if sheet_name and sheet_name in self.data:
            df = self.data[sheet_name]
        else:
            df = list(self.data.values())[0]
            sheet_name = list(self.data.keys())[0]
        
        insights = []
        
        # Dynamic labels based on data type
        data_label = "tests" if self.data_type == 'tests' else "memberships"
        data_label_cap = "Tests" if self.data_type == 'tests' else "Memberships"
        type_label = "test" if self.data_type == 'tests' else "membership"
        type_label_cap = "Test" if self.data_type == 'tests' else "Membership"
        members_label = "tests performed" if self.data_type == 'tests' else "members"
        
        try:
            # Find data column based on data type
            data_col_name = 'Test' if self.data_type == 'tests' else 'Membership'
            data_col = None
            
            for col in df.columns:
                try:
                    col_str = str(col).lower()
                    if data_col_name.lower() in col_str or str(col).strip() == data_col_name:
                        data_col = col
                        break
                except:
                    continue
            
            if not data_col and len(df.columns) > 0:
                data_col = df.columns[0]
            
            # Get month columns (handle both string and numeric columns safely)
            month_columns = []
            for col in df.columns:
                try:
                    col_str = str(col).lower()
                    if any(month in col_str for month in 
                          ['januari', 'februari', 'mars', 'april', 'maj', 'juni', 
                           'juli', 'augusti', 'september', 'oktober', 'november', 'december']):
                        month_columns.append(col)
                except:
                    continue
            
            if data_col and month_columns:
                df_copy = df.copy()
                
                # Calculate totals from monthly data for each type
                type_totals = []
                type_names = []
                
                for idx, row in df_copy.iterrows():
                    monthly_sum = 0
                    for col in month_columns:
                        value = pd.to_numeric(row[col], errors='coerce')
                        if not pd.isna(value):
                            monthly_sum += value
                    
                    if monthly_sum > 0:
                        type_totals.append(monthly_sum)
                        type_names.append(str(row[data_col]).strip())
                
                if type_totals:
                    # Total count
                    total_count = sum(type_totals)
                    insights.append(f"📊 Total {data_label}: {total_count:,.0f}")
                    
                    # Top performer
                    max_idx = type_totals.index(max(type_totals))
                    top_type = type_names[max_idx]
                    top_count = type_totals[max_idx]
                    insights.append(f"🏆 Top performing {type_label}: {top_type} ({top_count:.0f} {members_label})")
                    
                    # Average per active type
                    avg_per_type = np.mean(type_totals)
                    insights.append(f"📈 Average {data_label} per active type: {avg_per_type:.1f}")
                
                # Monthly analysis
                monthly_totals = []
                for col in month_columns:
                    numeric_col = pd.to_numeric(df_copy[col], errors='coerce').fillna(0)
                    monthly_totals.append(numeric_col.sum())
                
                if monthly_totals:
                    # Best month
                    best_month_idx = np.argmax(monthly_totals)
                    best_month = month_columns[best_month_idx].replace(' 2024', '').replace(' 2025', '')
                    best_month_count = monthly_totals[best_month_idx]
                    insights.append(f"🗓️ Best performing month: {best_month} ({best_month_count:.0f} new {members_label})")
                    
                    # Worst month (only if there are multiple months with data)
                    non_zero_months = [i for i, total in enumerate(monthly_totals) if total > 0]
                    if len(non_zero_months) > 1:
                        worst_month_idx = np.argmin([monthly_totals[i] for i in non_zero_months])
                        actual_worst_idx = non_zero_months[worst_month_idx]
                        worst_month = month_columns[actual_worst_idx].replace(' 2024', '').replace(' 2025', '')
                        worst_month_count = monthly_totals[actual_worst_idx]
                        insights.append(f"📉 Lowest performing month: {worst_month} ({worst_month_count:.0f} new {members_label})")
                    
                    # Growth trend (if we have multiple months)
                    if len([x for x in monthly_totals if x > 0]) >= 3:
                        # Calculate simple trend
                        first_half = monthly_totals[:len(monthly_totals)//2]
                        second_half = monthly_totals[len(monthly_totals)//2:]
                        
                        first_avg = np.mean([x for x in first_half if x > 0]) if any(x > 0 for x in first_half) else 0
                        second_avg = np.mean([x for x in second_half if x > 0]) if any(x > 0 for x in second_half) else 0
                        
                        if first_avg > 0 and second_avg > 0:
                            growth_rate = ((second_avg - first_avg) / first_avg) * 100
                            if growth_rate > 5:
                                insights.append(f"📈 Growing trend: {growth_rate:.1f}% increase from first to second half")
                            elif growth_rate < -5:
                                insights.append(f"📉 Declining trend: {abs(growth_rate):.1f}% decrease from first to second half")
                            else:
                                insights.append(f"➡️ Stable trend: relatively consistent performance")
            
        except Exception as e:
            insights.append(f"⚠️ Error generating insights: {str(e)}")
        
        return insights if insights else ["📊 No insights available for this dataset"]
    
    def get_available_sheets(self):
        """
        Get list of available sheet names
        """
        return list(self.data.keys())
    
    def create_location_comparison_summary(self):
        """
        Create a summary comparison of all locations/sheets
        """
        comparison_data = []
        
        for sheet_name, df in self.data.items():
            try:
                # Extract location and year
                parts = sheet_name.split()
                location = parts[0] if len(parts) > 0 else sheet_name
                year = parts[1] if len(parts) > 1 and parts[1].isdigit() else 'Unknown'
                
                # Find data column based on data type
                data_col_name = 'Test' if self.data_type == 'tests' else 'Membership'
                data_col = None
                
                for col in df.columns:
                    if data_col_name.lower() in str(col).lower() or str(col).strip() == data_col_name:
                        data_col = col
                        break
                
                if not data_col and len(df.columns) > 0:
                    data_col = df.columns[0]
                
                # Get month columns (handle both string and numeric columns safely)
                month_columns = []
                for col in df.columns:
                    try:
                        col_str = str(col).lower()
                        if any(month in col_str for month in 
                              ['januari', 'februari', 'mars', 'april', 'maj', 'juni', 
                               'juli', 'augusti', 'september', 'oktober', 'november', 'december']):
                            month_columns.append(col)
                    except:
                        continue
                
                if data_col and month_columns:
                    # Calculate totals
                    total_count = 0
                    monthly_totals = []
                    active_types = 0
                    
                    # Calculate type totals
                    for idx, row in df.iterrows():
                        monthly_sum = 0
                        for col in month_columns:
                            value = pd.to_numeric(row[col], errors='coerce')
                            if not pd.isna(value):
                                monthly_sum += value
                        
                        if monthly_sum > 0:
                            total_count += monthly_sum
                            active_types += 1
                    
                    # Calculate monthly totals
                    for col in month_columns:
                        numeric_col = pd.to_numeric(df[col], errors='coerce').fillna(0)
                        monthly_totals.append(numeric_col.sum())
                    
                    # Find best and worst months
                    if monthly_totals:
                        best_month_idx = np.argmax(monthly_totals)
                        best_month = month_columns[best_month_idx].replace(' 2024', '').replace(' 2025', '')
                        best_month_count = monthly_totals[best_month_idx]
                        
                        # Calculate average
                        non_zero_months = [x for x in monthly_totals if x > 0]
                        avg_monthly = np.mean(non_zero_months) if non_zero_months else 0
                        
                        # Growth trend
                        if len(non_zero_months) >= 3:
                            first_half = monthly_totals[:len(monthly_totals)//2]
                            second_half = monthly_totals[len(monthly_totals)//2:]
                            
                            first_avg = np.mean([x for x in first_half if x > 0]) if any(x > 0 for x in first_half) else 0
                            second_avg = np.mean([x for x in second_half if x > 0]) if any(x > 0 for x in second_half) else 0
                            
                            growth_rate = ((second_avg - first_avg) / first_avg) * 100 if first_avg > 0 else 0
                        else:
                            growth_rate = 0
                        
                        data_label = "Tests" if self.data_type == 'tests' else "Memberships"
                        type_label = "Test Types" if self.data_type == 'tests' else "Membership Types"
                        
                        comparison_data.append({
                            'Location': location,
                            'Year': year,
                            'Sheet_Name': sheet_name,
                            f'Total_{data_label}': total_count,
                            f'Active_{type_label.replace(" ", "_")}': active_types,
                            'Best_Month': best_month,
                            'Best_Month_Count': best_month_count,
                            'Avg_Monthly': avg_monthly,
                            'Growth_Rate': growth_rate,
                            'Monthly_Totals': monthly_totals,
                            'Month_Names': [col.replace(' 2024', '').replace(' 2025', '') for col in month_columns]
                        })
                        
            except Exception as e:
                print(f"Error processing {sheet_name} for comparison: {e}")
                continue
        
        return pd.DataFrame(comparison_data)
    
    def create_location_comparison_charts(self):
        """
        Create comparison charts across all locations
        """
        comparison_df = self.create_location_comparison_summary()
        
        if comparison_df.empty:
            return None, None, None
        
        # Dynamic column names and labels
        data_label = "Tests" if self.data_type == 'tests' else "Memberships"
        total_col_name = f'Total_{data_label}'
        
        # 1. Total comparison
        fig_totals = px.bar(
            comparison_df.sort_values(total_col_name, ascending=False),
            x='Sheet_Name',
            y=total_col_name,
            color='Location',
            title=f'Total {data_label} by Location & Year',
            text=total_col_name
        )
        fig_totals.update_traces(texttemplate='%{text}', textposition='outside')
        fig_totals.update_layout(
            xaxis_title='Location & Year',
            yaxis_title=f'Total {data_label}',
            xaxis={'tickangle': 45},
            height=500,
            margin=dict(t=60, b=80, l=60, r=60),  # Reasonable margins
            yaxis=dict(range=[0, comparison_df[total_col_name].max() * 1.15])  # Modest extra space
        )
        
        # 2. Monthly performance comparison
        monthly_data = []
        for idx, row in comparison_df.iterrows():
            for i, (month, total) in enumerate(zip(row['Month_Names'], row['Monthly_Totals'])):
                monthly_data.append({
                    'Location': row['Sheet_Name'],
                    'Month': month,
                    'Month_Order': i,
                    'Total': total
                })
        
        monthly_df = pd.DataFrame(monthly_data)
        
        fig_monthly = px.line(
            monthly_df,
            x='Month_Order',
            y='Total',
            color='Location',
            title='Monthly Performance Comparison Across Locations',
            markers=True
        )
        
        # Update x-axis to show month names
        if not monthly_df.empty:
            month_names = monthly_df[monthly_df['Location'] == monthly_df['Location'].iloc[0]]['Month'].tolist()
            fig_monthly.update_xaxes(
                tickmode='array',
                tickvals=list(range(len(month_names))),
                ticktext=month_names,
                tickangle=45
            )
        
        fig_monthly.update_layout(
            xaxis_title='Month',
            yaxis_title='New Memberships',
            height=500
        )
        
        # 3. Growth rate comparison
        fig_growth = px.bar(
            comparison_df.sort_values('Growth_Rate', ascending=False),
            x='Sheet_Name',
            y='Growth_Rate',
            color='Growth_Rate',
            color_continuous_scale=['red', 'yellow', 'green'],
            title='Growth Rate Comparison (First Half vs Second Half)',
            text='Growth_Rate'
        )
        fig_growth.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
        fig_growth.update_layout(
            xaxis_title='Location & Year',
            yaxis_title='Growth Rate (%)',
            xaxis={'tickangle': 45},
            height=500,
            margin=dict(t=60, b=80, l=60, r=60),  # Reasonable margins
            yaxis=dict(range=[min(comparison_df['Growth_Rate'].min() * 1.1, -10), 
                             max(comparison_df['Growth_Rate'].max() * 1.1, 10)])  # Modest extra space
        )
        
        return fig_totals, fig_monthly, fig_growth
    
    def create_membership_distribution_comparison(self):
        """
        Create comparison of type distributions across locations
        """
        comparison_data = []
        
        for sheet_name, df in self.data.items():
            try:
                # Find data column based on data type
                data_col_name = 'Test' if self.data_type == 'tests' else 'Membership'
                data_col = None
                
                for col in df.columns:
                    if data_col_name.lower() in str(col).lower() or str(col).strip() == data_col_name:
                        data_col = col
                        break
                
                if not data_col and len(df.columns) > 0:
                    data_col = df.columns[0]
                
                # Get month columns for calculating totals (handle both string and numeric columns safely)
                month_columns = []
                for col in df.columns:
                    try:
                        col_str = str(col).lower()
                        if any(month in col_str for month in 
                              ['januari', 'februari', 'mars', 'april', 'maj', 'juni', 
                               'juli', 'augusti', 'september', 'oktober', 'november', 'december']):
                            month_columns.append(col)
                    except:
                        continue
                
                if data_col and month_columns:
                    for idx, row in df.iterrows():
                        # Sum up monthly values for each type
                        monthly_sum = 0
                        for col in month_columns:
                            value = pd.to_numeric(row[col], errors='coerce')
                            if not pd.isna(value):
                                monthly_sum += value
                        
                        if monthly_sum > 0:  # Only include types with actual data
                            type_name = str(row[data_col]).strip()
                            comparison_data.append({
                                'Location': sheet_name,
                                'Type': type_name,
                                'Total': monthly_sum
                            })
                            
            except Exception as e:
                print(f"Error processing {sheet_name} for membership distribution: {e}")
                continue
        
        if not comparison_data:
            return None
        
        comparison_df = pd.DataFrame(comparison_data)
        
        # Create stacked bar chart showing type distribution by location
        data_label = "Tests" if self.data_type == 'tests' else "Members"
        type_label = "Test Type" if self.data_type == 'tests' else "Membership Type"
        
        fig = px.bar(
            comparison_df,
            x='Location',
            y='Total',
            color='Type',
            title=f'{type_label} Distribution Across Locations',
            labels={'Total': f'Number of {data_label}', 'Location': 'Location & Year'}
        )
        
        fig.update_layout(
            xaxis_title='Location & Year',
            yaxis_title=f'Number of {data_label}',
            xaxis={'tickangle': 45},
            height=600,  # Normal height
            legend_title=f'{type_label}s',
            margin=dict(t=60, b=80, l=60, r=60)  # Reasonable margins
        )
        
        return fig
    
    def create_membership_trends_by_type(self, sheet_name=None):
        """
        Create a line chart showing monthly trends for each type separately
        """
        if sheet_name and sheet_name in self.data:
            df = self.data[sheet_name]
        else:
            df = list(self.data.values())[0]
            sheet_name = list(self.data.keys())[0]
        
        # Find data column based on data type
        data_col_name = 'Test' if self.data_type == 'tests' else 'Membership'
        data_col = None
        
        for col in df.columns:
            if data_col_name.lower() in str(col).lower() or str(col).strip() == data_col_name:
                data_col = col
                break
        
        if not data_col and len(df.columns) > 0:
            data_col = df.columns[0]
        
        # Get month columns
        month_columns = [col for col in df.columns if any(month in str(col).lower() for month in 
                        ['januari', 'februari', 'mars', 'april', 'maj', 'juni', 
                         'juli', 'augusti', 'september', 'oktober', 'november', 'december'])]
        
        if not data_col or not month_columns:
            return None
        
        # Prepare data for multi-line chart
        trend_data = []
        
        for _, row in df.iterrows():
            type_name = str(row[data_col]).strip()
            
            # Skip summary rows
            if any(skip_word in type_name.lower() for skip_word in 
                  ['summa', 'total', 'sum', 'totalt']):
                continue
            
            # Add data for each month (including 0 values)
            for month in month_columns:
                try:
                    value = pd.to_numeric(row[month], errors='coerce')
                    if pd.isna(value):
                        value = 0  # Treat NaN as 0
                    
                    # Extract month name for better display
                    month_name = month.split()[0] if ' ' in month else month
                    
                    trend_data.append({
                        'Month': month_name,
                        'Type': type_name,
                        'Count': value,
                        'Month_Order': month_columns.index(month)
                    })
                except:
                    continue
        
        if not trend_data:
            return None
        
        trend_df = pd.DataFrame(trend_data)
        
        # Filter out types that have ALL zero values (no activity at all)
        # But keep types that have at least some non-zero values
        type_activity = trend_df.groupby('Type')['Count'].sum()
        types_with_activity = type_activity[type_activity > 0].index
        trend_df = trend_df[trend_df['Type'].isin(types_with_activity)]
        
        if len(trend_df) == 0:
            return None
        
        # Sort by month order for proper line progression
        trend_df = trend_df.sort_values(['Type', 'Month_Order'])
        
        # Ensure proper month ordering for the x-axis
        month_order = ['Januari', 'Februari', 'Mars', 'April', 'Maj', 'Juni', 
                      'Juli', 'Augusti', 'September', 'Oktober', 'November', 'December']
        
        # Create a proper month category
        trend_df['Month_Cat'] = pd.Categorical(trend_df['Month'], categories=month_order, ordered=True)
        
        # Create multi-line chart with custom colors (darker yellow)
        custom_colors = [
            '#1f77b4',  # blue
            '#ff7f0e',  # orange  
            '#d62728',  # red
            '#2ca02c',  # green
            '#9467bd',  # purple
            '#8c564b',  # brown
            '#e377c2',  # pink
            '#7f7f7f',  # gray
            '#bcbd22',  # darker olive/yellow
            '#17becf',  # cyan
            '#aec7e8',  # light blue
            '#ffbb78',  # light orange
            '#98df8a',  # light green
            '#ff9896',  # light red
            '#c5b0d5',  # light purple
            '#c49c94'   # light brown
        ]
        
        # Create dynamic labels based on data type
        data_label = "Tests" if self.data_type == 'tests' else "Memberships"
        count_label = f"New {data_label}" if self.data_type == 'tests' else "New Memberships"
        
        fig = px.line(
            trend_df,
            x='Month_Cat',
            y='Count',
            color='Type',
            title=f'Monthly {data_label} Trends by Type - {sheet_name}',
            labels={'Count': count_label, 'Month_Cat': 'Month'},
            color_discrete_sequence=custom_colors,
            category_orders={'Month_Cat': month_order}
        )
        
        fig.update_traces(
            mode='lines+markers', 
            line=dict(width=2), 
            marker=dict(size=6),
            hovertemplate='<b>%{fullData.name}</b><br>' +
                         'Month: %{x}<br>' +
                         'New Memberships: %{y}<br>' +
                         '<extra></extra>'  # Remove trace box
        )
        
        fig.update_layout(
            xaxis_title='Month',
            yaxis_title=count_label,
            height=600,
            legend_title=f'{data_label} Types',
            margin=dict(t=80, b=80, l=60, r=60),
            xaxis=dict(
                type='category',  # Treat x-axis as categorical
                categoryorder='array',
                categoryarray=month_order
            ),
            legend=dict(
                orientation="v",
                yanchor="top",
                y=1,
                xanchor="left",
                x=1.02
            ),
            hovermode='closest'  # Show only the closest point/line
        )
        
        # Rotate x-axis labels for better readability
        fig.update_xaxes(tickangle=45)
        
        return fig
    
    def create_two_location_distribution_comparison(self, location_1, location_2):
        """
        Create a stacked bar chart comparing distribution between two locations
        """
        comparison_data = []
        
        # Dynamic column name and labels
        data_col_name = 'Test' if self.data_type == 'tests' else 'Membership'
        data_label = "Tests" if self.data_type == 'tests' else "Memberships"
        
        for sheet_name in [location_1, location_2]:
            if sheet_name not in self.data:
                continue
                
            df = self.data[sheet_name]
            
            # Find data column
            data_col = None
            for col in df.columns:
                try:
                    col_str = str(col).lower()
                    if data_col_name.lower() in col_str or str(col).strip() == data_col_name:
                        data_col = col
                        break
                except:
                    continue
            
            if not data_col and len(df.columns) > 0:
                data_col = df.columns[0]
            
            # Get month columns for calculating totals
            month_columns = [col for col in df.columns if any(month in str(col).lower() for month in 
                            ['januari', 'februari', 'mars', 'april', 'maj', 'juni', 
                             'juli', 'augusti', 'september', 'oktober', 'november', 'december'])]
            
            if data_col and month_columns:
                for _, row in df.iterrows():
                    data_type_name = str(row[data_col]).strip()
                    
                    # Skip summary rows
                    if any(skip_word in data_type_name.lower() for skip_word in 
                          ['summa', 'total', 'sum', 'totalt']):
                        continue
                    
                    # Calculate total from monthly data
                    total = 0
                    for month in month_columns:
                        try:
                            value = pd.to_numeric(row[month], errors='coerce')
                            if pd.notna(value):
                                total += value
                        except:
                            continue
                    
                    if total > 0:
                        comparison_data.append({
                            'Location': sheet_name,
                            'Type': data_type_name,
                            'Total': total
                        })
        
        if not comparison_data:
            return None
        
        comparison_df = pd.DataFrame(comparison_data)
        
        # Create stacked bar chart showing distribution by location
        type_label = "Test Types" if self.data_type == 'tests' else "Membership Types"
        count_label = f"Number of {data_label}"
        
        fig = px.bar(
            comparison_df,
            x='Location',
            y='Total',
            color='Type',
            title=f'{type_label} Distribution: {location_1} vs {location_2}',
            labels={'Total': count_label, 'Location': 'Location & Year'},
            color_discrete_sequence=px.colors.qualitative.Set3
        )
        
        fig.update_layout(
            xaxis_title='Location & Year',
            yaxis_title=count_label,
            height=600,
            legend_title=type_label,
            margin=dict(t=80, b=80, l=60, r=60),
            legend=dict(
                orientation="v",
                yanchor="top",
                y=1,
                xanchor="left", 
                x=1.02
            )
        )
        
        return fig
    
    def export_summary(self, output_file='membership_summary.xlsx'):
        """
        Export a summary of all sheets to Excel
        """
        try:
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                for sheet_name, df in self.data.items():
                    df.to_excel(writer, sheet_name=sheet_name[:31], index=False)  # Excel sheet names max 31 chars
                
                # Add comparison summary
                comparison_df = self.create_location_comparison_summary()
                if not comparison_df.empty:
                    comparison_df.to_excel(writer, sheet_name='Location_Comparison', index=False)
            
            print(f"Summary exported to {output_file}")
            return True
        except Exception as e:
            print(f"Error exporting summary: {e}")
            return False
