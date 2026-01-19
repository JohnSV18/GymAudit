"""
Gym Membership Audit Tool - Main Streamlit Application
A frictionless tool for auditing gym membership data
Supports multiple membership types and locations
"""

import streamlit as st
import json
import pandas as pd
from pathlib import Path
import plotly.express as px
import plotly.graph_objects as go

from core.red_flags import get_locations, get_membership_types, load_config
from core.audit_engine import AuditEngine
from utils.statistics import AuditStatistics


# Page configuration
st.set_page_config(
    page_title="Gym Membership Audit Tool",
    page_icon="🏋️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Load settings
@st.cache_resource
def load_settings():
    """Load application settings"""
    settings_path = Path("config/settings.json")
    if settings_path.exists():
        with open(settings_path, 'r') as f:
            return json.load(f)
    return {}


@st.cache_resource
def load_audit_config():
    """Load the red flag configuration"""
    return load_config()


def get_audit_engine(membership_type: str, location: str):
    """Create audit engine for the specified membership type and location"""
    return AuditEngine(membership_type, location, output_folder='outputs')


def format_currency(amount):
    """Format amount as currency"""
    return f"${amount:,.2f}"


def display_metric_card(label, value, help_text=None, color=None):
    """Display a metric in a card-like format"""
    if color:
        st.markdown(
            f"""
            <div style="padding: 1rem; border-radius: 0.5rem; background-color: {color}20; border-left: 4px solid {color};">
                <h4 style="margin: 0; color: {color};">{label}</h4>
                <h2 style="margin: 0.5rem 0 0 0;">{value}</h2>
            </div>
            """,
            unsafe_allow_html=True
        )
    else:
        st.metric(label=label, value=value, help=help_text)


def main():
    """Main application"""

    settings = load_settings()
    config = load_audit_config()

    # Header
    st.title("🏋️ Gym Membership Audit Tool")
    st.markdown("### Automated Red Flag Detection for Membership Data")

    # Help/Info Section
    with st.expander("ℹ️  **What am I looking at? Click here for help**", expanded=False):
        st.markdown("""
        ### 📊 Understanding Your Results

        ####  **Total Files**
        How many spreadsheets you uploaded for audit.

        #### **Total Records**
        The number of membership accounts that were checked in all your files.

        #### **⚠️ Flagged**
        Accounts that have one or more red flags (data problems). These accounts need review!
        - **Yellow highlighting** = Problem found
        - The percentage shows what portion of your data has issues

        #### **Financial Impact**
        **This is the total dollar value of data discrepancies found.**

        It includes:
        - **Missing Dues**: Money that should have been collected but wasn't
          - Example: Member paid $0 but should have paid the expected amount → Impact = expected amount
        - **Outstanding Balances**: Amounts owed (debits) or refunds due (credits)
          - Example: Member has $50 balance → Impact = $50

        **Important:** Some accounts may contribute to BOTH categories.

        The Financial Impact is the sum of all these discrepancies. It represents potential revenue issues,
        not necessarily actual money owed.

        ---

        ### 🚨 What are Red Flags?

        Red flags are data anomalies that indicate something might be wrong with an account.
        The specific checks depend on the membership type selected:

        **Paid in Full Memberships (1 Year, 3 Month, 1 Month):**
        - Date range validation based on membership length
        - Dues amount must be at least 90% of expected price
        - Cycle value must match expected (0 for PIF memberships)
        - Balance must be $0.00

        **Month to Month:**
        - Expiration year should be 2099
        - Cycle should be 1
        - Draft date within 3 months of join date
        - End draft year should be 2099
        - Balance must be $0.00

        **Yellow rows in your Excel reports = accounts that need attention!**
        """)

    st.markdown("---")

    # Sidebar
    with st.sidebar:
        st.header("About")
        st.info(
            "This tool automatically audits gym membership spreadsheets and flags "
            "accounts with data anomalies.\n\n"
            "**Select your location and membership type, then upload files!**"
        )

        # Location and Membership Type Selection
        st.header("Audit Settings")

        # Get available options from config
        locations = get_locations(config)
        membership_types = get_membership_types(config)

        # Location dropdown
        location_keys = list(locations.keys())
        location_labels = list(locations.values())
        selected_location_idx = st.selectbox(
            "Select Location",
            range(len(location_keys)),
            format_func=lambda x: location_labels[x],
            help="Choose the gym location for this audit"
        )
        selected_location = location_keys[selected_location_idx]

        # Membership type dropdown
        type_keys = list(membership_types.keys())
        type_labels = list(membership_types.values())
        selected_type_idx = st.selectbox(
            "Select Membership Type",
            range(len(type_keys)),
            format_func=lambda x: type_labels[x],
            help="Choose the membership type to audit"
        )
        selected_membership_type = type_keys[selected_type_idx]

        # Store selections in session state
        st.session_state['selected_location'] = selected_location
        st.session_state['selected_membership_type'] = selected_membership_type

        st.markdown("---")

        # Display current rules based on selection
        st.header("Red Flag Criteria")
        type_config = config['membership_types'].get(selected_membership_type, {})
        type_name = type_config.get('name', selected_membership_type)
        pricing = type_config.get('pricing', {})
        rules = type_config.get('rules', {})
        expected_dues = pricing.get(selected_location, 0) or 0

        st.markdown(f"**Rules for {type_name} at {locations[selected_location]}:**")

        # Show relevant rules based on membership type
        if rules.get('date_rule_type') == 'exact_range':
            min_days = rules.get('date_diff_min_days', 365)
            max_days = rules.get('date_diff_max_days', 366)
            st.markdown(f"- ⚠️ Date range: {min_days}-{max_days} days")
        elif rules.get('date_rule_type') == 'max_only':
            max_days = rules.get('date_diff_max_days', 31)
            st.markdown(f"- ⚠️ Max date range: {max_days} days")

        if rules.get('expected_exp_year'):
            st.markdown(f"- ⚠️ Exp year = {rules['expected_exp_year']}")

        if expected_dues > 0:
            threshold = rules.get('payment_threshold_percent', 90)
            min_dues = expected_dues * (threshold / 100)
            st.markdown(f"- ⚠️ Dues ≥ ${min_dues:.2f} ({threshold}% of ${expected_dues})")

        if rules.get('cycle_rule_type') == 'exact' and rules.get('expected_cycle') is not None:
            st.markdown(f"- ⚠️ Cycle = {rules['expected_cycle']}")
        elif rules.get('cycle_rule_type') == 'max' and rules.get('cycle_max') is not None:
            st.markdown(f"- ⚠️ Cycle ≤ {rules['cycle_max']}")

        if rules.get('check_balance', True):
            st.markdown(f"- ⚠️ Balance = ${rules.get('expected_balance', 0):.2f}")

        if rules.get('draft_date_max_months_from_join'):
            st.markdown(f"- ⚠️ Draft date within {rules['draft_date_max_months_from_join']} months of join")

        if rules.get('expected_end_draft_year'):
            st.markdown(f"- ⚠️ End draft year = {rules['expected_end_draft_year']}")

        st.markdown("---")
        st.caption(f"Version {settings.get('app', {}).get('version', '1.0.0')}")

    # Main content area
    tab1, tab2, tab3 = st.tabs(["📁 Upload & Audit", "📊 Pattern Analysis", "💰 Financial Summary"])

    # TAB 1: Upload & Audit
    with tab1:
        st.header("Upload Membership Files")

        # Show current selection
        current_location = st.session_state.get('selected_location', 'bqe')
        current_type = st.session_state.get('selected_membership_type', '1_year_paid_in_full')
        location_name = locations.get(current_location, current_location)
        type_name = membership_types.get(current_type, current_type)

        st.info(f"**Auditing:** {type_name} memberships at **{location_name}**  \n*Change settings in the sidebar if needed.*")

        # File uploader
        uploaded_files = st.file_uploader(
            "Choose CSV or Excel files",
            type=['csv', 'xlsx', 'xls'],
            accept_multiple_files=True,
            help="Upload one or more membership spreadsheets for audit"
        )

        if uploaded_files:
            st.success(f"✅ {len(uploaded_files)} file(s) uploaded")

            # Display uploaded files
            with st.expander("View uploaded files"):
                for idx, file in enumerate(uploaded_files, 1):
                    st.text(f"{idx}. {file.name}")

            # Process button
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                process_button = st.button(
                    "🚀 Process Files",
                    type="primary",
                    use_container_width=True
                )

            if process_button:
                # Progress bar
                progress_bar = st.progress(0)
                status_text = st.empty()

                # Get audit engine with current settings
                engine = get_audit_engine(current_type, current_location)

                # Process files
                status_text.text(f"Processing files for {type_name} at {location_name}...")
                results = engine.audit_multiple_uploaded_files(
                    uploaded_files,
                    generate_individual_reports=True,
                    generate_consolidated=len(uploaded_files) > 1
                )

                # Store the settings used for this audit
                results['audit_settings'] = {
                    'location': current_location,
                    'location_name': location_name,
                    'membership_type': current_type,
                    'membership_type_name': type_name
                }

                progress_bar.progress(100)
                status_text.empty()

                # Store results in session state
                st.session_state['audit_results'] = results

                # Display results
                st.markdown("---")
                st.header("📋 Audit Results")

                # Overall summary
                st.subheader("Overall Summary")
                col1, col2, col3, col4 = st.columns(4)

                with col1:
                    display_metric_card(
                        "Total Files",
                        results['successful_files'],
                        color="#4CAF50"
                    )

                with col2:
                    display_metric_card(
                        "Total Records",
                        f"{results['total_records']:,}",
                        color="#2196F3"
                    )

                with col3:
                    flagged_pct = (results['total_flagged'] / results['total_records'] * 100) if results['total_records'] > 0 else 0
                    display_metric_card(
                        "⚠️ Flagged",
                        f"{results['total_flagged']:,} ({flagged_pct:.1f}%)",
                        color="#FF9800"
                    )

                with col4:
                    display_metric_card(
                        "Financial Impact",
                        format_currency(results['total_financial_impact']),
                        color="#F44336"
                    )

                # Financial Impact Breakdown
                with st.expander("💡 **See Financial Impact Breakdown**", expanded=False):
                    st.markdown("### Where does the Financial Impact come from?")

                    # Get audit settings
                    audit_settings = results.get('audit_settings', {})
                    audit_type_name = audit_settings.get('membership_type_name', 'the selected membership type')
                    audit_location_name = audit_settings.get('location_name', 'the selected location')

                    total_dues = sum(r.get('total_dues_impact', 0) for r in results['file_results'] if r['success'])
                    total_balance = sum(r.get('total_balance_impact', 0) for r in results['file_results'] if r['success'])

                    col1, col2 = st.columns(2)

                    with col1:
                        st.metric(
                            "📉 Missing/Low Dues",
                            format_currency(total_dues),
                            help=f"Accounts that paid less than 90% of expected dues for {audit_type_name} at {audit_location_name}"
                        )
                        st.caption(f"{(total_dues / results['total_financial_impact'] * 100):.1f}% of total impact" if results['total_financial_impact'] > 0 else "")

                    with col2:
                        st.metric(
                            "💳 Outstanding Balances",
                            format_currency(total_balance),
                            help="Accounts with non-zero balances (debits or credits)"
                        )
                        st.caption(f"{(total_balance / results['total_financial_impact'] * 100):.1f}% of total impact" if results['total_financial_impact'] > 0 else "")

                    st.info(
                        "**Note:** Some accounts may contribute to both categories. For example, "
                        "an account with $0 dues AND a $50 balance would add to both totals. "
                        "This represents data discrepancy value, not necessarily actual money owed."
                    )

                st.markdown("---")

                # Per-file results
                st.subheader("Individual File Results")

                for file_result in results['file_results']:
                    if not file_result['success']:
                        st.error(f"❌ {file_result['filename']}: {file_result['error']}")
                        continue

                    with st.expander(f"📄 {file_result['filename']}", expanded=True):
                        # Metrics
                        col1, col2, col3 = st.columns(3)

                        with col1:
                            st.metric("Total Records", f"{file_result['total_records']:,}")

                        with col2:
                            flagged_color = "🔴" if file_result['flagged_count'] > 0 else "🟢"
                            st.metric(
                                f"{flagged_color} Flagged",
                                f"{file_result['flagged_count']:,}",
                                f"{file_result['flagged_percentage']:.1f}%"
                            )

                        with col3:
                            st.metric(
                                "Financial Impact",
                                format_currency(file_result['total_financial_impact'])
                            )

                        # Flagged Member IDs
                        if file_result['flagged_member_ids']:
                            st.markdown("**⚠️ Flagged Member IDs:**")

                            # Limit display
                            max_display = 50
                            member_ids = file_result['flagged_member_ids'][:max_display]
                            remaining = len(file_result['flagged_member_ids']) - max_display

                            # Display as chips/badges
                            ids_html = " ".join([
                                f'<span style="background-color: #FFF3CD; color: #856404; padding: 0.25rem 0.5rem; border-radius: 0.25rem; margin: 0.25rem; display: inline-block;">{mid}</span>'
                                for mid in member_ids
                            ])
                            st.markdown(ids_html, unsafe_allow_html=True)

                            if remaining > 0:
                                st.caption(f"... and {remaining} more")

                        # Download button
                        if file_result.get('report_path'):
                            report_path = Path(file_result['report_path'])
                            if report_path.exists():
                                with open(report_path, 'rb') as f:
                                    st.download_button(
                                        label=f"📥 Download Audit Report",
                                        data=f,
                                        file_name=report_path.name,
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                        use_container_width=True
                                    )

                # Consolidated report download
                if results.get('consolidated_report_path') and len(uploaded_files) > 1:
                    st.markdown("---")
                    st.subheader("📊 Consolidated Report")

                    consolidated_path = Path(results['consolidated_report_path'])
                    if consolidated_path.exists():
                        with open(consolidated_path, 'rb') as f:
                            st.download_button(
                                label="📥 Download Consolidated Audit Report",
                                data=f,
                                file_name=consolidated_path.name,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="primary",
                                use_container_width=True
                            )

        else:
            # Instructions when no files uploaded
            st.info("👆 Upload one or more membership spreadsheets to get started")

            st.markdown("### How to Use")
            st.markdown("""
            1. **Select Location**: Choose your gym location in the sidebar
            2. **Select Membership Type**: Choose the membership type you're auditing
            3. **Upload Files**: Click the file uploader above and select your CSV or Excel files
            4. **Process**: Click the "Process Files" button
            5. **Review Results**: See summary statistics and flagged member IDs
            6. **Download Reports**: Get highlighted Excel reports with detailed notes

            **Supported Formats:** CSV, Excel (.xlsx, .xls)
            """)

    # TAB 2: Pattern Analysis
    with tab2:
        st.header("📊 Pattern Analysis")

        if 'audit_results' in st.session_state and st.session_state['audit_results']:
            results = st.session_state['audit_results']

            # Combine all audit results
            all_audit_results = []
            for file_result in results['file_results']:
                if file_result['success']:
                    all_audit_results.extend(file_result['audit_results'])

            if all_audit_results:
                # Create statistics object
                stats = AuditStatistics(all_audit_results)

                # Red flag counts
                st.subheader("Red Flags by Type")
                flag_counts = stats.get_red_flag_counts()

                if flag_counts:
                    # Create DataFrame
                    df_flags = pd.DataFrame([
                        {"Red Flag Type": k, "Count": v}
                        for k, v in sorted(flag_counts.items(), key=lambda x: x[1], reverse=True)
                    ])

                    # Bar chart
                    fig = px.bar(
                        df_flags,
                        x="Count",
                        y="Red Flag Type",
                        orientation='h',
                        title="Red Flag Distribution",
                        color="Count",
                        color_continuous_scale="Reds"
                    )
                    fig.update_layout(height=400, showlegend=False)
                    st.plotly_chart(fig, use_container_width=True)

                st.markdown("---")

                # Red flag combinations
                st.subheader("Common Red Flag Combinations")
                combinations = stats.get_most_common_combinations(10)

                if combinations:
                    df_combos = pd.DataFrame(combinations, columns=["Combination", "Count"])
                    st.dataframe(df_combos, use_container_width=True)
                else:
                    st.info("No accounts with multiple red flags")

                st.markdown("---")

                # By date range
                st.subheader("Red Flags by Join Date Period")
                date_stats = stats.group_by_join_date_range()

                if date_stats:
                    df_dates = pd.DataFrame([
                        {
                            "Period": period,
                            "Total": data['total'],
                            "Flagged": data['flagged'],
                            "Flag %": data['flag_percentage']
                        }
                        for period, data in sorted(date_stats.items())
                    ])

                    fig = go.Figure()
                    fig.add_trace(go.Bar(
                        name='Clean',
                        x=df_dates['Period'],
                        y=df_dates['Total'] - df_dates['Flagged'],
                        marker_color='lightgreen'
                    ))
                    fig.add_trace(go.Bar(
                        name='Flagged',
                        x=df_dates['Period'],
                        y=df_dates['Flagged'],
                        marker_color='orange'
                    ))

                    fig.update_layout(
                        barmode='stack',
                        title="Records by Join Date Period",
                        xaxis_title="Period",
                        yaxis_title="Count",
                        height=400
                    )
                    st.plotly_chart(fig, use_container_width=True)

                st.markdown("---")

                # By sales rep
                st.subheader("Red Flags by Sales Representative")
                rep_stats = stats.group_by_sales_rep()

                if rep_stats:
                    df_reps = pd.DataFrame([
                        {
                            "Sales Rep": rep,
                            "Total Records": data['total'],
                            "Flagged": data['flagged'],
                            "Clean": data['clean'],
                            "Flag %": f"{data['flag_percentage']:.1f}%"
                        }
                        for rep, data in sorted(rep_stats.items(), key=lambda x: x[1]['flag_percentage'], reverse=True)
                    ])

                    st.dataframe(df_reps, use_container_width=True)

        else:
            st.info("📁 Upload and process files in the 'Upload & Audit' tab to see pattern analysis")

    # TAB 3: Financial Summary
    with tab3:
        st.header("💰 Financial Summary")

        if 'audit_results' in st.session_state and st.session_state['audit_results']:
            results = st.session_state['audit_results']

            # Combine all audit results
            all_audit_results = []
            for file_result in results['file_results']:
                if file_result['success']:
                    all_audit_results.extend(file_result['audit_results'])

            if all_audit_results:
                stats = AuditStatistics(all_audit_results)
                financial_summary = stats.get_financial_summary()

                # Key metrics
                col1, col2, col3 = st.columns(3)

                with col1:
                    display_metric_card(
                        "Total Financial Impact",
                        format_currency(financial_summary['total_impact']),
                        color="#F44336"
                    )

                with col2:
                    display_metric_card(
                        "Accounts with Impact",
                        f"{financial_summary['accounts_with_impact']:,}",
                        color="#FF9800"
                    )

                with col3:
                    display_metric_card(
                        "Avg Impact per Flagged Account",
                        format_currency(financial_summary['average_impact_per_flagged_account']),
                        color="#9C27B0"
                    )

                st.markdown("---")

                # Impact Breakdown
                st.subheader("📊 Financial Impact Breakdown")

                # Calculate breakdown
                total_dues = sum(r.get('dues_impact', 0) for r in all_audit_results)
                total_balance = sum(r.get('balance_impact', 0) for r in all_audit_results)

                # Get audit settings for dynamic info
                audit_settings = results.get('audit_settings', {})
                audit_type_name = audit_settings.get('membership_type_name', 'the selected membership type')
                audit_location_name = audit_settings.get('location_name', 'the selected location')

                # Info box explaining calculation
                st.info(
                    "**How is Financial Impact calculated?**\n\n"
                    "The Financial Impact is the sum of:\n"
                    f"- **Missing/Low Dues**: Expected dues for {audit_type_name} at {audit_location_name} minus what was actually paid (flagged if < 90%)\n"
                    "- **Outstanding Balances**: Absolute value of any non-zero balances\n\n"
                    "Some accounts may contribute to both categories."
                )

                col1, col2 = st.columns([1, 1])

                with col1:
                    # Pie chart
                    if total_dues > 0 or total_balance > 0:
                        fig = px.pie(
                            values=[total_dues, total_balance],
                            names=['Missing/Low Dues', 'Outstanding Balances'],
                            title="Impact by Category",
                            color_discrete_sequence=['#FF6B6B', '#4ECDC4']
                        )
                        fig.update_traces(textposition='inside', textinfo='percent+label')
                        st.plotly_chart(fig, use_container_width=True)

                with col2:
                    st.markdown("### **Breakdown Details**")
                    st.markdown(f"**Missing/Low Dues**")
                    st.markdown(f"💰 {format_currency(total_dues)}")
                    st.caption(f"{(total_dues / financial_summary['total_impact'] * 100):.1f}% of total" if financial_summary['total_impact'] > 0 else "")

                    st.markdown("")
                    st.markdown(f"**Outstanding Balances**")
                    st.markdown(f"💳 {format_currency(total_balance)}")
                    st.caption(f"{(total_balance / financial_summary['total_impact'] * 100):.1f}% of total" if financial_summary['total_impact'] > 0 else "")

                    st.markdown("")
                    st.markdown(f"**Total**")
                    st.markdown(f"🔴 {format_currency(financial_summary['total_impact'])}")

                st.markdown("---")

                # Top impact accounts
                st.subheader("Top Accounts by Financial Impact")
                top_accounts = stats.get_top_impact_accounts(20)

                if top_accounts:
                    df_top = pd.DataFrame([
                        {
                            "Member ID": acc['member_id'],
                            "Name": acc['member_name'],
                            "Financial Impact": format_currency(acc['financial_impact']),
                            "Red Flags": acc['flag_count'],
                            "Issues": ", ".join(acc['red_flags'][:2]) + ("..." if len(acc['red_flags']) > 2 else "")
                        }
                        for acc in top_accounts
                    ])

                    st.dataframe(df_top, use_container_width=True, hide_index=True)

                st.markdown("---")

                # Expired vs Active
                st.subheader("Active vs Expired Memberships")
                exp_stats = stats.get_expired_vs_active_stats()

                df_exp = pd.DataFrame([
                    {
                        "Status": status.title(),
                        "Total": data['total'],
                        "Flagged": data['flagged'],
                        "Flag %": f"{data['flag_percentage']:.1f}%"
                    }
                    for status, data in exp_stats.items()
                    if data['total'] > 0
                ])

                st.dataframe(df_exp, use_container_width=True, hide_index=True)

        else:
            st.info("📁 Upload and process files in the 'Upload & Audit' tab to see financial summary")


if __name__ == "__main__":
    main()
