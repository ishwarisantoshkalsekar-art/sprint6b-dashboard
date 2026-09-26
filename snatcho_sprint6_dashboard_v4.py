# ============================================================
# SNATCHO — SPRINT 6
# FOUNDER INTELLIGENCE DASHBOARD V4
# ============================================================

import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="SNATCHO | Sprint 6 Founder Intelligence",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ============================================================
# FILE PATHS
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

ANALYTICAL_FILE = BASE_DIR / "Snatcho_Clean_Analytical_Dataset.xlsx"
STUDENT_FILE = BASE_DIR / "student_summary.csv"
SEGMENT_FILE = BASE_DIR / "user_segment_summary.csv"


# ============================================================
# CUSTOM CSS
# ============================================================

st.markdown(
    """
    <style>

    /* ---------- GENERAL ---------- */

    .stApp {
        background-color: #f5f7fa;
    }

    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 3rem;
        max-width: 1500px;
    }

    /* ---------- SIDEBAR ---------- */

    section[data-testid="stSidebar"] {
        background-color: #eef1f5;
    }

    section[data-testid="stSidebar"] .block-container {
        padding-top: 2rem;
    }

    .brand {
        font-size: 28px;
        font-weight: 800;
        color: #26364a;
        margin-bottom: 4px;
    }

    .brand-subtitle {
        font-size: 13px;
        color: #7a8492;
        margin-bottom: 25px;
    }

    /* ---------- PAGE HEADER ---------- */

    .page-title {
        font-size: 34px;
        font-weight: 800;
        color: #182230;
        margin-bottom: 4px;
    }

    .page-subtitle {
        font-size: 15px;
        color: #667085;
        margin-bottom: 25px;
    }

    /* ---------- KPI CARDS ---------- */

    .kpi-card {
        background: #ffffff;
        border: 1px solid #e4e7ec;
        border-radius: 16px;
        padding: 22px;
        min-height: 145px;
        box-shadow: 0 3px 12px rgba(16, 24, 40, 0.05);
    }

    .kpi-title {
        font-size: 13px;
        color: #667085;
        font-weight: 650;
        margin-bottom: 12px;
    }

    .kpi-value {
        font-size: 31px;
        font-weight: 800;
        color: #101828;
        line-height: 1.1;
        margin-bottom: 9px;
    }

    .kpi-subtitle {
        font-size: 12px;
        color: #98a2b3;
    }

    /* ---------- SECTION ---------- */

    .section-title {
        font-size: 22px;
        font-weight: 750;
        color: #182230;
        margin-top: 28px;
        margin-bottom: 12px;
    }

    .section-subtitle {
        font-size: 13px;
        color: #667085;
        margin-bottom: 15px;
    }

    /* ---------- INFO CARDS ---------- */

    .info-card {
        background: #ffffff;
        border: 1px solid #e4e7ec;
        border-radius: 14px;
        padding: 20px;
        margin-bottom: 12px;
        box-shadow: 0 2px 8px rgba(16, 24, 40, 0.04);
    }

    .info-title {
        font-size: 14px;
        font-weight: 700;
        color: #344054;
        margin-bottom: 8px;
    }

    .info-text {
        font-size: 14px;
        color: #667085;
        line-height: 1.6;
    }

    /* ---------- DECISION CARD ---------- */

    .decision-card {
        background: #ffffff;
        border: 1px solid #d0d5dd;
        border-radius: 16px;
        padding: 25px;
        margin-top: 18px;
        box-shadow: 0 4px 12px rgba(16, 24, 40, 0.05);
    }

    .decision-label {
        font-size: 12px;
        font-weight: 750;
        color: #667085;
        letter-spacing: 0.08em;
        text-transform: uppercase;
        margin-bottom: 8px;
    }

    .decision-text {
        font-size: 25px;
        font-weight: 800;
        color: #182230;
        margin-bottom: 8px;
    }

    .decision-description {
        font-size: 14px;
        color: #667085;
        line-height: 1.6;
    }

    /* ---------- GAP CARD ---------- */

    .gap-card {
        background: #fffaf0;
        border: 1px solid #f0dfb5;
        border-radius: 14px;
        padding: 18px;
        margin-bottom: 12px;
    }

    .gap-title {
        font-weight: 750;
        color: #7a4e00;
        margin-bottom: 6px;
    }

    .gap-text {
        font-size: 13px;
        color: #795548;
        line-height: 1.5;
    }

    /* ---------- EVIDENCE CARD ---------- */

    .evidence-card {
        background: #f8fafc;
        border-left: 4px solid #667085;
        padding: 15px 18px;
        margin: 10px 0;
        border-radius: 8px;
    }

    .evidence-title {
        font-weight: 700;
        color: #344054;
        margin-bottom: 4px;
    }

    .evidence-text {
        color: #667085;
        font-size: 13px;
        line-height: 1.5;
    }

    /* ---------- TABLE ---------- */

    .table-title {
        font-size: 16px;
        font-weight: 750;
        color: #344054;
        margin-top: 20px;
        margin-bottom: 10px;
    }

    /* ---------- FOOTER ---------- */

    .footer {
        text-align: center;
        color: #98a2b3;
        font-size: 12px;
        padding-top: 35px;
        padding-bottom: 15px;
    }

    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# LOAD DATA
# ============================================================

@st.cache_data
def load_data():

    if not ANALYTICAL_FILE.exists():
        return None, None, None

    try:
        df = pd.read_excel(
            ANALYTICAL_FILE,
            sheet_name="Analytical Dataset"
        )

        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
        )

        student_summary = None
        segment_summary = None

        if STUDENT_FILE.exists():
            try:
                student_summary = pd.read_csv(STUDENT_FILE)
            except Exception:
                student_summary = None

        if SEGMENT_FILE.exists():
            try:
                segment_summary = pd.read_csv(SEGMENT_FILE)
            except Exception:
                segment_summary = None

        return df, student_summary, segment_summary

    except Exception as e:
        st.error(f"Unable to load analytical dataset: {e}")
        return None, None, None


df, student_summary, segment_summary = load_data()


# ============================================================
# FILE ERROR
# ============================================================

if df is None:

    st.error(
        "Analytical dataset not found."
    )

    st.info(
        "Make sure this file is in the same folder as the Python dashboard:"
    )

    st.code(
        "Snatcho_Clean_Analytical_Dataset(5).xlsx"
    )

    st.stop()


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def clean_series(column):

    if column not in df.columns:
        return None

    return (
        df[column]
        .dropna()
        .astype(str)
        .str.strip()
    )


def percentage(column, target):

    series = clean_series(column)

    if series is None or len(series) == 0:
        return None

    target_clean = str(target).strip().lower()

    matches = (
        series
        .str.lower()
        .eq(target_clean)
        .sum()
    )

    return round((matches / len(series)) * 100, 1)


def percentage_contains(column, keyword):

    series = clean_series(column)

    if series is None or len(series) == 0:
        return None

    matches = (
        series
        .str.lower()
        .str.contains(str(keyword).lower(), na=False)
        .sum()
    )

    return round((matches / len(series)) * 100, 1)


def count_unique(column):

    series = clean_series(column)

    if series is None:
        return None

    return series.nunique()


def show_kpi(title, value, subtitle):

    if value is None:
        value = "N/A"

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-title">{title}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-subtitle">{subtitle}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


def show_info(title, text):

    st.markdown(
        f"""
        <div class="info-card">
            <div class="info-title">{title}</div>
            <div class="info-text">{text}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


def show_gap(title, text):

    st.markdown(
        f"""
        <div class="gap-card">
            <div class="gap-title">{title}</div>
            <div class="gap-text">{text}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


def show_decision(title, description):

    st.markdown(
        f"""
        <div class="decision-card">
            <div class="decision-label">Founder Decision Signal</div>
            <div class="decision-text">{title}</div>
            <div class="decision-description">{description}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


def show_bar_chart(column, title):

    series = clean_series(column)

    if series is None or series.empty:
        st.info(f"No data available for {title}.")
        return

    counts = (
        series
        .value_counts()
        .reset_index()
    )

    counts.columns = ["Response", "Count"]

    fig = px.bar(
        counts,
        x="Response",
        y="Count",
        title=title,
        text="Count"
    )

    fig.update_layout(
        height=400,
        template="plotly_white",
        xaxis_title="",
        yaxis_title="Responses",
        showlegend=False
    )

    fig.update_traces(
        textposition="outside"
    )

    st.plotly_chart(
        fig,
        use_container_width=True
    )


def show_pie_chart(column, title):

    series = clean_series(column)

    if series is None or series.empty:
        st.info(f"No data available for {title}.")
        return

    counts = (
        series
        .value_counts()
        .reset_index()
    )

    counts.columns = ["Response", "Count"]

    fig = px.pie(
        counts,
        names="Response",
        values="Count",
        title=title,
        hole=0.45
    )

    fig.update_layout(
        height=400,
        template="plotly_white"
    )

    st.plotly_chart(
        fig,
        use_container_width=True
    )


def make_cross_tab(column1, column2):

    if column1 not in df.columns or column2 not in df.columns:
        return None

    result = pd.crosstab(
        df[column1],
        df[column2]
    )

    return result


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.markdown(
        """
        <div class="brand">SNATCHO</div>
        <div class="brand-subtitle">
            Sprint 6 • Founder Intelligence
        </div>
        """,
        unsafe_allow_html=True
    )

    st.divider()

    page = st.radio(
        "Dashboard",
        [
            "Founder Overview",
            "User Intelligence",
            "Product Intelligence",
            "Growth & Strategy"
        ]
    )

    st.divider()

    st.caption(
        f"Analytical responses: {len(df):,}"
    )

    st.caption(
        "Evidence base: Survey / Analytical Dataset"
    )

    st.caption(
        "Live pilot data: Not available"
    )


# ============================================================
# PAGE 1 — FOUNDER OVERVIEW
# ============================================================

if page == "Founder Overview":

    st.markdown(
        '<div class="page-title">Founder Overview</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="page-subtitle">'
        'Sprint 6 decision layer — turning user evidence into product signals'
        '</div>',
        unsafe_allow_html=True
    )

    # --------------------------------------------------------
    # KPI CALCULATIONS
    # --------------------------------------------------------

    comparison_pain = percentage_contains(
        "Comparison Pain",
        "yes"
    )

    choice_confusion = percentage_contains(
        "Choice Confusion",
        "yes"
    )

    automatic_comparison = percentage_contains(
        "Automatic Price Comparison",
        "yes"
    )

    price_alerts = percentage_contains(
        "Price Drop Alerts",
        "yes"
    )

    trial_intent = percentage_contains(
        "₹500 Trial Intent",
        "definitely"
    )

    # --------------------------------------------------------
    # KPI ROW
    # --------------------------------------------------------

    cols = st.columns(5)

    with cols[0]:
        show_kpi(
            "Comparison Pain",
            f"{comparison_pain}%" if comparison_pain is not None else "N/A",
            "Survey evidence"
        )

    with cols[1]:
        show_kpi(
            "Choice Confusion",
            f"{choice_confusion}%" if choice_confusion is not None else "N/A",
            "Survey evidence"
        )

    with cols[2]:
        show_kpi(
            "Auto Comparison",
            f"{automatic_comparison}%" if automatic_comparison is not None else "N/A",
            "Feature demand"
        )

    with cols[3]:
        show_kpi(
            "Price Alerts",
            f"{price_alerts}%" if price_alerts is not None else "N/A",
            "Interest signal"
        )

    with cols[4]:
        show_kpi(
            "Trial Intent",
            f"{trial_intent}%" if trial_intent is not None else "N/A",
            "Survey intent"
        )

    # --------------------------------------------------------
    # CORE PROBLEM
    # --------------------------------------------------------

    st.markdown(
        '<div class="section-title">Core User Problem</div>',
        unsafe_allow_html=True
    )

    col1, col2 = st.columns(2)

    with col1:

        show_info(
            "Comparison Friction",
            "Survey evidence indicates that users experience difficulty "
            "comparing similar products and evaluating alternatives before purchase."
        )

    with col2:

        show_info(
            "Decision Support Opportunity",
            "Users express demand for automated comparison, price monitoring "
            "and decision-support features."
        )

    # --------------------------------------------------------
    # PAIN LANDSCAPE
    # --------------------------------------------------------

    st.markdown(
        '<div class="section-title">Pain Landscape</div>',
        unsafe_allow_html=True
    )

    col1, col2 = st.columns(2)

    with col1:

        if "Comparison Pain" in df.columns:
            show_bar_chart(
                "Comparison Pain",
                "Comparison Pain Distribution"
            )

    with col2:

        if "Choice Confusion" in df.columns:
            show_bar_chart(
                "Choice Confusion",
                "Choice Confusion Distribution"
            )

    # --------------------------------------------------------
    # COMMERCIAL SIGNAL
    # --------------------------------------------------------

    st.markdown(
        '<div class="section-title">Commercial Signal</div>',
        unsafe_allow_html=True
    )

    col1, col2 = st.columns(2)

    with col1:

        if "₹500 Trial Intent" in df.columns:
            show_bar_chart(
                "₹500 Trial Intent",
                "Willingness to Try / Pay"
            )

    with col2:

        if "Price Drop Alerts" in df.columns:
            show_bar_chart(
                "Price Drop Alerts",
                "Price Drop Alert Interest"
            )

    # --------------------------------------------------------
    # VALIDATION GAP
    # --------------------------------------------------------

    st.markdown(
        '<div class="section-title">Sprint 6 Validation Gap</div>',
        unsafe_allow_html=True
    )

    show_gap(
        "Live Product Behaviour",
        "The available dataset contains survey evidence. "
        "A controlled 30-user live pilot was not completed, so activation, "
        "repeat usage, recommendation success and real purchase behaviour "
        "should not be presented as measured live-product metrics."
    )

    show_decision(
        "CONTINUE VALIDATING",
        "The survey evidence establishes a meaningful user problem and "
        "feature interest, but live product evidence is still required "
        "before making strong claims about activation, retention or "
        "recommendation reliability."
    )


# ============================================================
# PAGE 2 — USER INTELLIGENCE
# ============================================================

elif page == "User Intelligence":

    st.markdown(
        '<div class="page-title">User Intelligence</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="page-subtitle">'
        'Who responded, how they shop and where comparison friction appears'
        '</div>',
        unsafe_allow_html=True
    )

    # --------------------------------------------------------
    # USER KPIs
    # --------------------------------------------------------

    total_users = len(df)

    student_pct = percentage(
        "Student Flag",
        "Yes"
    )

    frequent_shoppers = percentage_contains(
        "Shopping Frequency",
        "weekly"
    )

    cross_platform = percentage_contains(
        "Cross-Platform Checking",
        "yes"
    )

    wrong_purchase = percentage_contains(
        "Wrong Purchase",
        "yes"
    )

    cols = st.columns(4)

    with cols[0]:
        show_kpi(
            "Survey Responses",
            total_users,
            "Analytical dataset"
        )

    with cols[1]:
        show_kpi(
            "Students",
            f"{student_pct}%" if student_pct is not None else "N/A",
            "Respondent mix"
        )

    with cols[2]:
        show_kpi(
            "Cross-Platform Checking",
            f"{cross_platform}%" if cross_platform is not None else "N/A",
            "Shopping behaviour"
        )

    with cols[3]:
        show_kpi(
            "Wrong Purchase Signal",
            f"{wrong_purchase}%" if wrong_purchase is not None else "N/A",
            "Reported experience"
        )

    # --------------------------------------------------------
    # USER PROFILE
    # --------------------------------------------------------

    st.markdown(
        '<div class="section-title">User Profile</div>',
        unsafe_allow_html=True
    )

    col1, col2 = st.columns(2)

    with col1:

        if "Student Flag" in df.columns:
            show_pie_chart(
                "Student Flag",
                "Student vs Non-Student"
            )

    with col2:

        if "User Type" in df.columns:
            show_bar_chart(
                "User Type",
                "User Type Distribution"
            )

    # --------------------------------------------------------
    # DEMOGRAPHICS
    # --------------------------------------------------------

    st.markdown(
        '<div class="section-title">Demographic Structure</div>',
        unsafe_allow_html=True
    )

    col1, col2 = st.columns(2)

    with col1:

        if "Age Group" in df.columns:
            show_bar_chart(
                "Age Group",
                "Age Group"
            )

    with col2:

        if "Shopping Frequency" in df.columns:
            show_bar_chart(
                "Shopping Frequency",
                "Shopping Frequency"
            )

    # --------------------------------------------------------
    # SHOPPING BEHAVIOUR
    # --------------------------------------------------------

    st.markdown(
        '<div class="section-title">Shopping Behaviour</div>',
        unsafe_allow_html=True
    )

    col1, col2 = st.columns(2)

    with col1:

        if "Spend Per Order" in df.columns:
            show_bar_chart(
                "Spend Per Order",
                "Typical Spend Per Order"
            )

    with col2:

        if "Comparison Time" in df.columns:
            show_bar_chart(
                "Comparison Time",
                "Time Spent Comparing Products"
            )

    # --------------------------------------------------------
    # CROSS PLATFORM
    # --------------------------------------------------------

    if "Cross-Platform Checking" in df.columns:

        st.markdown(
            '<div class="section-title">Cross-Platform Behaviour</div>',
            unsafe_allow_html=True
        )

        show_bar_chart(
            "Cross-Platform Checking",
            "Do Users Check Multiple Platforms?"
        )

    # --------------------------------------------------------
    # STUDENT VS NON STUDENT
    # --------------------------------------------------------

    if (
        "Student Flag" in df.columns
        and "Comparison Pain" in df.columns
    ):

        st.markdown(
            '<div class="section-title">Student vs Non-Student Comparison Pain</div>',
            unsafe_allow_html=True
        )

        cross = pd.crosstab(
            df["Student Flag"],
            df["Comparison Pain"],
            normalize="index"
        ) * 100

        cross = cross.round(1)

        st.dataframe(
            cross,
            use_container_width=True
        )

    # --------------------------------------------------------
    # INTERPRETATION
    # --------------------------------------------------------

    show_info(
        "Founder Interpretation",
        "The survey can identify user segments and shopping behaviour, "
        "but it should not be treated as a substitute for observed product "
        "behaviour. Student and non-student differences are useful for "
        "segmentation and hypothesis formation."
    )


# ============================================================
# PAGE 3 — PRODUCT INTELLIGENCE
# ============================================================

elif page == "Product Intelligence":

    st.markdown(
        '<div class="page-title">Product Intelligence</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="page-subtitle">'
        'Feature demand, adoption signals, trust requirements and barriers'
        '</div>',
        unsafe_allow_html=True
    )

    # --------------------------------------------------------
    # PRODUCT KPIs
    # --------------------------------------------------------

    install_intent = percentage_contains(
        "Install Intent",
        "yes"
    )

    early_access = percentage_contains(
        "Early Access",
        "yes"
    )

    alert_interest = percentage_contains(
        "Price Drop Alerts",
        "yes"
    )

    automatic_feature = percentage_contains(
        "Automatic Price Comparison",
        "yes"
    )

    cols = st.columns(4)

    with cols[0]:
        show_kpi(
            "Install Intent",
            f"{install_intent}%" if install_intent is not None else "N/A",
            "Survey signal"
        )

    with cols[1]:
        show_kpi(
            "Early Access",
            f"{early_access}%" if early_access is not None else "N/A",
            "Survey signal"
        )

    with cols[2]:
        show_kpi(
            "Price Alerts",
            f"{alert_interest}%" if alert_interest is not None else "N/A",
            "Feature interest"
        )

    with cols[3]:
        show_kpi(
            "Auto Comparison",
            f"{automatic_feature}%" if automatic_feature is not None else "N/A",
            "Feature demand"
        )

    # --------------------------------------------------------
    # TRIAL / ADOPTION
    # --------------------------------------------------------

    st.markdown(
        '<div class="section-title">Adoption Signals</div>',
        unsafe_allow_html=True
    )

    col1, col2 = st.columns(2)

    with col1:

        if "Install Intent" in df.columns:
            show_bar_chart(
                "Install Intent",
                "Install Intent"
            )

    with col2:

        if "Early Access" in df.columns:
            show_bar_chart(
                "Early Access",
                "Early Access Interest"
            )

    # --------------------------------------------------------
    # FEATURE PRIORITY
    # --------------------------------------------------------

    st.markdown(
        '<div class="section-title">Feature Demand</div>',
        unsafe_allow_html=True
    )

    col1, col2 = st.columns(2)

    with col1:

        if "Feature Priority" in df.columns:
            show_bar_chart(
                "Feature Priority",
                "Feature Priority"
            )

    with col2:

        if "Most Useful Benefit" in df.columns:
            show_bar_chart(
                "Most Useful Benefit",
                "Most Useful Benefit"
            )

    # --------------------------------------------------------
    # FORMAT
    # --------------------------------------------------------

    st.markdown(
        '<div class="section-title">Preferred Product Format</div>',
        unsafe_allow_html=True
    )

    if "Preferred Tool Format" in df.columns:

        show_bar_chart(
            "Preferred Tool Format",
            "Preferred Tool Format"
        )

    # --------------------------------------------------------
    # TRUST
    # --------------------------------------------------------

    st.markdown(
        '<div class="section-title">Trust Requirements</div>',
        unsafe_allow_html=True
    )

    if "Trust Builder" in df.columns:

        show_bar_chart(
            "Trust Builder",
            "What Builds Recommendation Trust?"
        )

    # --------------------------------------------------------
    # BARRIERS
    # --------------------------------------------------------

    st.markdown(
        '<div class="section-title">Adoption Barriers</div>',
        unsafe_allow_html=True
    )

    if "Adoption Barrier" in df.columns:

        show_bar_chart(
            "Adoption Barrier",
            "What Could Stop Users From Adopting?"
        )

    # --------------------------------------------------------
    # PRICE
    # --------------------------------------------------------

    if "₹500 Trial Intent" in df.columns:

        st.markdown(
            '<div class="section-title">Trial / Payment Intent</div>',
            unsafe_allow_html=True
        )

        show_bar_chart(
            "₹500 Trial Intent",
            "Response to ₹500 Trial"
        )

    # --------------------------------------------------------
    # PRODUCT VALIDATION GAP
    # --------------------------------------------------------

    st.markdown(
        '<div class="section-title">Product Validation Gap</div>',
        unsafe_allow_html=True
    )

    show_gap(
        "Survey Intent ≠ Product Usage",
        "Install intent, early-access interest and feature demand are "
        "survey signals. They do not establish actual activation, "
        "successful recommendations, repeat usage or retention."
    )

    show_gap(
        "Recommendation Reliability",
        "Live product testing is required to measure whether SNAIA actually "
        "returns relevant products, accurate prices, functioning links and "
        "useful comparisons."
    )


# ============================================================
# PAGE 4 — GROWTH & STRATEGY
# ============================================================

elif page == "Growth & Strategy":

    st.markdown(
        '<div class="page-title">Growth & Strategy</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="page-subtitle">'
        'Segment signals, acquisition evidence and strategic assumptions'
        '</div>',
        unsafe_allow_html=True
    )

    # --------------------------------------------------------
    # GROWTH KPIs
    # --------------------------------------------------------

    total_users = len(df)

    survey_sources = count_unique(
        "Survey Source"
    )

    students = percentage(
        "Student Flag",
        "Yes"
    )

    non_students = None

    if students is not None:
        non_students = round(100 - students, 1)

    cols = st.columns(4)

    with cols[0]:
        show_kpi(
            "Evidence Base",
            total_users,
            "Survey respondents"
        )

    with cols[1]:
        show_kpi(
            "Survey Sources",
            survey_sources if survey_sources is not None else "N/A",
            "Reported acquisition sources"
        )

    with cols[2]:
        show_kpi(
            "Students",
            f"{students}%" if students is not None else "N/A",
            "Respondent segment"
        )

    with cols[3]:
        show_kpi(
            "Non-Students",
            f"{non_students}%" if non_students is not None else "N/A",
            "Respondent segment"
        )

    # --------------------------------------------------------
    # SURVEY ACQUISITION SOURCE
    # --------------------------------------------------------

    st.markdown(
        '<div class="section-title">Survey Acquisition Sources</div>',
        unsafe_allow_html=True
    )

    if "Survey Source" in df.columns:

        show_bar_chart(
            "Survey Source",
            "Where Survey Respondents Came From"
        )

    else:

        show_gap(
            "Growth Channel Data Unavailable",
            "No survey source field is available in the analytical dataset."
        )

    # --------------------------------------------------------
    # STUDENT / NON STUDENT
    # --------------------------------------------------------

    st.markdown(
        '<div class="section-title">Segment Intelligence</div>',
        unsafe_allow_html=True
    )

    col1, col2 = st.columns(2)

    with col1:

        if "Student Flag" in df.columns:
            show_pie_chart(
                "Student Flag",
                "Student / Non-Student Mix"
            )

    with col2:

        if "User Type" in df.columns:
            show_bar_chart(
                "User Type",
                "User Type Distribution"
            )

    # --------------------------------------------------------
    # STRATEGIC ASSUMPTION REGISTER
    # --------------------------------------------------------

    st.markdown(
        '<div class="section-title">Strategic Assumption Register</div>',
        unsafe_allow_html=True
    )

    assumptions = pd.DataFrame(
        {
            "Assumption": [
                "Users experience comparison friction",
                "Users want automated comparison",
                "Users value price monitoring",
                "Users will trust AI recommendations",
                "Users will repeatedly use SNAIA",
                "SNAIA can reliably return useful products",
                "Users will convert from interest to actual usage"
            ],
            "Current Evidence": [
                "Survey",
                "Survey",
                "Survey",
                "Survey / Trust questions",
                "Not measured",
                "Not sufficiently measured",
                "Not measured"
            ],
            "Status": [
                "SUPPORTED BY SURVEY",
                "SUPPORTED BY SURVEY",
                "SUPPORTED BY SURVEY",
                "REQUIRES PRODUCT VALIDATION",
                "DATA NOT AVAILABLE",
                "REQUIRES PRODUCT VALIDATION",
                "DATA NOT AVAILABLE"
            ]
        }
    )

    st.dataframe(
        assumptions,
        use_container_width=True,
        hide_index=True
    )

    # --------------------------------------------------------
    # BIGGEST PRODUCT PROBLEM
    # --------------------------------------------------------

    st.markdown(
        '<div class="section-title">Biggest Product Problem to Validate</div>',
        unsafe_allow_html=True
    )

    show_info(
        "Recommendation Reliability",
        "The most important product-level validation question is whether "
        "SNAIA can consistently transform a shopping query into relevant, "
        "trustworthy and actionable recommendations."
    )

    # --------------------------------------------------------
    # BIGGEST GROWTH OPPORTUNITY
    # --------------------------------------------------------

    st.markdown(
        '<div class="section-title">Growth Opportunity to Test</div>',
        unsafe_allow_html=True
    )

    show_info(
        "High-Intent Shopping Moments",
        "Survey evidence suggests that comparison-heavy shopping moments "
        "are a natural context for the product. Growth experiments should "
        "test acquisition and activation around these specific use cases."
    )

    # --------------------------------------------------------
    # DATA GAPS
    # --------------------------------------------------------

    st.markdown(
        '<div class="section-title">Critical Data Gaps</div>',
        unsafe_allow_html=True
    )

    show_gap(
        "Live Acquisition Funnel",
        "No validated live channel → signup → activation funnel is available."
    )

    show_gap(
        "Activation",
        "No reliable live-product activation rate is available."
    )

    show_gap(
        "Retention",
        "No repeat-user or retention dataset is available."
    )

    show_gap(
        "Recommendation Success",
        "No controlled live-user recommendation success dataset is available."
    )

    show_gap(
        "Conversion",
        "No validated purchase/conversion behaviour is available."
    )

    # --------------------------------------------------------
    # FOUNDER DECISION
    # --------------------------------------------------------

    show_decision(
        "VALIDATE BEFORE SCALING",
        "The survey establishes demand-side evidence, while several "
        "critical product and growth assumptions remain unvalidated. "
        "The next evidence step should focus on controlled real-user "
        "product behaviour rather than treating survey intent as usage."
    )


# ============================================================
# FOOTER
# ============================================================

st.markdown(
    """
    <div class="footer">
        SNATCHO • Sprint 6 • Founder Intelligence Dashboard V4
        <br>
        Evidence shown from the analytical survey dataset.
        Live product metrics are not fabricated where unavailable.
    </div>
    """,
    unsafe_allow_html=True
)