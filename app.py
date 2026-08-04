import streamlit as st
import pandas as pd
from datetime import datetime
from itertools import combinations
from openpyxl import load_workbook
import os
import ssl
import certifi

ssl._create_default_https_context = ssl._create_unverified_context

os.environ["SSL_CERT_FILE"] = certifi.where()
os.environ["REQUESTS_CA_BUNDLE"] = certifi.where()

import gspread
from google.oauth2.service_account import Credentials

# ==================================================
# PAGE CONFIG
# ==================================================

st.set_page_config(
    page_title="PvP IB Cricket Dashboard",
    layout="wide"
)

st.markdown("""
<style>

/* Mobile phones */
@media (max-width: 768px) {

    html, body, [class*="css"] {
        font-size: 20px !important;
    }

    p, div, label {
        font-size: 18px !important;
    }

    h1 {
        font-size: 36px !important;
    }

    h2 {
        font-size: 30px !important;
    }

    h3 {
        font-size: 24px !important;
    }

    button {
        font-size: 18px !important;
    }
}

</style>
""", unsafe_allow_html=True)


# ==================================================
# CONFIG
# ==================================================

MATCH_SHEET = "Online_Match_Entry"
CALCULATED_POINTS_SHEET = "Calculated_Points_Table"

WIN_POINTS = 2
TIE_POINTS = 1
LOSS_POINTS = 0
NO_RESULT_POINTS = 0

MAX_OVERS = 10
MAX_WICKETS = 5
ACCESS_FILE = "Access.xlsx"

# ==========================================
# GOOGLE SHEETS CONNECTION
# ==========================================

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# Force Python to use certifi certificates
os.environ["SSL_CERT_FILE"] = certifi.where()
os.environ["REQUESTS_CA_BUNDLE"] = certifi.where()

creds = Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=SCOPES
)

client = gspread.authorize(creds)

season_sheets = {
    
    "Season 2": "PvP IB Cricket Dashboard - Season 2",
    "Season 3": "PvP IB Cricket Dashboard - Season 3"
}



# ==================================================
# SESSION STATE
# ==================================================

if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

if "role" not in st.session_state:
    st.session_state["role"] = None

if "username" not in st.session_state:
    st.session_state["username"] = None

# ==================================================
# HELPER FOR ACCESS
# ==================================================

def load_users():

    try:

        sheet = client.open(
            "Access"
        ).worksheet(
            "Users"
        )

        return pd.DataFrame(
            sheet.get_all_records()
        )

    except Exception:

        return pd.DataFrame(
            columns=[
                "Username",
                "Password",
                "Role",
                "Status"
            ]
        )


def load_access_requests():

    try:

        sheet = client.open(
            "Access"
        ).worksheet(
            "Access_Requests"
        )

        return pd.DataFrame(
            sheet.get_all_records()
        )

    except Exception:

        return pd.DataFrame(
            columns=[
                "Username",
                "Email",
                "RequestedOn",
                "Status"
            ]
        )


season_options = [
    "Season 2",
    "Season 3"
]

# ==================================================
# TITLE AND SEASON SELECTOR
# ==================================================

title_col, season_col = st.columns([8, 2])

with title_col:
    st.title("🏏 PvP IB Cricket Dashboard")

with season_col:
    st.markdown("**🏆 SEASON**")

    season = st.selectbox(
        "",
        season_options,
        label_visibility="collapsed"
    )

#FILE = season_files[season]

#if not os.path.exists(FILE):
#    st.error(
#        f"Excel file not found: {FILE}. Please create this file in the same folder as app.py."
#    )
#    st.stop()

# ==================================================
# LOAD TEAMS MASTER
# ==================================================

try:
    sheet = client.open(
        season_sheets[season]
    ).worksheet(
        "Teams_Master"
    )

    teams_df = pd.DataFrame(
        sheet.get_all_records()
    )

    teams_df.columns = teams_df.columns.str.strip()

except Exception as e:
    st.error(
        f"Unable to read Teams_Master sheet from. Error: {e}"
    )
    st.stop()

required_team_columns = ["Group", "Team"]

for col in required_team_columns:
    if col not in teams_df.columns:
        st.error(
            f"Column '{col}' is missing in Teams_Master sheet."
        )
        st.stop()

groups = {}

group_order = (
    teams_df[
        ["SortOrder", "Group"]
    ]
    .drop_duplicates()
    .sort_values("SortOrder")
)

for _, row in group_order.iterrows():

    group_name = row["Group"]

    groups[group_name] = (
        teams_df.loc[
            teams_df["Group"] == group_name,
            "Team"
        ]
        .dropna()
        .tolist()
    )

all_teams = []

for team_list in groups.values():
    all_teams.extend(team_list)

# ==================================================
# LOAD ONLINE MATCH ENTRIES
# ==================================================

def load_match_entries():

    try:

        sheet = client.open(
            season_sheets[season]
        ).worksheet(
            "Online_Match_Entry"
        )

        history = pd.DataFrame(
            sheet.get_all_records()
        )

        if "Status" not in history.columns:
            history["Status"] = "Active"

        return history

    except Exception as e:

        st.error(f"Google Sheet Error: {e}")

        columns = [
            "Date",
            "Group",
            "TeamA",
            "RunsA",
            "WicketsA",
            "OversA",
            "TeamB",
            "RunsB",
            "WicketsB",
            "OversB",
            "Winner",
            "Status"
        ]

        return pd.DataFrame(columns=columns)


def load_knockout_matches():

    try:

        sheet = client.open(
            season_sheets[season]
        ).worksheet(
            "Knockout_Matches"
        )

        return pd.DataFrame(
            sheet.get_all_records()
        )

    except Exception as e:

        st.error(f"Knockout Sheet Error: {e}")

        return pd.DataFrame(
            columns=[
                "Stage",
                "Match",
                "TeamA",
                "RunsA",
                "WicketsA",
                "OversA",
                "TeamB",
                "RunsB",
                "WicketsB",
                "OversB",
                "Winner"
            ]
        )

    
# ==================================================
# SAVE MATCH ENTRY
# ==================================================
def save_match(match_data):


    sheet = client.open(
        season_sheets[season]
    ).worksheet(
        "Online_Match_Entry"
    )

    sheet.append_row(match_data)

# ==================================================
# CRICKET OVERS CONVERSION
# ==================================================
# Input 7.3 means 7 overs and 3 balls.
# Internally convert to decimal overs: 7 + 3/6 = 7.5

def convert_overs(value):

    try:
        value = float(value)

    except Exception:
        return 0.0

    whole_overs = int(value)
    decimal_part = round(value - whole_overs, 1)

    balls = int(round(decimal_part * 10, 0))

    if balls >= 0 and balls <= 5:
        return whole_overs + (balls / 6)

    return value

# ==================================================
# DISPLAY OVERS IN CRICKET FORMAT
# ==================================================
# Internal 7.5 decimal overs becomes 7.3 cricket overs.

def decimal_to_cricket_overs(value):

    try:
        value = float(value)

    except Exception:
        return "0.0"

    whole_overs = int(value)
    balls = round((value - whole_overs) * 6)

    if balls == 6:
        whole_overs += 1
        balls = 0

    return f"{whole_overs}.{balls}"

# ==================================================
# ICC STYLE OVERS FOR NRR
# ==================================================
# Tournament rule:
# 10 overs maximum
# 5 wickets maximum
#
# If team is all out for 5 wickets, NRR uses full 10 overs.
# Display still shows actual overs played.

def get_overs_for_nrr(overs_entered, wickets_lost):

    actual_overs = convert_overs(overs_entered)

    try:
        wickets_lost = int(wickets_lost)

    except Exception:
        wickets_lost = 0

    if wickets_lost >= MAX_WICKETS:
        return float(MAX_OVERS)

    return actual_overs

# ==================================================
# CALCULATE POINTS TABLE
# ==================================================

def calculate_points_table(group_name, team_list, match_history):

    table = {}

    for team in team_list:

        table[team] = {
            "Team": team,
            "Played": 0,
            "Wins": 0,
            "Losses": 0,
            "Ties": 0,
            "NR": 0,
            "Points": 0,
            "RunsFor": 0,
            "RunsAgainst": 0,
            "ActualOversFor": 0.0,
            "ActualOversAgainst": 0.0,
            "NrrOversFor": 0.0,
            "NrrOversAgainst": 0.0
        }

    if match_history.empty:

        result = pd.DataFrame(table.values())

        if result.empty:
            return pd.DataFrame(
                columns=[
                    "Rank",
                    "Team",
                    "Played",
                    "Wins",
                    "Losses",
                    "Ties",
                    "NR",
                    "Points",
                    "RunsFor",
                    "ActualOversFor",
                    "RunsAgainst",
                    "ActualOversAgainst",
                    "Scored",
                    "Conceded",
                    "NRR"
                ]
            )

        result["NRR"] = 0.000
        result["Scored"] = "0 / 0.0"
        result["Conceded"] = "0 / 0.0"
        result["Rank"] = range(1, len(result) + 1)

        return result[
            [
                "Rank",
                "Team",
                "Played",
                "Wins",
                "Losses",
                "Ties",
                "NR",
                "Points",
                "RunsFor",
                "ActualOversFor",
                "RunsAgainst",
                "ActualOversAgainst",
                "Scored",
                "Conceded",
                "NRR"
            ]
        ]

    for _, row in match_history.iterrows():

        status = row.get("Status", "Active")

        if status == "Deleted":
            continue

        row_group = row.get("Group", group_name)

        if pd.notna(row_group) and row_group != group_name:
            continue

        team_a = row.get("TeamA")
        team_b = row.get("TeamB")

        if team_a not in team_list or team_b not in team_list:
            continue

        runs_a = int(row.get("RunsA", 0))
        runs_b = int(row.get("RunsB", 0))

        wickets_a = int(row.get("WicketsA", 0))
        wickets_b = int(row.get("WicketsB", 0))

        actual_overs_a = convert_overs(
            row.get("OversA", 0)
        )

        actual_overs_b = convert_overs(
            row.get("OversB", 0)
        )

        nrr_overs_a = get_overs_for_nrr(
            row.get("OversA", 0),
            wickets_a
        )

        nrr_overs_b = get_overs_for_nrr(
            row.get("OversB", 0),
            wickets_b
        )

        winner = row.get("Winner", "")

        table[team_a]["Played"] += 1
        table[team_b]["Played"] += 1

        table[team_a]["RunsFor"] += runs_a
        table[team_a]["RunsAgainst"] += runs_b
        table[team_a]["ActualOversFor"] += actual_overs_a
        table[team_a]["ActualOversAgainst"] += actual_overs_b
        table[team_a]["NrrOversFor"] += nrr_overs_a
        table[team_a]["NrrOversAgainst"] += nrr_overs_b

        table[team_b]["RunsFor"] += runs_b
        table[team_b]["RunsAgainst"] += runs_a
        table[team_b]["ActualOversFor"] += actual_overs_b
        table[team_b]["ActualOversAgainst"] += actual_overs_a
        table[team_b]["NrrOversFor"] += nrr_overs_b
        table[team_b]["NrrOversAgainst"] += nrr_overs_a

        if winner == team_a:

            table[team_a]["Wins"] += 1
            table[team_a]["Points"] += WIN_POINTS

            table[team_b]["Losses"] += 1

        elif winner == team_b:

            table[team_b]["Wins"] += 1
            table[team_b]["Points"] += WIN_POINTS

            table[team_a]["Losses"] += 1

        elif winner == "Tie":

            table[team_a]["Ties"] += 1
            table[team_b]["Ties"] += 1

            table[team_a]["Points"] += TIE_POINTS
            table[team_b]["Points"] += TIE_POINTS

        elif winner == "No Result":

            table[team_a]["NR"] += 1
            table[team_b]["NR"] += 1

    result = pd.DataFrame(list(table.values()))

    if result.empty:

        return pd.DataFrame(
            columns=[
                "Rank",
                "Team",
                "Played",
                "Wins",
                "Losses",
                "Ties",
                "Points",
                "RunsFor",
                "ActualOversFor",
                "RunsAgainst",
                "ActualOversAgainst",
                "Scored",
                "Conceded",
                "NRR"
            ]
        )

    def calculate_nrr(row):

        if row["NrrOversFor"] == 0 or row["NrrOversAgainst"] == 0:
            return 0.000

        run_rate_for = row["RunsFor"] / row["NrrOversFor"]
        run_rate_against = row["RunsAgainst"] / row["NrrOversAgainst"]

        return round(run_rate_for - run_rate_against, 4)

    result["NRR"] = result.apply(
        calculate_nrr,
        axis=1
    )
    result["NRR"] = result["NRR"].apply(
        lambda x: f"{x:.4f}"
    )


    
    result["Scored"] = (
        result["RunsFor"].astype(int).astype(str)
        + " / "
        + result["ActualOversFor"].apply(decimal_to_cricket_overs)
    )

    result["Conceded"] = (
        result["RunsAgainst"].astype(int).astype(str)
        + " / "
        + result["ActualOversAgainst"].apply(decimal_to_cricket_overs)
    )

    result = result.sort_values(
        by=[
            "Points",
            "NRR",
            "Wins",
            "RunsFor"
        ],
        ascending=[
            False,
            False,
            False,
            False
        ]
    ).reset_index(drop=True)

    result["Rank"] = range(1, len(result) + 1)

    result["ActualOversFor"] = result["ActualOversFor"].round(2)
    result["ActualOversAgainst"] = result["ActualOversAgainst"].round(2)

    result = result[
        [
            "Rank",
            "Team",
            "Played",
            "Wins",
            "Losses",
            "Ties",
            "NR",
            "Points",
            "RunsFor",
            "ActualOversFor",
            "RunsAgainst",
            "ActualOversAgainst",
            "Scored",
            "Conceded",
            "NRR"
        ]
    ]

    return result

def calculate_player_stats(match_history):


    stats = {}

    for _, row in match_history.iterrows():

        if row.get("Status", "Active") == "Deleted":
            continue

        if str(row.get("Winner", "")).strip() == "":
            continue
          
        team_a = row["TeamA"]
        team_b = row["TeamB"]

        runs_a = int(row.get("RunsA", 0) or 0)
        runs_b = int(row.get("RunsB", 0) or 0)

        wickets_a = int(row.get("WicketsA", 0) or 0)
        wickets_b = int(row.get("WicketsB", 0) or 0)

        winner = row["Winner"]

        for player in [team_a, team_b]:

            if player not in stats:

                stats[player] = {
                    "Player": player,
                    "Group": "",
                    "Matches": 0,
                    "Wins": 0,
                    "Losses": 0,
                    "Runs Scored": 0,
                    "Runs Conceded": 0,
                    "Wickets Lost": 0,
                    "Wickets Taken": 0,
                    "Highest Score": 0
                }

        # TEAM A

        stats[team_a]["Matches"] += 1
        stats[team_a]["Runs Scored"] += runs_a
        stats[team_a]["Runs Conceded"] += runs_b
        stats[team_a]["Wickets Lost"] += wickets_a
        stats[team_a]["Wickets Taken"] += wickets_b
        stats[team_a]["Highest Score"] = max(
            stats[team_a]["Highest Score"],
            runs_a
        )

        # TEAM B

        stats[team_b]["Matches"] += 1
        stats[team_b]["Runs Scored"] += runs_b
        stats[team_b]["Runs Conceded"] += runs_a
        stats[team_b]["Wickets Lost"] += wickets_b
        stats[team_b]["Wickets Taken"] += wickets_a
        stats[team_b]["Highest Score"] = max(
            stats[team_b]["Highest Score"],
            runs_b
        )

        if winner == team_a:

            stats[team_a]["Wins"] += 1
            stats[team_b]["Losses"] += 1

        elif winner == team_b:

            stats[team_b]["Wins"] += 1
            stats[team_a]["Losses"] += 1

    #Build Player -> group mapping

    team_group_map = {}

    for group_name, team_list in groups.items():
        for team in team_list:
            team_group_map[team] = group_name

    for player in stats:
        stats[player]["Group"] = team_group_map.get(
            player,
            "Unknown"
        )

    player_df = pd.DataFrame(stats.values())

    if not player_df.empty:

        player_df["Average Score"] = (
            player_df["Runs Scored"] /
            player_df["Matches"]
        ).round(2)

        player_df["Win %"] = (
            player_df["Wins"] /
            player_df["Matches"] * 100
        ).round(1)

    return player_df

# ==================================================
# GENERATE GROUP FIXTURES
# ==================================================

def generate_group_fixtures(team_list):

    fixtures = []

    for team_a, team_b in combinations(team_list, 2):

        fixtures.append({
            "TeamA": team_a,
            "TeamB": team_b
        })

    return pd.DataFrame(fixtures)



# ==================================================
# LOAD MATCH HISTORY AND CURRENT TABLES
# ==================================================

match_history = load_match_entries()

knockout_df = load_knockout_matches()


knockout_history = pd.DataFrame()

if not knockout_df.empty:

    knockout_history = pd.DataFrame({
        "Date": "",
        "Group": knockout_df["Stage"],
        "TeamA": knockout_df["TeamA"],
        "RunsA": knockout_df["RunsA"],
        "WicketsA": knockout_df["WicketsA"],
        "OversA": knockout_df["OversA"],
        "TeamB": knockout_df["TeamB"],
        "RunsB": knockout_df["RunsB"],
        "WicketsB": knockout_df["WicketsB"],
        "OversB": knockout_df["OversB"],
        "Winner": knockout_df["Winner"],
        "Status": "Active"
    })

tournament_history = pd.concat(
    [
        match_history,
        knockout_history
    ],
    ignore_index=True
)

numeric_cols = [
    "RunsA",
    "RunsB",
    "WicketsA",
    "WicketsB",
    "OversA",
    "OversB"
]

for col in numeric_cols:

    tournament_history[col] = pd.to_numeric(
        tournament_history[col],
        errors="coerce"
    ).fillna(0)

required_cols = [
    "Date",
    "Group",
    "TeamA",
    "RunsA",
    "WicketsA",
    "OversA",
    "TeamB",
    "RunsB",
    "WicketsB",
    "OversB",
    "Winner",
    "Status"
]

for col in required_cols:
    if col not in match_history.columns:
        match_history[col] = ""

elite_df = calculate_points_table(
    "Elite",
    groups["Elite"],
    match_history
)

super_df = calculate_points_table(
    "Super",
    groups["Super"],
    match_history
)

golden_df = calculate_points_table(
    "Golden",
    groups["Golden"],
    match_history
)




challenger_df = calculate_points_table(
    "Challenger",
    groups["Challenger"],
    match_history
)

player_stats_df = calculate_player_stats(
    tournament_history
)

#st.write(player_stats_df.columns.tolist())

# ==================================================
# TOP SUMMARY
# ==================================================

c1, c2, c3, c4, right = st.columns(
    [1, 1, 1, 1, 1]
)

with c1:
    st.metric("Elite Teams", len(groups["Elite"]))

with c2:
    st.metric("Super Teams", len(groups["Super"]))

with c3:
    st.metric("Golden Teams", len(groups["Golden"]))

with c4:
    st.metric("Challenger Teams", len(groups["Challenger"]))

# ==================================================
# ADMIN ACCESS REQUEST ALERT
# ==================================================

if st.session_state.get("role") == "Admin":

    try:

        requests_df = load_access_requests() 

        pending_count = len(
            requests_df[
                requests_df["Status"] == "Pending"
            ]
        )

        if pending_count > 0:

            st.warning(
                f"🔔 {pending_count} Pending Request(s) - Open User Management tab to approve."
                )

    except Exception:
        pass



# ==================================================
# DISPLAY GROUP TABLE
# ==================================================

def show_group(title, table_df, color):

    st.markdown(
        f"""
        <div style="
            background-color:{color};
            color:white;
            padding:12px;
            border-radius:8px;
            font-size:24px;
            font-weight:bold;
            margin-top:15px;
            margin-bottom:15px;">
            {title}
        </div>
        """,
        unsafe_allow_html=True
    )

    if table_df.empty:

        st.warning("No teams found for this group.")
        return

    leader = table_df.iloc[0]["Team"]

    m1, m2, m3 = st.columns(3)

    with m1:
        st.metric("Teams", len(table_df))

    with m2:
        st.metric("Leader", leader)



    display_df = table_df[
        [
            "Rank",
            "Team",
            "Played",
            "Wins",
            "Losses",
            "Ties",
            "NR",
            "Points",
            "NRR",
            "Scored",
            "Conceded"
        ]
    ]

    display_df = display_df.rename(
        columns={
            "Played": "P",
            "Wins": "W",
            "Losses": "L",
            "Ties": "T",
            "NR": "NR",
            "Points": "Pts"
        }
    )

    left, center, right = st.columns([1, 3, 1])

    with center:
        #st.table(display_df)

        st.dataframe(
            display_df,
            use_container_width=True,
            hide_index=True,
            height=300
        )




# ==================================================
# TABS
# ==================================================

tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs(
    [
        "🔥 Group Stage",
        "📅 Fixtures",
        "🏆 Knockout Stage",
        "📝 Match Entry",
        "🗑 Delete Match",
        "👑 User Management",
        "📊 Tournament Stats"
    ]
)

# ==================================================
# Group Stage
# ==================================================

with tab1:

    st.subheader("🏏 Group Stage")

    colors = [
        "#1F4E78",  # Blue
        "#198754",  # Green
        "#DAA520",  # Gold
        "#DC3545",  # Red
        "#6F42C1",  # Purple
        "#0DCAF0"   # Cyan
    ]

    for i, group_name in enumerate(groups.keys()):

        group_df = calculate_points_table(
            group_name,
            groups[group_name],
            match_history
        )

        show_group(
            f"🏆 {group_name} Points Table",
            group_df,
            colors[i % len(colors)]
        )
# ==================================================
# FIXTURES
# ==================================================

with tab2:

    st.subheader("📅 Group Stage Fixtures")

    selected_group = st.radio(
        "Select Group",
        list(groups.keys()),
        horizontal=True
    )

    show_pending_only = st.toggle(
        "⏳ Show Pending Matches Only"
    )

    fixtures_df = generate_group_fixtures(
        groups[selected_group]
    )

    if "Group" in match_history.columns:

        played_matches = match_history[
            match_history["Group"] == selected_group
        ]

    else:

        played_matches = pd.DataFrame()


    played_results = {}

    for _, row in played_matches.iterrows():

        match_key = tuple(
            sorted([
                row["TeamA"],
                row["TeamB"]
            ])
        )

        played_results[match_key] = row["Winner"]

    fixtures_df["Status"] = fixtures_df.apply(
        lambda x:
        "✅ Played"
        if tuple(
            sorted([
                x["TeamA"],
                x["TeamB"]
            ])
        ) in played_results
        else "⏳ Pending",
        axis=1
        )

    fixtures_df["Result"] = fixtures_df.apply(
        lambda x:
        f"🏆 {played_results.get(tuple(sorted([x['TeamA'], x['TeamB']])), '')}"
        if tuple(
            sorted([
                x["TeamA"],
                x["TeamB"]
            ])
        ) in played_results
        else "",
        axis=1
    )

    completed = len(
        fixtures_df[
            fixtures_df["Status"] == "✅ Played"
        ]
    )

    pending = len(
        fixtures_df[
            fixtures_df["Status"] == "⏳ Pending"
        ]
    )

    c1, c2 = st.columns(2)

    with c1:
        st.metric(
            "✅ Completed",
            completed
        )

    with c2:
        st.metric(
            "⏳ Pending",
            pending
        )

    if show_pending_only:

        fixtures_df = fixtures_df[
            fixtures_df["Status"] == "⏳ Pending"
        ]

    def highlight_matches(row):

        if row["Status"] == "⏳ Pending":
            return ["background-color: #FFF9C4"] * len(row)

        if row["Status"] == "✅ Played":
            return ["background-color: #E8F5E9"] * len(row)

        return [""] * len(row)


    st.dataframe(
        fixtures_df.style.apply(
            highlight_matches,
            axis=1
        ),
        use_container_width=True,
        hide_index=True
    )




# ==================================================
# Knockout Stage
# ==================================================


with tab3:

    st.subheader("🏆 Knockout Stage")

    st.info(
        "Knockout stage will be activated after group stage completion. " \
        "Top two of each group will fight for Glory"
    )

    st.markdown("### 🏏 Tournament Bracket")

    # ======================================
    # Build Winner Map
    # ======================================

    winner_map = {}

    for _, row in knockout_df.iterrows():

        match = str(row["Match"]).strip()
        winner = str(row["Winner"]).strip()

        if winner:
            winner_map[match] = winner

    def resolve_team_name(team_name):
        match = str(row["Match"]).strip()
        winner = str(row["Winner"]).strip()

        if winner:
            winner_map[match] = winner
    def resolve_team_name(team_name):
        team_name = str(team_name)

        if team_name.startswith("Winner "):

            match_id = team_name.replace(
                "Winner ",
                ""
            ).strip()

            return winner_map.get(
                match_id,
                team_name
            )
        return team_name

    c1, c2, c3 = st.columns(3)

    with c1:
        qf_matches = knockout_df[
                knockout_df["Stage"] == "Quarter Final"
            ]

        for _, row in qf_matches.iterrows():

                st.success(
                    f"""
                    {row['Match']}

                    {row['TeamA']} {row['RunsA']}/{row['WicketsA']}
                    ({row['OversA']})

                    vs

                    {row['TeamB']} {row['RunsB']}/{row['WicketsB']}
                    ({row['OversB']})

                    🏆 Winner: {row['Winner']}
                    """
                )
    with c2:

        sf_df = knockout_df[
            knockout_df["Stage"] == "Semi Final"
        ]

        for _, row in sf_df.iterrows():

            team_a = resolve_team_name(
                row["TeamA"]
            )

            team_b = resolve_team_name(
                row["TeamB"]
            )

            st.info(
                f"""
                {row['Match']}

                {team_a}
                vs
                {team_b}
                """
            )


    with c3:

        final_df = knockout_df[
            knockout_df["Stage"] == "Final"
        ]

        for _, row in final_df.iterrows():

            st.warning(
                f"""
                🏆 {row['Match']}

                {row['TeamA']}
                vs
                {row['TeamB']}
                """
            )
# ==================================================
# MATCH ENTRY TAB
# ==================================================

with tab4:

    if st.session_state.get("role") not in [
        "Admin",
        "Scorekeeper"
    ]:

        st.warning(
            "You do not have permission to add match results. Please login as Admin or Scorekeeper from the User Access section at the bottom."
        )

    else:

        st.markdown(
            """
            <div style="
                background-color:#6C757D;
                color:white;
                padding:12px;
                border-radius:8px;
                font-size:24px;
                font-weight:bold;
                margin-bottom:15px;">
                📝 Enter Match Result
            </div>
            """,
            unsafe_allow_html=True
        )

        st.info(
            "Tournament rule applied: 10 overs maximum, 5 wickets maximum. "
            "If a team loses 5 wickets, NRR uses full 10 overs for calculation, while the table displays actual overs."
        )

        selected_group = st.selectbox(
            "Select Group",
            [
                "Elite",
                "Super",
                "Golden",
                "Challenger"
            ]
        )

        group_teams = groups[selected_group]

        if not group_teams:
            st.error(
                f"No teams configured for {selected_group} in Teams_Master."
            )

        else:

            col1, col2 = st.columns(2)

            with col1:

                team_a = st.selectbox(
                    "Team A",
                    group_teams,
                    key="team_a"
                )

                runs_a = st.number_input(
                    "Runs A",
                    min_value=0,
                    value=0,
                    key="runs_a"
                )

                wickets_a = st.number_input(
                    "Wickets A",
                    min_value=0,
                    max_value=MAX_WICKETS,
                    value=0,
                    key="wickets_a"
                )

                overs_a = st.number_input(
                    "Overs A",
                    min_value=0.0,
                    max_value=float(MAX_OVERS),
                    value=0.0,
                    step=0.1,
                    key="overs_a"
                )

            with col2:

                team_b_options = [
                    team for team in group_teams
                    if team != team_a
                ]

                team_b = st.selectbox(
                    "Team B",
                    team_b_options,
                    key="team_b"
                )

                runs_b = st.number_input(
                    "Runs B",
                    min_value=0,
                    value=0,
                    key="runs_b"
                )

                wickets_b = st.number_input(
                    "Wickets B",
                    min_value=0,
                    max_value=MAX_WICKETS,
                    value=0,
                    key="wickets_b"
                )

                overs_b = st.number_input(
                    "Overs B",
                    min_value=0.0,
                    max_value=float(MAX_OVERS),
                    value=0.0,
                    step=0.1,
                    key="overs_b"
                )

            result_type = st.radio(
                "Result Type",
                ["Normal Result", "No Result"],
                horizontal=True
            )

            if st.button("💾 Save Match Result"):

                if team_a == team_b:

                    st.error("Team A and Team B cannot be the same.")

                elif overs_a <= 0 or overs_b <= 0:

                    st.error("Overs must be greater than 0 for both teams.")

                else:



                    if result_type == "No Result":

                        winner = "No Result"

                    else:

                        if runs_a > runs_b:
                            winner = team_a

                        elif runs_b > runs_a:
                            winner = team_b

                        else:
                            winner = "Tie"

                    match_data = [
                        datetime.now().strftime("%Y-%m-%d %H:%M"),
                        selected_group,
                        team_a,
                        runs_a,
                        wickets_a,
                        overs_a,
                        team_b,
                        runs_b,
                        wickets_b,
                        overs_b,
                        winner,
                        "Active"
                    ]

                    try:

                        save_match(match_data)

                        updated_history = load_match_entries()

                        updated_elite_df = calculate_points_table(
                            "Elite",
                            groups["Elite"],
                            updated_history
                        )

                        updated_super_df = calculate_points_table(
                            "Super",
                            groups["Super"],
                            updated_history
                        )

                        updated_golden_df = calculate_points_table(
                            "Golden",
                            groups["Golden"],
                            updated_history
                        )

                        updated_challenger_df = calculate_points_table(
                            "Challenger",
                            groups["Challenger"],
                            updated_history
                        )

                        #write_calculated_points_to_excel(
                        #    updated_elite_df,
                        #    updated_super_df,
                        #    updated_golden_df,
                        #    updated_challenger_df
                        #)

                        st.success(
                            f"✅ Match saved successfully. Winner: {winner}"
                        )

                        st.rerun()

                    except PermissionError:

                        st.error(
                            "Permission denied. Please close the Excel workbook and then click Save again."
                        )

                    except Exception as e:

                        import traceback

                        st.code(traceback.format_exc())

                        st.error(f"Error while saving match: {e}")

            st.markdown("---")

            st.subheader("Recent Online Match Entries")

            latest_history = load_match_entries()

            if latest_history.empty:

                st.info("No online match entries yet.")

            else:

                active_latest_history = latest_history[
                    latest_history["Status"] != "Deleted"
                ]

                st.dataframe(
                    active_latest_history.tail(20),
                    hide_index=True,
                    use_container_width=True
                )

# ==================================================
# DELETE MATCH TAB
# ==================================================

with tab5:

    if st.session_state.get("role") != "Admin":

        st.warning(
            "Only Admin can delete matches. Please login as Admin from the User Access section at the bottom."
        )

    else:

        st.subheader("🗑 Delete Match")

        history = load_match_entries()

        if history.empty:

            st.info("No match entries found.")

        else:

            if "Status" not in history.columns:
                history["Status"] = "Active"

            active_matches = history[
                history["Status"] != "Deleted"
            ].copy()

            if active_matches.empty:

                st.info("No active matches found.")

            else:

                active_matches["MatchLabel"] = (
                    active_matches["Date"].astype(str)
                    + " | "
                    + active_matches["Group"].astype(str)
                    + " | "
                    + active_matches["TeamA"].astype(str)
                    + " vs "
                    + active_matches["TeamB"].astype(str)
                )

                selected_match = st.selectbox(
                    "Select Match to Delete",
                    active_matches["MatchLabel"]
                )

                confirm_delete = st.checkbox(
                    "I confirm this match should be deleted"
                )

                if confirm_delete and st.button(
                    "🗑 Mark Match as Deleted"
                ):

                    row_index = active_matches[
                        active_matches["MatchLabel"] == selected_match
                    ].index[0]

                    history.loc[
                        row_index,
                        "Status"
                    ] = "Deleted"

                    try:

                        with pd.ExcelWriter(
                            FILE,
                            engine="openpyxl",
                            mode="a",
                            if_sheet_exists="replace"
                        ) as writer:

                            history.to_excel(
                                writer,
                                sheet_name=MATCH_SHEET,
                                index=False
                            )

                        updated_history = load_match_entries()

                        updated_elite_df = calculate_points_table(
                            "Elite",
                            groups["Elite"],
                            updated_history
                        )

                        updated_super_df = calculate_points_table(
                            "Super",
                            groups["Super"],
                            updated_history
                        )

                        updated_golden_df = calculate_points_table(
                            "Golden",
                            groups["Golden"],
                            updated_history
                        )

                        updated_challenger_df = calculate_points_table(
                            "Challenger",
                            groups["Challenger"],
                            updated_history
                        )

                        write_calculated_points_to_excel(
                            updated_elite_df,
                            updated_super_df,
                            updated_golden_df,
                            updated_challenger_df
                        )

                        st.success(
                            "✅ Match marked as Deleted"
                        )

                        st.rerun()

                    except PermissionError:

                        st.error(
                            "Permission denied. Please close the Excel workbook and then try again."
                        )

                    except Exception as e:

                        st.error(f"Error while deleting match: {e}")

# ==================================================
# USER MANAGEMENT
# ==================================================

with tab6:

    if st.session_state.get("role") not in [
        "Admin",
        "Scorekeeper"
    ]:

        st.warning(
            "Only Admin or Scorekeeper can access User Management. Contact admin to submit Match Entry as."
        )

    else:

        st.subheader("👑 User Management")

        users_df = load_users()
        requests_df = load_access_requests()

       

        pending = requests_df[
            requests_df["Status"] == "Pending"
        ]

        pending_count = len(pending)

        if pending_count > 0:
            st.error(
                    f"🔔 {pending_count} Pending Access Request(s)"
            )

        else:
            st.success(
                "✅ No Pending Requests"
            )
        if pending.empty:

            st.info("No pending requests.")

        else:

            st.dataframe(
                pending,
                use_container_width=True
            )

            selected_user = st.selectbox(
                "Select Request",
                pending["Username"]
            )

            role = st.selectbox(
                "Assign Role",
                [
                    "Viewer",
                    "Scorekeeper",
                    "Admin"
                ]
            )

            temp_password = st.text_input(
                "Temporary Password",
                value="Temp123"
            )

            col1, col2 = st.columns(2)

# =====================================
# APPROVE USER
# =====================================

            with col1:

                if st.button("✅ Approve User"):

                    selected_row = pending[
                        pending["Username"] == selected_user
                    ].iloc[0]

                    new_user = pd.DataFrame(
                        [[
                            selected_row["Username"],
                            temp_password,
                            role,
                            "Approved"
                        ]],
                        columns=[
                            "Username",
                            "Password",
                            "Role",
                            "Status"
                        ]
                    )

                    users_df = pd.concat(
                        [users_df, new_user],
                        ignore_index=True
                    )

                    requests_sheet = client.open(
                        "Access"
                    ).worksheet(
                        "Access_Requests"
                    )

                    records = requests_sheet.get_all_records()

                    for i, record in enumerate(records, start=2):
                        if record["Username"] == selected_user:
                            requests_sheet.update_cell(i, 4, "Approved")
                            break
                    users_sheet = client.open(
                        "Access"
                    ).worksheet(
                        "Users"
                    )

                    users_sheet.append_row([
                        selected_row["Username"],
                        temp_password,
                        role,
                        "Approved"
                    ])

                    st.success(
                        f"{selected_user} approved as {role}"
                    )

                    st.rerun()

            # =====================================
            # REJECT USER
            # =====================================

            with col2:

                if st.button("❌ Reject User"):

                    requests_sheet = client.open(
                        "Access"
                    ).worksheet(
                        "Access_Requests"
                    )

                    records = requests_sheet.get_all_records()

                    for i, record in enumerate(records, start=2):
                        if record["Username"] == selected_user:
                            requests_sheet.update_cell(i, 4, "Rejected")
                            break

                    with pd.ExcelWriter(
                        ACCESS_FILE,
                        engine="openpyxl",
                        mode="a",
                        if_sheet_exists="replace"
                    ) as writer:

                        users_df.to_excel(
                            writer,
                            sheet_name="Users",
                            index=False
                        )

                        requests_df.to_excel(
                            writer,
                            sheet_name="Access_Requests",
                            index=False
                        )

                    st.success(
                        f"{selected_user} request rejected"
                    )

                    st.rerun()

# ==================================================
# TOURNAMENT STATS
# ==================================================

with tab7:

    st.subheader("📊 Tournament Statistics")

    if player_stats_df.empty:

        st.info("No statistics available yet.")

    else:

        active_matches = tournament_history[
            tournament_history["Status"] != "Deleted"
        ]

        total_matches = len(active_matches)

        total_runs = (
            active_matches["RunsA"].sum()
            +
            active_matches["RunsB"].sum()
        )

        total_wickets = (
            active_matches["WicketsA"].sum()
            +
            active_matches["WicketsB"].sum()
        )

        average_score = round(
            total_runs / (total_matches * 2),
            2
        ) if total_matches > 0 else 0

        c1, c2, c3, c4 = st.columns(4)

        with c1:
            st.metric(
                "Matches Played",
                total_matches
            )

        with c2:
            st.metric(
                "Total Runs",
                int(total_runs)
            )

        with c3:
            st.metric(
                "Total Wickets",
                int(total_wickets)
            )

        with c4:
            st.metric(
                "Average Score",
                average_score
            )

        st.markdown("### 🏆 Leaderboards")

        orange_cap = player_stats_df.sort_values(
            "Runs Scored",
            ascending=False
        ).iloc[0]

        purple_cap = player_stats_df.sort_values(
            "Wickets Taken",
            ascending=False
        ).iloc[0]

        most_wins = player_stats_df.sort_values(
            "Wins",
            ascending=False
        ).iloc[0]

        highest_score_player = player_stats_df.sort_values(
            "Highest Score",
            ascending=False
        ).iloc[0]

        c1, c2, c3, c4 = st.columns(4)

        with c1:
            st.metric(
                "🟠 Orange Cap",
                orange_cap["Player"],
                int(orange_cap["Runs Scored"])
            )

        with c2:
            st.metric(
                "🟣 Purple Cap",
                purple_cap["Player"],
                int(purple_cap["Wickets Taken"])
            )

        with c3:
            st.metric(
                "🏆 Most Wins",
                most_wins["Player"],
                int(most_wins["Wins"])
            )

        with c4:
            st.metric(
                "💥 Highest Score",
                highest_score_player["Player"],
                int(highest_score_player["Highest Score"])
            )

        st.markdown("### Full Player Statistics")

        display_cols = [
            "Player",
            "Group",
            "Matches",
            "Wins",
            "Losses",
            "Runs Scored",
            "Runs Conceded",
            "Highest Score",
            "Average Score",
            "Wickets Lost",
            "Wickets Taken",
            "Win %"
        ]

        st.dataframe(
            player_stats_df[display_cols].sort_values(
                "Runs Scored",
                ascending=False
            ),
            use_container_width=True,
            hide_index=True
        )



# ==================================================
# LOGIN / REQUEST ACCESS FOOTER
# ==================================================

st.markdown("---")
st.subheader("🔐 User Access")

if st.session_state["logged_in"]:

    col1, col2 = st.columns([4, 1])

    with col1:
        st.success(
            f"✅ Logged in as {st.session_state['username']} "
            f"({st.session_state['role']})"
        )

    with col2:
        if st.button("Logout"):

            st.session_state["logged_in"] = False
            st.session_state["role"] = None
            st.session_state["username"] = None

            st.rerun()

else:

    col1, col2 = st.columns(2)

    # LOGIN
    with col1:

        with st.expander("🔐 User Login", expanded=False):

            username = st.text_input(
                "Username",
                key="login_username"
            )

            password = st.text_input(
                "Password",
                type="password",
                key="login_password"
            )

            if st.button(
                "Login",
                key="login_button"
            ):

                users_df = load_users()

                user_match = users_df[
                    (users_df["Username"] == username)
                    &
                    (users_df["Password"] == password)
                    &
                    (users_df["Status"] == "Approved")
                ]

                if not user_match.empty:

                    st.session_state["logged_in"] = True
                    st.session_state["role"] = user_match.iloc[0]["Role"]
                    st.session_state["username"] = username

                    st.rerun()

                else:

                    st.error("Invalid credentials")

    # REQUEST ACCESS
    with col2:

        with st.expander("📝 Request Access", expanded=False):

            req_username = st.text_input(
                "Username",
                key="request_username"
            )

            req_email = st.text_input(
                "Email",
                key="request_email"
            )

            if st.button(
                "Submit Request",
                key="request_submit"
            ):

                if req_username.strip() == "" or req_email.strip() == "":

                    st.error("Please enter both username and email.")

                else:

                    try:

                        requests_df = load_access_requests()

                        new_row = pd.DataFrame(
                            [[
                                req_username,
                                req_email,
                                datetime.now().strftime(
                                    "%Y-%m-%d %H:%M"
                                ),
                                "Pending"
                            ]],
                            columns=[
                                "Username",
                                "Email",
                                "RequestedOn",
                                "Status"
                            ]
                        )

                        requests_df = pd.concat(
                            [requests_df, new_row],
                            ignore_index=True
                        )

                        sheet = client.open(
                            "Access"
                        ).worksheet(
                            "Access_Requests"
                        )

                        sheet.append_row([
                            req_username,
                            req_email,
                            datetime.now().strftime(
                                "%Y-%m-%d %H:%M"
                            ),
                            "Pending"
                        ])

                        st.success(
                            "✅ Access request submitted"
                        )

                    except Exception as e:

                        st.error(
                            f"Error: {e}"
                        )