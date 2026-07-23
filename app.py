import streamlit as st
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
import os

# ==================================================
# PAGE CONFIG
# ==================================================

st.set_page_config(
    page_title="PvP IB Cricket Dashboard",
    layout="wide"
)

# ==================================================
# CONFIG
# ==================================================

MATCH_SHEET = "Online_Match_Entry"
CALCULATED_POINTS_SHEET = "Calculated_Points_Table"

WIN_POINTS = 2
TIE_POINTS = 1
LOSS_POINTS = 0

MAX_OVERS = 10
MAX_WICKETS = 5
ACCESS_FILE = "Access.xlsx"

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
        return pd.read_excel(
            ACCESS_FILE,
            sheet_name="Users"
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
        return pd.read_excel(
            ACCESS_FILE,
            sheet_name="Access_Requests"
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


season_files = {
    "Season 2": "PvP IB Cricket Dashboard - Season 2.xlsx",
    "Season 3": "PvP IB Cricket Dashboard - Season 3.xlsx"
}

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
        list(season_files.keys()),
        label_visibility="collapsed"
    )

FILE = season_files[season]

if not os.path.exists(FILE):
    st.error(
        f"Excel file not found: {FILE}. Please create this file in the same folder as app.py."
    )
    st.stop()

# ==================================================
# LOAD TEAMS MASTER
# ==================================================

try:
    teams_df = pd.read_excel(
        FILE,
        sheet_name="Teams_Master"
    )

    teams_df.columns = teams_df.columns.str.strip()

except Exception as e:
    st.error(
        f"Unable to read Teams_Master sheet from {FILE}. Error: {e}"
    )
    st.stop()

required_team_columns = ["Group", "Team"]

for col in required_team_columns:
    if col not in teams_df.columns:
        st.error(
            f"Column '{col}' is missing in Teams_Master sheet."
        )
        st.stop()

groups = {
    "Elite": teams_df.loc[
        teams_df["Group"] == "Elite",
        "Team"
    ].dropna().tolist(),

    "Super": teams_df.loc[
        teams_df["Group"] == "Super",
        "Team"
    ].dropna().tolist(),

    "Golden": teams_df.loc[
        teams_df["Group"] == "Golden",
        "Team"
    ].dropna().tolist(),

    "Challenger": teams_df.loc[
        teams_df["Group"] == "Challenger",
        "Team"
    ].dropna().tolist()
}

all_teams = (
    groups["Elite"]
    + groups["Super"]
    + groups["Golden"]
    + groups["Challenger"]
)

# ==================================================
# LOAD ONLINE MATCH ENTRIES
# ==================================================

def load_match_entries():

    try:
        history = pd.read_excel(
            FILE,
            sheet_name=MATCH_SHEET
        )

        if "Status" not in history.columns:
            history["Status"] = "Active"

        return history

    except Exception:

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

# ==================================================
# SAVE MATCH ENTRY
# ==================================================

def save_match(match_data):

    wb = load_workbook(FILE)

    if MATCH_SHEET not in wb.sheetnames:

        ws = wb.create_sheet(MATCH_SHEET)

        ws.append([
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
        ])

    else:

        ws = wb[MATCH_SHEET]

        headers = [cell.value for cell in ws[1]]

        if "Group" not in headers:
            ws.insert_cols(2)
            ws.cell(row=1, column=2).value = "Group"

        headers = [cell.value for cell in ws[1]]

        if "Status" not in headers:
            ws.cell(row=1, column=12).value = "Status"

    ws.append(match_data)

    wb.save(FILE)

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
            table[team_b]["Points"] += LOSS_POINTS

        elif winner == team_b:

            table[team_b]["Wins"] += 1
            table[team_b]["Points"] += WIN_POINTS

            table[team_a]["Losses"] += 1
            table[team_a]["Points"] += LOSS_POINTS

        else:

            table[team_a]["Ties"] += 1
            table[team_b]["Ties"] += 1

            table[team_a]["Points"] += TIE_POINTS
            table[team_b]["Points"] += TIE_POINTS

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

        return round(run_rate_for - run_rate_against, 3)

    result["NRR"] = result.apply(
        calculate_nrr,
        axis=1
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

# ==================================================
# WRITE CALCULATED POINTS TABLE BACK TO EXCEL
# ==================================================

def write_calculated_points_to_excel(
    elite_df,
    super_df,
    golden_df,
    challenger_df
):

    wb = load_workbook(FILE)

    if CALCULATED_POINTS_SHEET in wb.sheetnames:
        del wb[CALCULATED_POINTS_SHEET]

    ws = wb.create_sheet(CALCULATED_POINTS_SHEET)

    def write_group(title, dataframe, start_row):

        ws.cell(row=start_row, column=1).value = title

        headers = list(dataframe.columns)

        for col_index, header in enumerate(headers, start=1):
            ws.cell(row=start_row + 1, column=col_index).value = header

        for row_index, row in enumerate(
            dataframe.itertuples(index=False),
            start=start_row + 2
        ):
            for col_index, value in enumerate(row, start=1):
                ws.cell(row=row_index, column=col_index).value = value

        return start_row + len(dataframe) + 4

    row_position = 1

    row_position = write_group(
        "Elite Points Table",
        elite_df,
        row_position
    )

    row_position = write_group(
        "Super Points Table",
        super_df,
        row_position
    )

    row_position = write_group(
        "Golden Points Table",
        golden_df,
        row_position
    )

    row_position = write_group(
        "Challenger Points Table",
        challenger_df,
        row_position
    )

    wb.save(FILE)

# ==================================================
# LOAD MATCH HISTORY AND CURRENT TABLES
# ==================================================

match_history = load_match_entries()

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

        requests_df = pd.read_excel(
            ACCESS_FILE,
            sheet_name="Access_Requests"
        )

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
            "Points",
            "Scored",
            "Conceded",
            "NRR"
        ]
    ]

    left, center, right = st.columns([1, 3, 1])

    with center:
        st.table(display_df)

# ==================================================
# TABS
# ==================================================

tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs(
    [
        "🏆 Elite",
        "⭐ Super",
        "🥇 Golden",
        "🔥 Challenger",
        "📝 Match Entry",
        "🗑 Delete Match",
        "👑 User Management"
    ]
)

# ==================================================
# ELITE TAB
# ==================================================

with tab1:

    show_group(
        "🏆 Elite Points Table",
        elite_df,
        "#1F4E78"
    )

# ==================================================
# SUPER TAB
# ==================================================

with tab2:

    show_group(
        "⭐ Super Points Table",
        super_df,
        "#198754"
    )

# ==================================================
# GOLDEN TAB
# ==================================================

with tab3:

    show_group(
        "🥇 Golden Points Table",
        golden_df,
        "#DAA520"
    )

# ==================================================
# CHALLENGER TAB
# ==================================================

with tab4:

    show_group(
        "🔥 Challenger Points Table",
        challenger_df,
        "#DC3545"
    )

# ==================================================
# MATCH ENTRY TAB
# ==================================================

with tab5:

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

            if st.button("💾 Save Match Result"):

                if team_a == team_b:

                    st.error("Team A and Team B cannot be the same.")

                elif overs_a <= 0 or overs_b <= 0:

                    st.error("Overs must be greater than 0 for both teams.")

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

                        write_calculated_points_to_excel(
                            updated_elite_df,
                            updated_super_df,
                            updated_golden_df,
                            updated_challenger_df
                        )

                        st.success(
                            f"✅ Match saved successfully. Winner: {winner}"
                        )

                        st.rerun()

                    except PermissionError:

                        st.error(
                            "Permission denied. Please close the Excel workbook and then click Save again."
                        )

                    except Exception as e:

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

with tab6:

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

with tab7:

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

        st.markdown("### Pending Requests")

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

        requests_df.loc[
            requests_df["Username"] == selected_user,
            "Status"
        ] = "Approved"

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
            f"{selected_user} approved as {role}"
        )

        st.rerun()

# =====================================
# REJECT USER
# =====================================

with col2:

    if st.button("❌ Reject User"):

        requests_df.loc[
            requests_df["Username"] == selected_user,
            "Status"
        ] = "Rejected"

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

                        with pd.ExcelWriter(
                            ACCESS_FILE,
                            engine="openpyxl",
                            mode="a",
                            if_sheet_exists="replace"
                        ) as writer:

                            requests_df.to_excel(
                                writer,
                                sheet_name="Access_Requests",
                                index=False
                            )

                        st.success(
                            "✅ Access request submitted"
                        )

                    except Exception as e:

                        st.error(
                            f"Error: {e}"
                        )