import streamlit as st
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook

# ==================================================
# CONFIG
# ==================================================

FILE = "PvP IB Cricket Dashboard.xlsx"

MATCH_SHEET = "Online_Match_Entry"
CALCULATED_POINTS_SHEET = "Calculated_Points_Table"

WIN_POINTS = 2
TIE_POINTS = 1
LOSS_POINTS = 0

# Your tournament rule
MAX_OVERS = 10
MAX_WICKETS = 5

st.set_page_config(
    page_title="PvP IB Cricket Dashboard",
    layout="wide"
)

# ==================================================
# TITLE
# ==================================================

st.title("🏏 PvP IB Cricket Dashboard")

# ==================================================
# READ BASE TEAM STRUCTURE FROM EXCEL
# ==================================================

points_raw = pd.read_excel(
    FILE,
    sheet_name="Points_Tables",
    header=None
)
# ==================================================
# BASE GROUP TEAM LISTS FROM YOUR EXCEL FILE
# ==================================================

teams_df = pd.read_excel(
    FILE,
    sheet_name="Teams_Master"
)

groups = {
    "Elite": teams_df.loc[
        teams_df["Group"] == "Elite",
        "Team"
    ].tolist(),

    "Super": teams_df.loc[
        teams_df["Group"] == "Super",
        "Team"
    ].tolist(),

    "Golden": teams_df.loc[
        teams_df["Group"] == "Golden",
        "Team"
    ].tolist(),

    "Challenger": teams_df.loc[
        teams_df["Group"] == "Challenger",
        "Team"
    ].tolist()
}






all_teams = (
    groups["Elite"] +
    groups["Super"] +
    groups["Golden"] +
    groups["Challenger"]
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

        # Add Group column if old sheet does not have it
        headers = [cell.value for cell in ws[1]]
        if "Status" not in headers:
            ws.cell(row=1, column=12).value = "Status"

        if "Group" not in headers:
            ws.insert_cols(2)
            ws.cell(row=1, column=2).value = "Group"

    ws.append(match_data)

    wb.save(FILE)

# ==================================================
# CRICKET OVERS CONVERSION
# ==================================================
# 9.3 means 9 overs and 3 balls
# Cricket conversion: 9 + 3/6 = 9.5 decimal overs

def convert_overs(value):

    try:
        value = float(value)
    except:
        return 0.0

    whole_overs = int(value)
    decimal_part = round(value - whole_overs, 1)

    balls = int(round(decimal_part * 10, 0))

    if balls >= 0 and balls <= 5:
        return whole_overs + (balls / 6)

    return value

# ==================================================
# ICC STYLE OVERS FOR NRR
# ==================================================
# Your rules:
# 10 overs maximum
# 5 wickets maximum
#
# ICC rule:
# If team is all out, use full quota of overs.
# In your league, all out means 5 wickets.

def get_overs_for_nrr(overs_entered, wickets_lost):

    overs = convert_overs(overs_entered)

    try:
        wickets_lost = int(wickets_lost)
    except:
        wickets_lost = 0

    if wickets_lost >= MAX_WICKETS:
        return float(MAX_OVERS)

    return overs

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
            "OversFor": 0.0,
            "RunsAgainst": 0,
            "OversAgainst": 0.0
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
                    "OversFor",
                    "RunsAgainst",
                    "OversAgainst",
                    "Scored",
                    "Conceded",
                    "NRR"
                ]
            )

        result["NRR"] = 0.000

        result["Scored"] = "0 / 0"
        result["Conceded"] = "0 / 0"

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
                "OversFor",
                "RunsAgainst",
                "OversAgainst",
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

        # ICC-style NRR overs logic for 10-over, 5-wicket league
        overs_a_nrr = get_overs_for_nrr(
            row.get("OversA", 0),
            wickets_a
        )

        overs_b_nrr = get_overs_for_nrr(
            row.get("OversB", 0),
            wickets_b
        )

        winner = row.get("Winner", "")

        # Played
        table[team_a]["Played"] += 1
        table[team_b]["Played"] += 1

        # Runs and overs for Team A
        table[team_a]["RunsFor"] += runs_a
        table[team_a]["RunsAgainst"] += runs_b
        table[team_a]["OversFor"] += overs_a_nrr
        table[team_a]["OversAgainst"] += overs_b_nrr

        # Runs and overs for Team B
        table[team_b]["RunsFor"] += runs_b
        table[team_b]["RunsAgainst"] += runs_a
        table[team_b]["OversFor"] += overs_b_nrr
        table[team_b]["OversAgainst"] += overs_a_nrr

        # Result points
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
                "OversFor",
                "RunsAgainst",
                "OversAgainst",
                "Scored",
                "Conceded",
                "NRR"
            ]
        )
    #st.write("DEBUG GROUP:", group_name)
    #st.write("TEAM LIST:", team_list)
    #st.write(result.head())

    # ICC NRR formula:
    # NRR = (Total Runs Scored / Total Overs Faced)
    #       -
    #       (Total Runs Conceded / Total Overs Bowled)

    def calculate_nrr(row):

        if row["OversFor"] == 0 or row["OversAgainst"] == 0:
            return 0.000

        run_rate_for = row["RunsFor"] / row["OversFor"]
        run_rate_against = row["RunsAgainst"] / row["OversAgainst"]

        return round(run_rate_for - run_rate_against, 3)

    # ==========================================
    # ICC NRR
    # ==========================================


    result["NRR"] = result.apply(
    calculate_nrr,
    axis=1
)

    # ==========================================
    # DISPLAY COLUMNS
    # ==========================================

    result["Scored"] = (
    result["RunsFor"].astype(int).astype(str)
    + " / "
    + result["OversFor"].round(2).astype(str)
)

    result["Conceded"] = (
    result["RunsAgainst"].astype(int).astype(str)
    + " / "
    + result["OversAgainst"].round(2).astype(str)
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

    result["OversFor"] = result["OversFor"].round(2)
    result["OversAgainst"] = result["OversAgainst"].round(2)

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
        "OversFor",
        "RunsAgainst",
        "OversAgainst",
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
# LOAD MATCH HISTORY AND CALCULATE CURRENT TABLES
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

c1, c2, c3, c4 = st.columns(4)

with c1:
    st.metric("Elite Teams", len(groups["Elite"]))

with c2:
    st.metric("Super Teams", len(groups["Super"]))

with c3:
    st.metric("Golden Teams", len(groups["Golden"]))

with c4:
    st.metric("Challenger Teams", len(groups["Challenger"]))

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

    with m3:
        st.metric(
            "Matches Played",
            int(table_df["Played"].sum())
        )

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

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
    [
        "🏆 Elite",
        "⭐ Super",
        "🥇 Golden",
        "🔥 Challenger",
        "📝 Match Entry",
        "🗑 Delete Match"
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
        "If a team loses 5 wickets, NRR uses full 10 overs as per ICC-style all-out logic."
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

        elif wickets_a >= MAX_WICKETS and overs_a > MAX_OVERS:

            st.error("Team A overs cannot be more than 10.")

        elif wickets_b >= MAX_WICKETS and overs_b > MAX_OVERS:

            st.error("Team B overs cannot be more than 10.")

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

                st.info(
                    "Calculated points table updated in Excel sheet: Calculated_Points_Table"
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

        st.dataframe(
            latest_history.tail(20),
            hide_index=True,
            use_container_width=True
        )
# ==================================================
# DELETE MATCH TAB
# ==================================================

with tab6:

    st.subheader("🗑 Delete Match")

    history = pd.read_excel(
        FILE,
        sheet_name=MATCH_SHEET
    )

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
            + active_matches["TeamA"]
            + " vs "
            + active_matches["TeamB"]
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
                active_matches["MatchLabel"]
                == selected_match
            ].index[0]

            history.loc[
                row_index,
                "Status"
            ] = "Deleted"

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

            st.success(
                "✅ Match marked as Deleted"
            )

            st.info(
                "Deleted matches will no longer affect points tables."
            )


            updated_history = pd.read_excel(
            FILE,
            sheet_name=MATCH_SHEET
            )

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

            st.rerun()