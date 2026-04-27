import streamlit as st
import pandas as pd
from docx import Document
from docx.table import Table
from docx.shared import Pt, RGBColor
import io
import re

# ── UI BRANDING
st.set_page_config(page_title="Launchpad Sprint Operations", layout="wide", page_icon="🚀")
st.markdown("""
    <style>
    .stApp { background: linear-gradient(to bottom, #ffffff, #f0f2f6); }
    .section-header { color: #008751; font-size: 1.2rem; font-weight: 700; text-transform: uppercase; margin-bottom: 15px; }
    .stButton>button { background: linear-gradient(135deg, #E65100 0%, #FF9800 100%); color: white; border: none; padding: 12px 30px; font-weight: bold; border-radius: 50px; }
    [data-testid="stMetric"] { background-color: white; padding: 15px; border-radius: 10px; border-left: 5px solid #E65100; box-shadow: 0 2px 5px rgba(0,0,0,0.05); }
    </style>
    """, unsafe_allow_html=True)

with st.sidebar:
    st.image("logooooo-removebg-preview.png", use_container_width=True)
    st.divider()
    st.caption("Admin: Stephanie Nwangwu")

st.markdown("<h1 style='text-align: center; color: #E65100;'>Launchpad Sprint Operations</h1>", unsafe_allow_html=True)
st.divider()

# ── STEP 1: SESSION IDENTITY
st.markdown("<p class='section-header'>Step 1: Session Identity</p>", unsafe_allow_html=True)
c1, c2, c3 = st.columns(3)
with c1:
    week_val  = st.text_input("Week Number", value="2")
    total_val = st.text_input("Total Weeks", value="4")
    topic     = st.text_input("Session Topic", value="Execution Framework")
with c2:
    date_str   = st.text_input("Day & Date", value="Sunday, 12th April 2026")
    time_range = st.text_input("Session Time Range", value="8:00 PM – 11:10 PM")
with c3:
    facil    = st.text_input("Facilitator", value="Temilade Salami")
    dur_text = st.text_input("Total Duration", value="3 hours 10 minutes")

team_input = st.text_area(
    "Team Member List (comma separated)",
    value="Sybil Obeng-Sintim, Paseal Njoku, Janet Isesele, Oluwasanmi Awe, Vivian Nesiama, Edidiong Udoudom, Grace Adu-Yeboah, Stephanie Nwangwu"
)

# ── STEP 2: DATA INGESTION
st.markdown("<p class='section-header'>Step 2: Data Ingestion</p>", unsafe_allow_html=True)
u1, u2, u3 = st.columns(3)
with u1: tracker_file  = st.file_uploader("Official Mentee Tracker (CSV)", type="csv")
with u2: report_file   = st.file_uploader("Google Meet Log (CSV)", type="csv")
with u3: template_file = st.file_uploader("Word Master Template (.docx)", type="docx")

# ── NAME FIX DICTIONARY
# Add new entries here whenever a mentee joins with a different display name
NAME_FIXES = {
    "vickysmane":               "victory chiamaka eze",
    "vickysmane chinedu":       "victory chiamaka eze",
    "dd sylvia":                "chinedu divine favour",
    "chukky igunbor":           "georgina samuel",
    "chukky igbunbor":          "georgina samuel",
    "pearl owusu -twum":        "pearl owusu-twum",
    "pearl owusu  twum":        "pearl owusu-twum",
    "damilola adesodun":        "damilola bankole adesodun",
    "sharon nngwama":           "sharon ngwama",
    "tega ishaya":              "phoebe ishaya oghenetega",
    "phoebe ishaya oghenetega": "phoebe ishaya oghenetega",
    "moyin ojo":                "moyinoluwa ojo",
    "funke":                    "mary oloyede funke",
    "lela":                     "lela tony",
    # NOTE: "divine omeire" removed — she is a different person from chinedu divine favour
}

# ── HELPERS
def convert_to_seconds(t_str):
    if pd.isna(t_str) or t_str == "":
        return 0
    total = 0
    hr = re.search(r'(\d+)\s*hr',  str(t_str))
    mn = re.search(r'(\d+)\s*min', str(t_str))
    sc = re.search(r'(\d+)\s*s',   str(t_str))
    if hr: total += int(hr.group(1)) * 3600
    if mn: total += int(mn.group(1)) * 60
    if sc: total += int(sc.group(1))
    return total

if 'processed' not in st.session_state:
    st.session_state.processed = False

# ── PROCESS BUTTON
if st.button("🚀 PROCESS SESSION DATA"):
    if all([tracker_file, report_file, template_file]):

        team_list = [n.strip() for n in team_input.split(",")]

        tracker = pd.read_csv(tracker_file).dropna(subset=['First Name', 'Last Name'])
        tracker['Full Name'] = (
            tracker['First Name'].str.strip() + " " + tracker['Last Name'].str.strip()
        ).str.lower()

        pod_map     = tracker.set_index('Full Name')['Pod'].to_dict() if 'Pod' in tracker.columns else {}
        mentee_list = set(tracker['Full Name'])

        log = pd.read_csv(report_file).dropna(subset=['Participant Name'])
        log['RawName'] = log['Participant Name'].str.strip().str.lower()

        # ── IMPROVED MATCHING
        def get_best_match(raw_name):
            # 1. Direct fix lookup
            if raw_name in NAME_FIXES:
                return NAME_FIXES[raw_name]
            # 2. Exact match
            if raw_name in mentee_list:
                return raw_name
            # 3. Word overlap — prioritise surname match
            raw_parts = set(raw_name.replace('-', ' ').split())
            raw_last  = raw_name.split()[-1] if raw_name.split() else ""
            best_match, best_score = None, 0
            for m in mentee_list:
                m_parts = set(m.replace('-', ' ').split())
                m_last  = m.split()[-1] if m.split() else ""
                overlap = len(raw_parts & m_parts)
                score   = overlap + (2 if raw_last == m_last else 0)
                if score > best_score and (raw_last == m_last or overlap >= 2):
                    best_score = score
                    best_match = m
            return best_match if best_match else raw_name

        log['Name']    = log['RawName'].apply(get_best_match)
        log['Seconds'] = log['Attended Duration'].apply(convert_to_seconds)

        attendance    = log.groupby('Name')['Seconds'].sum().reset_index()
        mentees_found = attendance[attendance['Name'].isin(mentee_list)]

        # ── DEBUG: show unmatched names so you can add them to NAME_FIXES
        unmatched = attendance[~attendance['Name'].isin(mentee_list)].copy()
        if len(unmatched) > 0:
            st.warning(f"⚠️ {len(unmatched)} Meet name(s) couldn't be matched — add them to NAME_FIXES if they are real mentees:")
            st.dataframe(
                unmatched[['Name']].rename(columns={'Name': 'Unmatched Meet Name'}),
                hide_index=True
            )
        else:
            st.success("✅ All Meet names matched successfully")

        st.session_state.results = {
            "total_p":  len(log['RawName'].unique()),
            "expected": len(mentee_list),
            "present":  len(mentees_found),
            "rate":     f"{round((len(mentees_found) / len(mentee_list)) * 100, 1)}%",
            "high":     sorted(mentees_found[mentees_found['Seconds'] >= 10800]['Name'].str.title().tolist()),
            "mod":      sorted(mentees_found[(mentees_found['Seconds'] >= 7200)  & (mentees_found['Seconds'] < 10800)]['Name'].str.title().tolist()),
            "low":      sorted(mentees_found[(mentees_found['Seconds'] >= 3600)  & (mentees_found['Seconds'] < 7200)]['Name'].str.title().tolist()),
            "absent":   sorted([n.title() for n in mentee_list if n not in set(mentees_found['Name'])]),
            "pod_map":  pod_map,
            "team":     team_list,
        }
        st.session_state.processed = True
    else:
        st.error("Please upload all three files before processing.")

# ── DASHBOARD + REPORT GENERATION
if st.session_state.processed:
    r = st.session_state.results

    st.divider()
    st.markdown("<p class='section-header'>Sprint Dashboard</p>", unsafe_allow_html=True)
    d1, d2, d3, d4 = st.columns(4)
    d1.metric("Total Participants", r['total_p'])
    d2.metric("Expected Mentees",   r['expected'])
    d3.metric("Mentees Present",    r['present'])
    d4.metric("Attendance Rate",    r['rate'])

    st.markdown("<p class='section-header'>Step 3: Operations Observations</p>", unsafe_allow_html=True)
    absent_count = len(r['absent'])
    o_attn = st.text_area("Attendance Observation",
        value=f"{r['present']}/{r['expected']} mentees present.")
    o_eng  = st.text_area("Engagement Observation",
        value="High; majority stayed for the duration.")
    o_abs  = st.text_area("Absentees Note",
        value=f"{absent_count} mentee{'s' if absent_count != 1 else ''} missed the session.")

    if st.button("📝 GENERATE SPRINT REPORT"):
        ORANGE = RGBColor(230, 81, 0)
        GREEN  = RGBColor(0, 135, 81)

        # ── Pod grouping helper
        def get_names_with_pods(name_list):
            pods_dict = {}
            for name in name_list:
                pod = r['pod_map'].get(name.lower(), "Unknown Pod")
                pods_dict.setdefault(pod, []).append(name)
            pods_out  = "\n".join(str(p) for p in pods_dict)
            names_out = "\n".join(", ".join(members) for members in pods_dict.values())
            return pods_out, names_out

        low_pods, low_names = get_names_with_pods(r['low'])
        abs_pods, abs_names = get_names_with_pods(r['absent'])

        replacements = {
            "{{week_number}}":        week_val,
            "{{total_weeks}}":        total_val,
            "{{session_name}}":       topic,
            "{{session_day_date}}":   date_str,
            "{{session_time}}":       time_range,
            "{{total_participants}}": str(r['total_p']),
            "{{total_mentees}}":      str(r['expected']),
            "{{total_present}}":      str(r['present']),
            "{{attendance_rate}}":    r['rate'],
            "{{facilitator}}":        facil,
            "{{session_date}}":       date_str.split(',')[-1].strip() if ',' in date_str else date_str,
            "{{session_duration}}":   f"{dur_text} ({time_range})",
            "{{high_count}}":         str(len(r['high'])),    # number only — template has "mentees" after
            "{{moderate_count}}":     str(len(r['mod'])),     # number only
            "{{low_count}}":          str(len(r['low'])),     # number only
            "{{low_pod}}":            low_pods,
            "{{low_list}}":           low_names,
            "{{absent_pod}}":         abs_pods,
            "{{absent_list}}":        abs_names,
            "{{obs_attendance}}":     o_attn,
            "{{obs_engagement}}":     o_eng,
            "{{obs_absentees}}":      o_abs,
        }

        # ── Run-safe paragraph replacement (preserves bold/colour/size)
        def replace_in_paragraphs(paragraphs):
            for p in paragraphs:
                for k, v in replacements.items():
                    if k in p.text:
                        for run in p.runs:
                            if k in run.text:
                                run.text = run.text.replace(k, str(v))

        # ── Walk ALL tables in XML order (catches tables missed by doc.tables)
        def iter_all_tables(element):
            for child in element:
                tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                if tag == 'tbl':
                    yield Table(child, element)
                else:
                    yield from iter_all_tables(child)

        doc = Document(template_file)

        # ── FIX 1: replace inside Word header (WEEK 2/4 badge lives here)
        for section in doc.sections:
            replace_in_paragraphs(section.header.paragraphs)
            for table in section.header.tables:
                for row in table.rows:
                    seen_row = set()
                    for cell in row.cells:
                        if id(cell._tc) in seen_row: continue
                        seen_row.add(id(cell._tc))
                        replace_in_paragraphs(cell.paragraphs)

        # Body paragraphs
        replace_in_paragraphs(doc.paragraphs)

        # ── FIX 3: deduplicate PER ROW only (not globally) — avoids skipping
        #           cells whose _tc id appears across multiple tables in this template
        for table in iter_all_tables(doc.element.body):
            for row in table.rows:
                seen_row = set()  # reset per row
                for cell in row.cells:
                    if id(cell._tc) in seen_row:
                        continue
                    seen_row.add(id(cell._tc))

                    cell_text = cell.text

                    # Team members — plain list
                    if "{{team_members}}" in cell_text:
                        cell.text = ""
                        for member in r['team']:
                            p   = cell.add_paragraph(member)
                            run = p.runs[0] if p.runs else p.add_run(member)
                            run.font.name = 'Montserrat'
                            run.font.size = Pt(10)
                            run.font.color.rgb = GREEN
                        continue

                    # Name list blocks — check BEFORE replace_in_paragraphs touches the cell
                    list_map = {
                        "{{high_list}}":     (r['high'], 'columns'),
                        "{{moderate_list}}": (r['mod'],  'columns'),
                        "{{low_list}}":      (r['low'],  'plain'),
                        "{{absent_list}}":   (r['absent'], 'plain'),
                    }
                    handled = False
                    for tag, (names, style) in list_map.items():
                        if tag in cell_text:
                            cell.text = ""
                            if style == 'plain':
                                for name in names:
                                    p   = cell.add_paragraph(name)
                                    run = p.runs[0] if p.runs else p.add_run(name)
                                    run.font.name = 'Montserrat'
                                    run.font.size = Pt(10)
                            else:
                                mid  = (len(names) + 1) // 2
                                col1, col2 = names[:mid], names[mid:]
                                for i in range(max(len(col1), len(col2))):
                                    n1   = col1[i] if i < len(col1) else ""
                                    n2   = col2[i] if i < len(col2) else ""
                                    num2 = f"{i+mid+1}." if n2 else ""
                                    line = f"{i+1}. {n1:<35} {num2:<4} {n2}"
                                    p    = cell.add_paragraph(line)
                                    run  = p.runs[0] if p.runs else p.add_run(line)
                                    run.font.name = 'Montserrat'
                                    run.font.size = Pt(9)
                            handled = True
                            break

                    if not handled:
                        replace_in_paragraphs(cell.paragraphs)
                        # Bold + orange the four stats-bar numbers
                        for p in cell.paragraphs:
                            for run in p.runs:
                                for k in ["{{total_participants}}", "{{total_mentees}}",
                                          "{{total_present}}", "{{attendance_rate}}"]:
                                    if run.text == replacements.get(k, "___NO___"):
                                        run.font.bold      = True
                                        run.font.color.rgb = ORANGE
                                        run.font.size      = Pt(18)

        bio = io.BytesIO()
        doc.save(bio)
        st.balloons()
        st.download_button(
            "📥 DOWNLOAD FINAL REPORT",
            data=bio.getvalue(),
            file_name=f"Sprint_Week_{week_val}_Report.docx"
        )