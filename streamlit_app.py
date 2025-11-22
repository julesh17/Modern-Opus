import streamlit as st
import pandas as pd
import io
import re
import uuid
from datetime import datetime, date, time, timedelta
from dateutil import parser as dtparser
import pytz
from openpyxl import load_workbook
from icalendar import Calendar, Event, Timezone, TimezoneStandard, TimezoneDaylight
from streamlit_calendar import calendar as st_calendar

# ==============================================================================
# 1. CONFIGURATION & DESIGN "APPLE x CESI"
# ==============================================================================
st.set_page_config(page_title="Modern Opus", page_icon="📅", layout="wide")

# Palette CESI & Apple Style
CESI_YELLOW = "#FFC20E" 
CESI_BLACK = "#000000"
APPLE_GRAY = "#F5F5F7"
WHITE = "#FFFFFF"

st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    
    /* Global Reset */
    html, body, [class*="css"] {{
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
        background-color: {APPLE_GRAY};
        color: {CESI_BLACK};
    }}
    
    /* Header Minimaliste */
    .main-header {{
        background-color: {WHITE};
        padding: 1.5rem 0;
        border-bottom: 1px solid #E5E5E5;
        margin-bottom: 2rem;
        text-align: center;
    }}
    .main-header h1 {{
        font-weight: 700;
        font-size: 2.2rem;
        margin: 0;
        letter-spacing: -0.5px;
        color: {CESI_BLACK};
    }}
    
    /* Cards Metrics (Apple Style) */
    .metric-card {{
        background-color: {WHITE};
        border-radius: 12px;
        padding: 1.5rem;
        box-shadow: 0 4px 6px rgba(0,0,0,0.02);
        text-align: center;
        border: 1px solid #EAEAEA;
    }}
    .metric-value {{
        font-size: 2rem;
        font-weight: 700;
        color: {CESI_BLACK};
    }}
    .metric-label {{
        font-size: 0.85rem;
        font-weight: 500;
        color: #86868b; /* Apple Gray Text */
        text-transform: uppercase;
        margin-top: 5px;
    }}
    
    /* Tabs Élégants */
    .stTabs [data-baseweb="tab-list"] {{
        gap: 8px;
        background-color: transparent;
        padding-bottom: 10px;
    }}
    .stTabs [data-baseweb="tab"] {{
        height: 40px;
        background-color: {WHITE};
        border-radius: 8px;
        color: #1d1d1f;
        font-weight: 500;
        border: 1px solid #E5E5E5;
        padding: 0 20px;
        box-shadow: 0 1px 2px rgba(0,0,0,0.05);
    }}
    .stTabs [aria-selected="true"] {{
        background-color: {CESI_YELLOW};
        color: {CESI_BLACK};
        border-color: {CESI_YELLOW};
        font-weight: 600;
    }}

    /* Input & Select boxes */
    .stSelectbox > div > div {{
        border-radius: 8px;
        border: 1px solid #d2d2d7;
    }}

    /* Buttons */
    .stButton button {{
        background-color: {CESI_BLACK};
        color: {WHITE};
        border-radius: 8px;
        font-weight: 500;
        padding: 0.5rem 1rem;
        border: none;
        transition: transform 0.1s ease;
    }}
    .stButton button:hover {{
        background-color: #333333;
        transform: scale(1.02);
        color: {WHITE};
    }}
    
    /* Dataframes clean */
    [data-testid="stDataFrame"] {{
        border: 1px solid #E5E5E5;
        border-radius: 8px;
        overflow: hidden;
    }}
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 2. LOGIQUE MÉTIER (ROBUSTE & SANS RÉGRESSION)
# ==============================================================================

def normalize_group_label(x):
    if x is None: return None
    try:
        if pd.isna(x): return None
    except: pass
    s = str(x).strip()
    if not s: return None
    m = re.search(r'G\s*\.?\s*(\d+)', s, re.I)
    if m: return f'G {m.group(1)}'
    m2 = re.search(r'^(?:groupe)?\s*(\d+)$', s, re.I)
    if m2: return f'G {m2.group(1)}'
    return s

def is_time_like(x):
    if x is None: return False
    if isinstance(x, (pd.Timestamp, datetime, time)): return True
    s = str(x).strip()
    if not s: return False
    if re.match(r'^\d{1,2}[:hH]\d{2}(\s*[AaPp][Mm]\.?)?$', s): return True
    return False

def to_time(x):
    if x is None: return None
    if isinstance(x, time): return x
    if isinstance(x, pd.Timestamp): return x.to_pydatetime().time()
    if isinstance(x, datetime): return x.time()
    s = str(x).strip()
    if not s: return None
    s2 = s.replace('h', ':').replace('H', ':')
    try:
        dt = dtparser.parse(s2, dayfirst=True)
        return dt.time()
    except: return None

def to_date(x):
    if x is None: return None
    if isinstance(x, pd.Timestamp): return x.to_pydatetime().date()
    if isinstance(x, datetime): return x.date()
    if isinstance(x, date): return x
    s = str(x).strip()
    if not s: return None
    try:
        dt = dtparser.parse(s, dayfirst=True, fuzzy=True)
        return dt.date()
    except: return None

def get_merged_map_from_bytes(xls_bytes, sheet_name):
    wb = load_workbook(io.BytesIO(xls_bytes), data_only=True)
    ws = wb[sheet_name]
    merged_map = {}
    for merged in ws.merged_cells.ranges:
        r1, r2, c1, c2 = merged.min_row, merged.max_row, merged.min_col, merged.max_col
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                merged_map[(r - 1, c - 1)] = (r1 - 1, c1 - 1, r2 - 1, c2 - 1)
    return merged_map

@st.cache_data(show_spinner=False)
def parse_excel_engine(file_content, sheet_names_to_scan):
    """ Moteur de parsing unifié """
    results = {}
    
    for sheet in sheet_names_to_scan:
        try:
            # Renommage P1/P2 pour l'affichage
            promo_name = "P1" if "P1" in sheet.upper() else ("P2" if "P2" in sheet.upper() else sheet)

            df = pd.read_excel(io.BytesIO(file_content), sheet_name=sheet, header=None, engine='openpyxl')
            merged_map = get_merged_map_from_bytes(file_content, sheet)
            
            nrows, ncols = df.shape
            s_rows = [i for i in range(len(df)) if isinstance(df.iat[i,0], str) and re.match(r'^\s*S\s*\d+', df.iat[i,0].strip(), re.I)]
            h_rows = [i for i in range(len(df)) if isinstance(df.iat[i,0], str) and re.match(r'^\s*H\s*\d+', df.iat[i,0].strip(), re.I)]
            
            raw_events = []
            
            for r in h_rows:
                p_candidates = [s for s in s_rows if s <= r]
                if not p_candidates: continue
                p = max(p_candidates)
                date_row, group_row = p + 1, p + 2
                
                date_cols = [c for c in range(ncols) if date_row < nrows and to_date(df.iat[date_row, c]) is not None]
                
                for c in date_cols:
                    for col in (c, c + 1):
                        if col >= ncols: continue
                        
                        try: summary = df.iat[r, col]
                        except: summary = None
                        if pd.isna(summary) or summary is None: continue
                        summary_str = str(summary).strip()
                        if not summary_str: continue
                        
                        # Teachers
                        teachers = []
                        if (r+2) < nrows:
                            for off in range(2, 6):
                                if (r+off) >= nrows: break
                                try: t = df.iat[r+off, col]
                                except: t = None
                                if t and not pd.isna(t) and not is_time_like(t) and not isinstance(t, (int, float)):
                                    s_t = str(t).strip()
                                    if s_t: teachers.append(s_t)
                        teachers = list(dict.fromkeys(teachers))
                        
                        stop_idx = None
                        for off in range(1, 12):
                            idx = r + off
                            if idx >= nrows: break
                            try:
                                if is_time_like(df.iat[idx, col]):
                                    stop_idx = idx
                                    break
                            except: continue
                        if stop_idx is None: stop_idx = min(r+7, nrows)
                        
                        desc_parts = []
                        for idx in range(r+1, stop_idx):
                            if idx >= nrows: break
                            try: cell = df.iat[idx, col]
                            except: cell = None
                            if pd.isna(cell) or cell is None: continue
                            s_cell = str(cell).strip()
                            if not s_cell or to_date(cell) is not None: continue
                            if s_cell in teachers or s_cell == summary_str: continue
                            desc_parts.append(s_cell)
                        desc_text = " | ".join(dict.fromkeys(desc_parts))
                        
                        start_val, end_val = None, None
                        for off in range(1, 13):
                            idx = r + off
                            if idx >= nrows: break
                            try: v = df.iat[idx, col]
                            except: v = None
                            if is_time_like(v):
                                if start_val is None: start_val = v
                                elif end_val is None and v != start_val:
                                    end_val = v; break
                        if start_val is None or end_val is None: continue
                        start_t, end_t = to_time(start_val), to_time(end_val)
                        if start_t is None or end_t is None: continue
                        
                        d = to_date(df.iat[date_row, c])
                        if d is None: continue
                        dtstart, dtend = datetime.combine(d, start_t), datetime.combine(d, end_t)
                        
                        gl = normalize_group_label(df.iat[group_row, col] if group_row < nrows else None)
                        gl_next = normalize_group_label(df.iat[group_row, col+1] if (col+1) < ncols and group_row < nrows else None)
                        
                        groups = set()
                        is_left = (col == c)
                        if is_left:
                            merged_here = merged_map.get((r, col))
                            merged_right = merged_map.get((r, col+1))
                            if merged_here and merged_right and merged_here == merged_right:
                                if gl: groups.add(gl)
                                if gl_next: groups.add(gl_next)
                            else:
                                if gl: groups.add(gl)
                        else:
                            if gl: groups.add(gl)
                            
                        raw_events.append({
                            'summary': summary_str,
                            'teachers': set(teachers),
                            'descriptions': set([desc_text]) if desc_text else set(),
                            'start': dtstart,
                            'end': dtend,
                            'groups': groups
                        })
            
            merged = {}
            for e in raw_events:
                key = (e['summary'], e['start'], e['end'])
                if key not in merged:
                    merged[key] = {
                        'summary': e['summary'],
                        'teachers': set(),
                        'descriptions': set(),
                        'start': e['start'],
                        'end': e['end'],
                        'groups': set()
                    }
                merged[key]['teachers'].update(e.get('teachers', set()))
                merged[key]['descriptions'].update(e.get('descriptions', set()))
                merged[key]['groups'].update(e.get('groups', set()))
            
            final_events = []
            for v in merged.values():
                final_events.append({
                    'summary': v['summary'],
                    'teachers': sorted([t for t in v['teachers'] if t and str(t).lower() not in ['nan','none']]),
                    'description': " | ".join(sorted([d for d in v['descriptions'] if d])),
                    'start': v['start'],
                    'end': v['end'],
                    'groups': sorted(list(v['groups']))
                })
            
            results[promo_name] = final_events
            
        except Exception as e:
            results[sheet] = []
            
    return results

def generate_ics(events, tzname='Europe/Paris'):
    cal = Calendar()
    cal.add('prodid', '-//Modern Opus//FR')
    cal.add('version', '2.0')
    
    tz = Timezone()
    tz.add('TZID', tzname)
    std = TimezoneStandard()
    std.add('DTSTART', datetime(1970, 10, 25, 3, 0, 0))
    std.add('TZOFFSETFROM', timedelta(hours=2))
    std.add('TZOFFSETTO', timedelta(hours=1))
    std.add('RRULE', {'freq': 'yearly', 'bymonth': 10, 'byday': '-1su'})
    tz.add_component(std)
    dst = TimezoneDaylight()
    dst.add('DTSTART', datetime(1970, 3, 29, 2, 0, 0))
    dst.add('TZOFFSETFROM', timedelta(hours=1))
    dst.add('TZOFFSETTO', timedelta(hours=2))
    dst.add('RRULE', {'freq': 'yearly', 'bymonth': 3, 'byday': '-1su'})
    tz.add_component(dst)
    cal.add_component(tz)

    timezone = pytz.timezone(tzname)

    for ev in events:
        e = Event()
        e.add('summary', ev['summary'])
        start_dt = ev['start']
        end_dt = ev['end']
        if start_dt.tzinfo is None: start_dt = timezone.localize(start_dt)
        if end_dt.tzinfo is None: end_dt = timezone.localize(end_dt)
        
        e.add('dtstart', start_dt)
        e.add('dtend', end_dt)
        e.add('dtstamp', datetime.now(timezone))
        e.add('uid', str(uuid.uuid4()))
        
        desc_lines = []
        if ev['description']: desc_lines.append(ev['description'])
        if ev['teachers']: desc_lines.append('Enseignant(s): ' + ', '.join(ev['teachers']))
        if ev['groups']: desc_lines.append('Groupes: ' + ', '.join(ev['groups']))
        e.add('description', '\n'.join(desc_lines))
        cal.add_component(e)
        
    return cal.to_ical()

# ==============================================================================
# 3. INTERFACE
# ==============================================================================

st.markdown('<div class="main-header"><h1>Modern Opus</h1></div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("Déposer le fichier Excel", type=['xlsx'])

if uploaded_file:
    file_bytes = uploaded_file.read()
    
    try:
        xls = pd.ExcelFile(io.BytesIO(file_bytes), engine='openpyxl')
        all_sheets = xls.sheet_names
        # Filtre strict : EDT + P1/P2
        promo_sheets = [s for s in all_sheets if "EDT" in s.upper() and ("P1" in s.upper() or "P2" in s.upper())]
        maquette_sheet = next((s for s in all_sheets if "maquette" in s.lower()), None)
    except Exception as e:
        st.error(f"Erreur fichier: {e}")
        st.stop()

    if not promo_sheets:
        st.error("Aucune feuille 'EDT P1' ou 'EDT P2' détectée.")
        st.stop()

    with st.spinner('Chargement...'):
        events_map = parse_excel_engine(file_bytes, promo_sheets)

    # Agrégation
    all_events_flat = []
    teachers_set = set()
    subjects_set = set()
    
    for promo_name, evs in events_map.items():
        for e in evs:
            e['promo_label'] = promo_name
            all_events_flat.append(e)
            for t in e['teachers']: teachers_set.add(t)
            subjects_set.add(e['summary'])
    
    # --- Metrics ---
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(f'<div class="metric-card"><div class="metric-value">{len(all_events_flat)}</div><div class="metric-label">Séances</div></div>', unsafe_allow_html=True)
    with c2:
        st.markdown(f'<div class="metric-card"><div class="metric-value">{len(promo_sheets)}</div><div class="metric-label">Promos</div></div>', unsafe_allow_html=True)
    with c3:
        st.markdown(f'<div class="metric-card"><div class="metric-value">{len(teachers_set)}</div><div class="metric-label">Enseignants</div></div>', unsafe_allow_html=True)
    
    st.write("") 

    tab_cal, tab_mail, tab_stats, tab_exam, tab_export, tab_maquette = st.tabs([
        "🗓️ Calendrier", 
        "✉️ Mails",
        "📊 Récapitulatifs",
        "🎓 Examens",
        "📥 Exports", 
        "📐 Maquette"
    ])

    # --- TAB 1: CALENDRIER ---
    with tab_cal:
        col_sel, col_view = st.columns([1, 4])
        with col_sel:
            cal_promo = st.selectbox("Promo", list(events_map.keys()), key="cal_p")
            view_mode = st.radio("Vue", ["Semaine", "Mois"])
        
        with col_view:
            cal_events = []
            for ev in events_map.get(cal_promo, []):
                cal_events.append({
                    "title": ev['summary'],
                    "start": ev['start'].isoformat(),
                    "end": ev['end'].isoformat(),
                    "backgroundColor": "#FFC20E",
                    "borderColor": "#FFC20E",
                    "textColor": "#000000",
                    "extendedProps": {"description": f"{', '.join(ev['teachers'])}"}
                })
            
            calendar_options = {
                "headerToolbar": {
                    "left": "today prev,next",
                    "center": "title",
                    "right": ""
                },
                "initialView": "timeGridWeek" if view_mode == "Semaine" else "dayGridMonth",
                "slotMinTime": "07:30:00",
                "slotMaxTime": "19:30:00",
                "contentHeight": "auto",
                "locale": "fr",
                "allDaySlot": False,
                "nowIndicator": True,
                "eventBorderColor": "transparent"
            }
            if cal_events:
                st_calendar(events=cal_events, options=calendar_options)

    # --- TAB 2: MAILS ---
    with tab_mail:
        c_m1, c_m2 = st.columns([1, 3])
        sorted_teachers = sorted(list(teachers_set))
        
        with c_m1:
            if sorted_teachers:
                chosen_teacher = st.selectbox("Destinataire", sorted_teachers)
                politesse = st.radio("Ton", ["Tutoiement", "Vouvoiement"])
            else:
                chosen_teacher = None
                st.info("Pas d'enseignants.")
        
        with c_m2:
            if chosen_teacher:
                t_evs = [e for e in all_events_flat if chosen_teacher in e['teachers']]
                clean_name = chosen_teacher.replace(',', '')
                parts = clean_name.split()
                
                if politesse == "Tutoiement":
                    prenom = parts[1] if len(parts) > 1 else parts[0]
                    intro = f"Bonjour {prenom},\n\nVoici le récapitulatif de tes interventions :"
                    closing = "Bien à toi,"
                else:
                    nom = parts[0]
                    intro = f"Bonjour M./Mme {nom},\n\nVeuillez trouver ci-dessous le récapitulatif de vos interventions :"
                    closing = "Cordialement,"
                
                body = ""
                months = {1:'janvier', 2:'février', 3:'mars', 4:'avril', 5:'mai', 6:'juin', 7:'juillet', 8:'août', 9:'septembre', 10:'octobre', 11:'novembre', 12:'décembre'}
                days = {0:'Lundi', 1:'Mardi', 2:'Mercredi', 3:'Jeudi', 4:'Vendredi', 5:'Samedi', 6:'Dimanche'}

                promos_present = sorted(list(set(e['promo_label'] for e in t_evs)))
                total_h_global = 0
                
                for promo in promos_present:
                    p_evs = [e for e in t_evs if e['promo_label'] == promo]
                    if not p_evs: continue
                    
                    body += f"\nPour la promo **{promo}** :\n"
                    by_subj = {}
                    for e in p_evs:
                        by_subj.setdefault(e['summary'], []).append(e)
                    
                    for subj, ev_list in by_subj.items():
                        ev_list.sort(key=lambda x: x['start'])
                        body += f"\nMatière : {subj}\n"
                        for ev in ev_list:
                            dt = ev['start']
                            day_str = f"{days[dt.weekday()]} {dt.day} {months[dt.month]}"
                            h_start = dt.strftime("%Hh%M")
                            h_end = ev['end'].strftime("%Hh%M")
                            dur = (ev['end'] - ev['start']).total_seconds()/3600
                            total_h_global += dur
                            
                            grp_txt = ""
                            if ev['groups'] and len(ev['groups']) == 1:
                                grp_txt = f" ({ev['groups'][0]})"
                            
                            body += f"- {day_str} de {h_start} à {h_end}{grp_txt}\n"
                
                final_txt = f"{intro}\n{body}\nTotal planifié : {total_h_global:g} heures.\n\n{closing}"
                st.text_area("Aperçu", value=final_txt, height=500)

    # --- TAB 3: RÉCAPITULATIFS ---
    with tab_stats:
        st.markdown("### Analyse")
        tabs_promo = st.tabs(list(events_map.keys()))
        
        for i, promo in enumerate(events_map.keys()):
            with tabs_promo[i]:
                evs_promo = events_map[promo]
                mode = st.radio(f"Vue {promo}", ["Par Matière", "Par Enseignant"], key=f"rad_{promo}", horizontal=True)
                
                if mode == "Par Matière":
                    subjs = sorted(list(set(e['summary'] for e in evs_promo)))
                    sel = st.selectbox(f"Matière ({promo})", subjs, key=f"sb_m_{promo}")
                    if sel:
                        data = []
                        for e in evs_promo:
                            if e['summary'] == sel:
                                data.append({
                                    "Date": e['start'].strftime("%d/%m/%Y"),
                                    "Heures": f"{e['start'].strftime('%H:%M')} - {e['end'].strftime('%H:%M')}",
                                    "Enseignant": ", ".join(e['teachers']),
                                    "Groupes": ", ".join(e['groups'])
                                })
                        st.dataframe(pd.DataFrame(data), use_container_width=True)
                
                else:
                    teachs = sorted(list(set(t for e in evs_promo for t in e['teachers'])))
                    sel = st.selectbox(f"Enseignant ({promo})", teachs, key=f"sb_t_{promo}")
                    if sel:
                        data = []
                        for e in evs_promo:
                            if sel in e['teachers']:
                                data.append({
                                    "Date": e['start'].strftime("%d/%m/%Y"),
                                    "Heures": f"{e['start'].strftime('%H:%M')} - {e['end'].strftime('%H:%M')}",
                                    "Matière": e['summary'],
                                    "Groupes": ", ".join(e['groups'])
                                })
                        st.dataframe(pd.DataFrame(data), use_container_width=True)

    # --- TAB 4: EXAMENS ---
    with tab_exam:
        st.markdown("### 🎓 Calendrier des Examens")
        p1_exams = []
        p2_exams = []
        
        for e in all_events_flat:
            if e['description'] and "EXAMEN" in e['description'].upper():
                row = {
                    "Date": e['start'].strftime("%d/%m/%Y"),
                    "Horaire": f"{e['start'].strftime('%H:%M')} - {e['end'].strftime('%H:%M')}",
                    "Matière": e['summary'],
                    "Description": e['description']
                }
                if e['promo_label'] == "P1":
                    p1_exams.append(row)
                elif e['promo_label'] == "P2":
                    p2_exams.append(row)
        
        c_ex1, c_ex2 = st.columns(2)
        with c_ex1:
            st.markdown("#### Promo P1")
            if p1_exams:
                st.dataframe(pd.DataFrame(p1_exams), hide_index=True)
            else:
                st.info("Aucun examen détecté (mot-clé 'EXAMEN' dans description).")
        
        with c_ex2:
            st.markdown("#### Promo P2")
            if p2_exams:
                st.dataframe(pd.DataFrame(p2_exams), hide_index=True)
            else:
                st.info("Aucun examen détecté.")

    # --- TAB 5: EXPORTS ---
    with tab_export:
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("#### Par Promo")
            for promo, evs in events_map.items():
                if evs:
                    ics = generate_ics(evs)
                    st.download_button(f"📥 {promo}.ics", data=ics, file_name=f"{promo}.ics", mime="text/calendar")
        with c2:
            st.markdown("#### Par Enseignant")
            if sorted_teachers:
                sel = st.multiselect("Choix", sorted_teachers, key="exp_sel")
                if st.button("Générer"):
                    evs = [e for e in all_events_flat if any(t in e['teachers'] for t in sel)]
                    ics = generate_ics(evs)
                    st.download_button("📥 Planning_Perso.ics", data=ics, file_name="planning.ics", mime="text/calendar")

    # --- TAB 6: MAQUETTE ---
    with tab_maquette:
        if maquette_sheet:
            mq_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=maquette_sheet, header=None, engine='openpyxl')
            rows_mq = []
            if mq_df.shape[1] > 12:
                for i in range(len(mq_df)):
                    subj = mq_df.iat[i, 2]
                    tgt = mq_df.iat[i, 12]
                    if pd.notna(subj) and str(subj).strip():
                        try: val = float(tgt)
                        except: val = 0
                        rows_mq.append({'Matière': str(subj).strip(), 'Cible': val})
            df_mq = pd.DataFrame(rows_mq)
            
            p_comp = st.selectbox("Comparer Promo", list(events_map.keys()))
            
            real = {}
            for e in events_map.get(p_comp, []):
                dur = (e['end'] - e['start']).total_seconds()/3600
                real[e['summary']] = real.get(e['summary'], 0) + dur
            
            res = []
            for _, r in df_mq.iterrows():
                m = r['Matière']
                c = r['Cible']
                r_val = real.get(m, 0)
                res.append({"Matière": m, "Prévu": c, "Réel": round(r_val, 2), "Ecart": round(r_val - c, 2)})
            
            df_res = pd.DataFrame(res)
            
            def style_ecart(v):
                if v < -2: return 'color: red; font-weight: bold'
                if v < 0: return 'color: orange'
                return 'color: green'
                
            st.dataframe(df_res.style.applymap(style_ecart, subset=['Ecart']), use_container_width=True)
        else:
            st.warning("Pas de feuille Maquette.")

else:
    # État initial (sans fichier)
    st.markdown("""
    <div style="text-align: center; margin-top: 50px; padding: 40px; background-color: #FFFFFF; border-radius: 12px; border: 1px solid #E5E5E5;">
        <h3 style="color: #000000; margin-bottom: 20px;">👋 Bienvenue sur Modern Opus</h3>
        <p style="color: #666666; font-size: 1.1em;">
            Pour commencer, déposez votre fichier Excel d'emploi du temps ci-dessus.<br>
            L'application détectera automatiquement les promos (P1/P2), les enseignants et les examens.
        </p>
        <div style="margin-top: 30px; font-size: 0.9em; color: #888;">
            Formats supportés : .xlsx avec feuilles "EDT P1" / "EDT P2"
        </div>
    </div>
    """, unsafe_allow_html=True)
