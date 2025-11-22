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
# 1. CONFIGURATION & DESIGN "MODERN OPUS" (Jaune & Noir)
# ==============================================================================
st.set_page_config(page_title="Modern Opus", page_icon="🎹", layout="wide")

st.markdown("""
<style>
    /* Import Font */
    @import url('https://fonts.googleapis.com/css2?family=Oswald:wght@400;600&family=Roboto:wght@300;400;700&display=swap');
    
    /* COULEURS : Noir (#121212), Jaune (#FFD700), Gris Foncé (#1E1E1E) */
    
    html, body, [class*="css"] {
        font-family: 'Roboto', sans-serif;
        color: #E0E0E0;
        background-color: #121212;
    }
    
    h1, h2, h3 {
        font-family: 'Oswald', sans-serif;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    
    /* Header Principal */
    .main-header {
        background: linear-gradient(135deg, #FFD700 0%, #FFA500 100%);
        padding: 2rem;
        border-radius: 0px 0px 20px 20px;
        color: #121212;
        margin-bottom: 2rem;
        text-align: center;
        box-shadow: 0 4px 15px rgba(255, 215, 0, 0.3);
    }
    .main-header h1 {
        color: #121212 !important;
        font-weight: 800;
        font-size: 3rem;
        margin: 0;
    }
    .main-header p {
        color: #121212;
        font-weight: 600;
        font-size: 1.2rem;
    }

    /* Cards Metrics */
    .metric-card {
        background-color: #1E1E1E;
        border-left: 5px solid #FFD700;
        padding: 1.5rem;
        border-radius: 5px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.3);
        text-align: center;
        transition: transform 0.2s;
    }
    .metric-card:hover {
        transform: scale(1.02);
        background-color: #252525;
    }
    .metric-value {
        font-size: 2.5rem;
        font-weight: 800;
        color: #FFD700;
        font-family: 'Oswald', sans-serif;
    }
    .metric-label {
        color: #B0B0B0;
        font-size: 0.9rem;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    
    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 10px;
        background-color: #121212;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        background-color: #1E1E1E;
        border-radius: 5px;
        color: #B0B0B0;
        font-weight: 600;
        border: 1px solid #333;
    }
    .stTabs [aria-selected="true"] {
        background-color: #FFD700;
        color: #121212;
        border: 1px solid #FFD700;
    }

    /* Dataframes */
    [data-testid="stDataFrame"] {
        border: 1px solid #333;
    }
    
    /* Custom Buttons */
    .stButton button {
        background-color: #FFD700;
        color: #121212;
        font-weight: bold;
        border-radius: 5px;
        border: none;
        transition: all 0.3s ease;
    }
    .stButton button:hover {
        background-color: #FFA500;
        color: black;
        box-shadow: 0 0 10px #FFD700;
    }

    /* Warning/Info boxes style */
    .stAlert {
        background-color: #1E1E1E;
        color: #E0E0E0;
        border: 1px solid #333;
    }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 2. MOTEUR DE PARSING ROBUSTE
# ==============================================================================

# --- Utilities ---
def normalize_group_label(x):
    if x is None: return None
    try:
        if pd.isna(x): return None
    except: pass
    s = str(x).strip()
    if not s: return None
    # [cite_start]Regex pour capturer G1, G 1, Groupe 1, etc. [cite: 2, 49]
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

# --- Core Parsing Function (Cached) ---
@st.cache_data(show_spinner=False)
def parse_excel_engine(file_content, sheet_names_to_scan):
    """
    Parsing optimisé. Ne traite que les feuilles valides.
    """
    results = {}
    
    for sheet in sheet_names_to_scan:
        try:
            # Lecture DataFrame
            df = pd.read_excel(io.BytesIO(file_content), sheet_name=sheet, header=None, engine='openpyxl')
            
            # Lecture Merged Cells
            merged_map = get_merged_map_from_bytes(file_content, sheet)
            
            nrows, ncols = df.shape
            # [cite_start]Détection lignes semaines (S xx) et lignes heures (H xx) [cite: 7, 53]
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
                        
                        # Summary
                        try: summary = df.iat[r, col]
                        except: summary = None
                        if pd.isna(summary) or summary is None: continue
                        summary_str = str(summary).strip()
                        if not summary_str: continue
                        
                        # [cite_start]Teachers - Correction Bug "1.222" [cite: 12, 64]
                        teachers = []
                        if (r+2) < nrows:
                            for off in range(2, 6):
                                if (r+off) >= nrows: break
                                try: t = df.iat[r+off, col]
                                except: t = None
                                # On vérifie que ce n'est pas un nombre
                                if t and not pd.isna(t) and not is_time_like(t) and not isinstance(t, (int, float)):
                                    s_t = str(t).strip()
                                    if s_t: teachers.append(s_t)
                        teachers = list(dict.fromkeys(teachers))
                        
                        # Stop Index / Time
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
                        
                        # Description
                        desc_parts = []
                        for idx in range(r+1, stop_idx):
                            if idx >= nrows: break
                            try: cell = df.iat[idx, col]
                            except: cell = None
                            if pd.isna(cell) or cell is None: continue
                            s_cell = str(cell).strip()
                            if not s_cell or to_date(cell) is not None: continue
                            # [cite_start]Exclusion enseignants/summary dans description [cite: 22, 75]
                            if s_cell in teachers or s_cell == summary_str: continue
                            desc_parts.append(s_cell)
                        desc_text = " | ".join(dict.fromkeys(desc_parts))
                        
                        # Start/End Time
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
                        
                        # Date
                        d = to_date(df.iat[date_row, c])
                        if d is None: continue
                        dtstart, dtend = datetime.combine(d, start_t), datetime.combine(d, end_t)
                        
                        # Groups
                        gl = normalize_group_label(df.iat[group_row, col] if group_row < nrows else None)
                        gl_next = normalize_group_label(df.iat[group_row, col+1] if (col+1) < ncols and group_row < nrows else None)
                        
                        groups = set()
                        is_left = (col == c)
                        if is_left:
                            merged_here = merged_map.get((r, col))
                            merged_right = merged_map.get((r, col+1))
                            # [cite_start]Si fusionné horizontalement, c'est les deux groupes [cite: 30, 89]
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
            
            # Fusion des doublons
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
            results[sheet] = final_events
            
        except Exception as e:
            print(f"Erreur parsing {sheet}: {e}")
            results[sheet] = []
            
    return results

# ==============================================================================
# 3. EXPORT & TOOLS
# ==============================================================================

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
# 4. INTERFACE PRINCIPALE
# ==============================================================================

st.markdown('<div class="main-header"><h1>MODERN OPUS</h1><p>L\'Excellence de la Planification.</p></div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("📂 Déposez votre fichier Excel (Format EDT P1/P2)", type=['xlsx'])

if uploaded_file:
    file_bytes = uploaded_file.read()
    
    try:
        xls = pd.ExcelFile(io.BytesIO(file_bytes), engine='openpyxl')
        all_sheets = xls.sheet_names
        # FILTRE STRICT DEMANDÉ : Doit contenir "EDT" ET ("P1" OU "P2")
        promo_sheets = [
            s for s in all_sheets 
            if "EDT" in s.upper() and ("P1" in s.upper() or "P2" in s.upper())
        ]
        # Maquette moins stricte
        maquette_sheet = next((s for s in all_sheets if "maquette" in s.lower()), None)
    except Exception as e:
        st.error(f"Erreur de lecture du fichier: {e}")
        st.stop()

    if not promo_sheets:
        st.error("⛔ Aucune feuille valide détectée. Le fichier doit contenir des feuilles nommées 'EDT P1' ou 'EDT P2'.")
        st.stop()

    with st.spinner('Analyse approfondie des données en cours...'):
        events_map = parse_excel_engine(file_bytes, promo_sheets)

    # Agrégation globale
    all_events_flat = []
    teachers_set = set()
    subjects_set = set()
    exam_events = []

    for sheet, evs in events_map.items():
        for e in evs:
            e['source_sheet'] = sheet # On stocke la source (ex: EDT P1)
            all_events_flat.append(e)
            for t in e['teachers']: teachers_set.add(t)
            subjects_set.add(e['summary'])
            
            # Détection examen
            text_blob = (e['summary'] + " " + e['description']).lower()
            if any(kw in text_blob for kw in ['examen', 'ds ', 'partiel', 'exam ']):
                exam_events.append(e)
    
    # --- DASHBOARD METRICS ---
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown(f'<div class="metric-card"><div class="metric-value">{len(all_events_flat)}</div><div class="metric-label">Séances Totales</div></div>', unsafe_allow_html=True)
    with col2:
        st.markdown(f'<div class="metric-card"><div class="metric-value">{len(promo_sheets)}</div><div class="metric-label">Promos</div></div>', unsafe_allow_html=True)
    with col3:
        st.markdown(f'<div class="metric-card"><div class="metric-value">{len(teachers_set)}</div><div class="metric-label">Enseignants</div></div>', unsafe_allow_html=True)
    with col4:
        st.markdown(f'<div class="metric-card"><div class="metric-value" style="color:white">{len(exam_events)}</div><div class="metric-label">Examens Détectés</div></div>', unsafe_allow_html=True)
    
    st.write("") 

    # --- TABS NAVIGATION ---
    tab_cal, tab_mail, tab_stats, tab_exam, tab_export, tab_maquette = st.tabs([
        "🗓️ Visualisation", 
        "✉️ Générateur de Mails",
        "📊 Récapitulatifs",
        "🎓 Examens",
        "📥 Exports ICS", 
        "📐 Comparatif Maquette"
    ])

    # ====================== TAB 1: CALENDRIER ======================
    with tab_cal:
        col_sel, col_cal = st.columns([1, 4])
        with col_sel:
            st.markdown("### Configuration")
            cal_promo = st.selectbox("Promo", promo_sheets, key="cal_sel")
            view_mode = st.radio("Vue", ["Semaine", "Mois"], index=0)
        
        with col_cal:
            cal_events = []
            for ev in events_map.get(cal_promo, []):
                # Design: Titre propre sans () vides
                teachers_str = ', '.join(ev['teachers'])
                title = ev['summary']
                if teachers_str:
                    title += f" ({teachers_str})"
                
                cal_events.append({
                    "title": title,
                    "start": ev['start'].isoformat(),
                    "end": ev['end'].isoformat(),
                    "backgroundColor": "#FFD700", # Jaune Modern Opus
                    "borderColor": "#B8860B",
                    "textColor": "#000000" # Texte Noir
                })
            
            calendar_options = {
                "headerToolbar": {
                    "left": "today prev,next",
                    "center": "title",
                    "right": ""
                },
                "initialView": "timeGridWeek" if view_mode == "Semaine" else "dayGridMonth",
                "slotMinTime": "07:00:00",
                "slotMaxTime": "20:00:00",
                "allDaySlot": False,
                "contentHeight": "auto",
                "locale": "fr"
            }
            if cal_events:
                st_calendar(events=cal_events, options=calendar_options)
            else:
                st.info("Aucun événement à afficher.")

    # ====================== TAB 2: GÉNÉRATEUR EMAILS (FIXED) ======================
    with tab_mail:
        c_mail1, c_mail2 = st.columns([1, 3])
        
        sorted_teachers = sorted(list(teachers_set))
        
        with c_mail1:
            st.markdown("### Paramètres")
            if not sorted_teachers:
                st.warning("Aucun enseignant.")
                chosen_teacher = None
            else:
                chosen_teacher = st.selectbox("Enseignant", sorted_teachers)
            
            politesse = st.radio("Ton", ["Tutoiement", "Vouvoiement"])
            
        with c_mail2:
            if chosen_teacher:
                # Filtre et Tri
                teacher_evs = [e for e in all_events_flat if chosen_teacher in e['teachers']]
                
                if not teacher_evs:
                    st.warning("Aucun cours trouvé pour cet enseignant.")
                else:
                    # 1. Groupement par Matière
                    grouped_by_subject = {}
                    for ev in teacher_evs:
                        grouped_by_subject.setdefault(ev['summary'], []).append(ev)
                    
                    # Construction du mail
                    parts = chosen_teacher.split()
                    first_name = parts[1] if len(parts) > 1 else chosen_teacher
                    
                    if politesse == "Tutoiement":
                        header = f"Bonjour {first_name},\n\nVoici tes prochaines interventions :"
                        footer = "Bien à toi,"
                    else:
                        header = "Bonjour,\n\nVeuillez trouver ci-dessous le détail de vos interventions :"
                        footer = "Cordialement,"
                    
                    body_content = ""
                    months = {1:'janvier', 2:'février', 3:'mars', 4:'avril', 5:'mai', 6:'juin', 7:'juillet', 8:'août', 9:'septembre', 10:'octobre', 11:'novembre', 12:'décembre'}
                    days = {0:'Lundi', 1:'Mardi', 2:'Mercredi', 3:'Jeudi', 4:'Vendredi', 5:'Samedi', 6:'Dimanche'}

                    for subject, evs_list in grouped_by_subject.items():
                        # Tri des événements de la matière par date
                        evs_list.sort(key=lambda x: x['start'])
                        
                        body_content += f"\nEn {subject} :\n"
                        
                        for ev in evs_list:
                            dt = ev['start']
                            day_str = f"{days[dt.weekday()]} {dt.day} {months[dt.month]} {dt.year}"
                            h_start = dt.strftime("%Hh%M")
                            h_end = ev['end'].strftime("%Hh%M")
                            
                            # Logique Groupe "Intelligente"
                            # Si 2 groupes (G1, G2) souvent présents, c'est la promo entière -> on affiche rien ou Promo
                            # Si un seul groupe spécifique -> on l'affiche
                            # On récupère aussi le nom de la feuille (EDT P1)
                            promo_name = ev['source_sheet']
                            
                            grp_display = ""
                            if ev['groups']:
                                # Si un seul groupe est mentionné, on le précise. Sinon on suppose promo entière
                                if len(ev['groups']) == 1:
                                    grp_display = f" ({ev['groups'][0]})"
                            
                            # Format demandé : dates + heures (promo + groupe si unique)
                            body_content += f"{day_str} de {h_start} à {h_end} ({promo_name}{grp_display})\n"
                    
                    total_h = sum([(e['end'] - e['start']).total_seconds()/3600 for e in teacher_evs])
                    
                    final_mail = f"{header}\n{body_content}\nTotal planifié : {total_h:g} heures.\n\n{footer}"
                    
                    st.text_area("Texte généré (Copier-coller)", value=final_mail, height=500)

    # ====================== TAB 3: RÉCAPITULATIFS (RESTORED) ======================
    with tab_stats:
        st.markdown("### Analyse détaillée")
        stat_mode = st.radio("Type de vue", ["Par Matière", "Par Enseignant"], horizontal=True)
        
        # [cite_start]Helper pour aplatir [cite: 114]
        def flatten_for_display(events_list):
            rows = []
            for ev in events_list:
                rows.append({
                    "Promo": ev['source_sheet'],
                    "Date": ev['start'].strftime("%d/%m/%Y"),
                    "Heure Début": ev['start'].strftime("%H:%M"),
                    "Heure Fin": ev['end'].strftime("%H:%M"),
                    "Matière": ev['summary'],
                    "Enseignant(s)": ", ".join(ev['teachers']),
                    "Groupes": ", ".join(ev['groups'])
                })
            return pd.DataFrame(rows)

        if stat_mode == "Par Matière":
            sel_subj = st.selectbox("Choisir une matière", sorted(list(subjects_set)))
            if sel_subj:
                filtered = [e for e in all_events_flat if e['summary'] == sel_subj]
                st.dataframe(flatten_for_display(filtered), use_container_width=True)
                
        elif stat_mode == "Par Enseignant":
            sel_teach = st.selectbox("Choisir un enseignant", sorted_teachers)
            if sel_teach:
                filtered = [e for e in all_events_flat if sel_teach in e['teachers']]
                st.dataframe(flatten_for_display(filtered), use_container_width=True)

    # ====================== TAB 4: EXAMENS (NEW) ======================
    with tab_exam:
        st.markdown("### 🎓 Calendrier des Examens")
        st.write("Evénements contenant 'examen', 'ds', 'partiel' dans le titre ou la description.")
        
        if not exam_events:
            st.info("Aucun examen détecté.")
        else:
            # Tri par date
            exam_events.sort(key=lambda x: x['start'])
            
            for exam in exam_events:
                dt = exam['start']
                with st.expander(f"{dt.strftime('%d/%m')} - {exam['summary']} ({exam['source_sheet']})"):
                    st.write(f"**Horaire :** {dt.strftime('%H:%M')} - {exam['end'].strftime('%H:%M')}")
                    st.write(f"**Description :** {exam['description']}")
                    st.write(f"**Groupes :** {', '.join(exam['groups']) if exam['groups'] else 'Tous'}")

    # ====================== TAB 5: EXPORTS ======================
    with tab_export:
        c_ex1, c_ex2 = st.columns(2)
        with c_ex1:
            st.markdown("#### Par Promo")
            for sheet in promo_sheets:
                evs = events_map.get(sheet, [])
                if evs:
                    ics = generate_ics(evs)
                    st.download_button(f"📅 {sheet}.ics", data=ics, file_name=f"{sheet}.ics", mime="text/calendar")
        
        with c_ex2:
            st.markdown("#### Par Enseignant")
            sel_t = st.multiselect("Enseignants", sorted_teachers)
            if sel_t:
                f_evs = [e for e in all_events_flat if any(t in e['teachers'] for t in sel_t)]
                if st.button("Générer ICS Enseignant"):
                    ics = generate_ics(f_evs)
                    st.download_button("📅 Télécharger", data=ics, file_name="Planning_Enseignants.ics", mime="text/calendar")

    # ====================== TAB 6: MAQUETTE ======================
    with tab_maquette:
        if not maquette_sheet:
            st.warning("Feuille 'Maquette' introuvable.")
        else:
            # [cite_start]Lecture Maquette simplifiée [cite: 108]
            mq_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=maquette_sheet, header=None, engine='openpyxl')
            rows_mq = []
            if mq_df.shape[1] > 12:
                for i in range(len(mq_df)):
                    subj = mq_df.iat[i, 2] # Col C
                    tgt = mq_df.iat[i, 12] # Col M
                    if pd.notna(subj) and str(subj).strip():
                        try: t_val = float(tgt)
                        except: t_val = 0
                        rows_mq.append({'Matière': str(subj).strip(), 'Cible': t_val})
            
            df_maquette = pd.DataFrame(rows_mq)
            
            comp_promo = st.selectbox("Comparer avec", promo_sheets)
            real_hours = {}
            for ev in events_map.get(comp_promo, []):
                d = (ev['end'] - ev['start']).total_seconds() / 3600
                real_hours[ev['summary']] = real_hours.get(ev['summary'], 0) + d
            
            comp_data = []
            for _, row in df_maquette.iterrows():
                m = row['Matière']
                cible = row['Cible']
                real = real_hours.get(m, 0)
                diff = real - cible
                comp_data.append({"Matière": m, "Prévu": cible, "Réel": round(real, 2), "Ecart": round(diff, 2)})
            
            res_df = pd.DataFrame(comp_data)
            
            # Style Jaune/Noir sur les écarts
            def color_diff(val):
                if val < -2: color = '#FF4B4B' # Rouge Streamlit
                elif val < 0: color = '#FFA500' # Orange
                else: color = '#00C851' # Vert
                return f'color: {color}; font-weight: bold'
            
            st.dataframe(res_df.style.applymap(color_diff, subset=['Ecart']), use_container_width=True)

else:
    st.info("👆 En attente de fichier...")
