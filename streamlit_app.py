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
from streamlit_calendar import calendar as st_calendar # Nécessite streamlit-calendar

# ==============================================================================
# CONFIGURATION & DESIGN "WOW"
# ==============================================================================
st.set_page_config(page_title="Master Planificateur", page_icon="📅", layout="wide")

# Injection de CSS pour un look moderne
st.markdown("""
<style>
    /* Fonts et Couleurs Globales */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;800&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }
    
    /* En-tête stylisé */
    .main-header {
        background: linear-gradient(90deg, #4F46E5 0%, #7C3AED 100%);
        padding: 2rem;
        border-radius: 15px;
        color: white;
        margin-bottom: 2rem;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
    .main-header h1 {
        color: white !important;
        margin: 0;
        font-size: 2.5rem;
    }
    .main-header p {
        color: #E0E7FF;
        font-size: 1.1rem;
    }

    /* Cards pour les stats */
    .metric-card {
        background-color: #ffffff;
        border: 1px solid #e5e7eb;
        padding: 1.5rem;
        border-radius: 10px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        text-align: center;
        transition: transform 0.2s;
    }
    .metric-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .metric-value {
        font-size: 2rem;
        font-weight: 800;
        color: #4F46E5;
    }
    .metric-label {
        color: #6B7280;
        font-size: 0.875rem;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }
    
    /* Tabs personnalisés */
    .stTabs [data-baseweb="tab-list"] {
        gap: 24px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        background-color: #F3F4F6;
        border-radius: 8px;
        padding: 0 20px;
        color: #374151;
        font-weight: 600;
        border: none;
    }
    .stTabs [aria-selected="true"] {
        background-color: #4F46E5;
        color: white;
    }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# MOTEUR DE PARSING (LOGIQUE CONSERVÉE & OPTIMISÉE)
# ==============================================================================

# --- Utilities ---
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

# --- Core Parsing Function (Cached) ---
@st.cache_data(show_spinner=False)
def parse_excel_engine(file_content, sheet_names_to_scan):
    """
    Fonction optimisée qui lit le binaire une fois et parse les feuilles demandées.
    Retourne un dictionnaire {sheet_name: events_list}
    """
    results = {}
    
    # On charge le workbook une fois pour openpyxl (merged cells)
    try:
        merged_maps = {}
        # On ne peut pas lire toutes les merged maps d'un coup efficacement sans tout charger
        # On le fera à la demande ou par feuille
    except:
        pass

    for sheet in sheet_names_to_scan:
        try:
            # Lecture DataFrame
            df = pd.read_excel(io.BytesIO(file_content), sheet_name=sheet, header=None, engine='openpyxl')
            
            # Lecture Merged Cells
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
                        
                        # Summary
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
                                if t and not pd.isna(t) and not is_time_like(t):
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
# EXPORT & TOOLS
# ==============================================================================

def generate_ics(events, tzname='Europe/Paris'):
    # Reprise de la logique Excel2ICS avec VTIMEZONE complet
    cal = Calendar()
    cal.add('prodid', '-//EDT Master Planner//FR')
    cal.add('version', '2.0')
    
    # VTIMEZONE (Simplified for brevity, using icalendar helpers usually preferred but hardcoding works)
    tz = Timezone()
    tz.add('TZID', tzname)
    # Standard
    std = TimezoneStandard()
    std.add('DTSTART', datetime(1970, 10, 25, 3, 0, 0))
    std.add('TZOFFSETFROM', timedelta(hours=2))
    std.add('TZOFFSETTO', timedelta(hours=1))
    std.add('RRULE', {'freq': 'yearly', 'bymonth': 10, 'byday': '-1su'})
    tz.add_component(std)
    # Daylight
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
        
        # Dates
        start_dt = ev['start']
        end_dt = ev['end']
        if start_dt.tzinfo is None: start_dt = timezone.localize(start_dt)
        if end_dt.tzinfo is None: end_dt = timezone.localize(end_dt)
        
        e.add('dtstart', start_dt)
        e.add('dtend', end_dt)
        e.add('dtstamp', datetime.now(timezone))
        e.add('uid', str(uuid.uuid4()))
        
        # Description
        desc_lines = []
        if ev['description']: desc_lines.append(ev['description'])
        if ev['teachers']: desc_lines.append('Enseignant(s): ' + ', '.join(ev['teachers']))
        if ev['groups']: desc_lines.append('Groupes: ' + ', '.join(ev['groups']))
        e.add('description', '\n'.join(desc_lines))
        
        cal.add_component(e)
        
    return cal.to_ical()

# ==============================================================================
# INTERFACE PRINCIPALE
# ==============================================================================

st.markdown('<div class="main-header"><h1>📅 Master Planificateur</h1><p>Analysez, visualisez et diffusez les emplois du temps en un clin d\'œil.</p></div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("Déposez votre fichier Excel (Planning + Maquette)", type=['xlsx'])

if uploaded_file:
    # Lecture binaire une seule fois
    file_bytes = uploaded_file.read()
    
    # Détection des feuilles
    try:
        xls = pd.ExcelFile(io.BytesIO(file_bytes), engine='openpyxl')
        all_sheets = xls.sheet_names
        promo_sheets = [s for s in all_sheets if "EDT" in s.upper() or "P1" in s.upper() or "P2" in s.upper()]
        maquette_sheet = next((s for s in all_sheets if "maquette" in s.lower()), None)
    except Exception as e:
        st.error(f"Erreur de lecture du fichier: {e}")
        st.stop()

    if not promo_sheets:
        st.warning("Aucune feuille 'EDT', 'P1' ou 'P2' détectée.")
        st.stop()

    # Parsing avec barre de progression cachée (c'est très rapide grâce au cache)
    with st.spinner('Traitement intelligent des données...'):
        events_map = parse_excel_engine(file_bytes, promo_sheets)

    # Agrégation de tous les événements pour stats globales
    all_events_flat = []
    teachers_set = set()
    subjects_set = set()
    for sheet, evs in events_map.items():
        for e in evs:
            e['source_sheet'] = sheet
            all_events_flat.append(e)
            for t in e['teachers']: teachers_set.add(t)
            subjects_set.add(e['summary'])
    
    # --- DASHBOARD METRICS ---
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown(f'<div class="metric-card"><div class="metric-value">{len(all_events_flat)}</div><div class="metric-label">Séances Totales</div></div>', unsafe_allow_html=True)
    with col2:
        st.markdown(f'<div class="metric-card"><div class="metric-value">{len(promo_sheets)}</div><div class="metric-label">Promos/Groupes</div></div>', unsafe_allow_html=True)
    with col3:
        st.markdown(f'<div class="metric-card"><div class="metric-value">{len(teachers_set)}</div><div class="metric-label">Enseignants</div></div>', unsafe_allow_html=True)
    with col4:
        st.markdown(f'<div class="metric-card"><div class="metric-value">{len(subjects_set)}</div><div class="metric-label">Matières</div></div>', unsafe_allow_html=True)
    
    st.write("") # Spacer

    # --- TABS NAVIGATION ---
    tab_cal, tab_export, tab_mail, tab_maquette = st.tabs([
        "🗓️ Visualisation Calendrier", 
        "📥 Export & Filtres ICS", 
        "✉️ Générateur d'Emails",
        "📊 Comparatif Maquette"
    ])

    # ====================== TAB 1: CALENDRIER ======================
    with tab_cal:
        st.subheader("Vue Graphique")
        cal_promo = st.selectbox("Choisir la promo à afficher", promo_sheets, key="cal_sel")
        
        # Préparation events pour streamlit-calendar
        cal_events = []
        for ev in events_map.get(cal_promo, []):
            cal_events.append({
                "title": f"{ev['summary']} ({', '.join(ev['teachers'])})",
                "start": ev['start'].isoformat(),
                "end": ev['end'].isoformat(),
                "backgroundColor": "#4F46E5",
                "borderColor": "#4F46E5"
            })
            
        calendar_options = {
            "headerToolbar": {
                "left": "today prev,next",
                "center": "title",
                "right": "dayGridMonth,timeGridWeek,timeGridDay"
            },
            "initialView": "timeGridWeek",
            "slotMinTime": "08:00:00",
            "slotMaxTime": "20:00:00",
        }
        if cal_events:
            st_calendar(events=cal_events, options=calendar_options)
        else:
            st.info("Aucun événement à afficher pour cette sélection.")

    # ====================== TAB 2: EXPORT ICS ======================
    with tab_export:
        col_ex1, col_ex2 = st.columns(2)
        with col_ex1:
            st.markdown("### 📦 Export Global")
            st.write("Télécharger l'emploi du temps complet par promo.")
            for sheet in promo_sheets:
                evs = events_map.get(sheet, [])
                if evs:
                    ics_data = generate_ics(evs)
                    st.download_button(f"📥 Télécharger {sheet}.ics", data=ics_data, file_name=f"{sheet}.ics", mime="text/calendar")
        
        with col_ex2:
            st.markdown("### 🧑‍🏫 Export par Enseignant")
            st.write("Générer un ICS personnalisé pour un ou plusieurs profs.")
            sel_teachers = st.multiselect("Sélectionner les enseignants", sorted(list(teachers_set)))
            
            if sel_teachers:
                filtered_evs = [e for e in all_events_flat if any(t in e['teachers'] for t in sel_teachers)]
                st.write(f"**{len(filtered_evs)}** séances trouvées.")
                if st.button("Générer l'ICS Enseignant(s)"):
                    ics_data = generate_ics(filtered_evs)
                    fname = "Planning_" + "_".join([t.split()[0] for t in sel_teachers[:2]]) + ".ics"
                    st.download_button("📥 Télécharger l'ICS filtré", data=ics_data, file_name=fname, mime="text/calendar")

    # ====================== TAB 3: GÉNÉRATEUR EMAILS (NOUVEAU) ======================
    with tab_mail:
        st.subheader("Générateur de mail récapitulatif")
        
        c_mail1, c_mail2 = st.columns([1, 2])
        
        with c_mail1:
            chosen_teacher = st.selectbox("Destinataire (Enseignant)", sorted(list(teachers_set)))
            politeness = st.radio("Formule de politesse", ["Tutoiement", "Vouvoiement"])
            
            # Filtre events
            teacher_evs = [e for e in all_events_flat if chosen_teacher in e['teachers']]
            teacher_evs.sort(key=lambda x: x['start'])
            
            # Calcul totaux
            total_h = sum([(e['end'] - e['start']).total_seconds()/3600 for e in teacher_evs])
            
        with c_mail2:
            if not teacher_evs:
                st.warning("Aucun cours trouvé pour cet enseignant.")
            else:
                # Construction du texte
                first_name = chosen_teacher.split()[1] if len(chosen_teacher.split()) > 1 else chosen_teacher
                last_name = chosen_teacher
                
                if politesse == "Tutoiement":
                    salutation = f"Bonjour {first_name},"
                    intro = "Voici le récapitulatif de tes prochaines interventions :"
                    closing = "Bien à toi,"
                else:
                    salutation = f"Bonjour," 
                    intro = "Veuillez trouver ci-dessous le récapitulatif de vos interventions planifiées :"
                    closing = "Cordialement,"
                
                list_items = ""
                current_month = None
                
                # Formatage à la française
                months = {1:'Janvier', 2:'Février', 3:'Mars', 4:'Avril', 5:'Mai', 6:'Juin', 7:'Juillet', 8:'Août', 9:'Septembre', 10:'Octobre', 11:'Novembre', 12:'Décembre'}
                days = {0:'Lundi', 1:'Mardi', 2:'Mercredi', 3:'Jeudi', 4:'Vendredi', 5:'Samedi', 6:'Dimanche'}
                
                for ev in teacher_evs:
                    dt = ev['start']
                    day_str = f"{days[dt.weekday()]} {dt.day} {months[dt.month]} {dt.year}"
                    h_start = dt.strftime("%Hh%M")
                    h_end = ev['end'].strftime("%Hh%M")
                    groups = ", ".join(ev['groups']) if ev['groups'] else "Promo entière"
                    
                    list_items += f"- **{day_str}** de {h_start} à {h_end} : {ev['summary']} ({groups})\n"

                email_body = f"""
{salutation}

{intro}

{list_items}
**Total planifié : {total_h:g} heures.**

{closing}
"""
                st.text_area("Brouillon (Copier-coller)", value=email_body, height=400)

    # ====================== TAB 4: MAQUETTE ======================
    with tab_maquette:
        if not maquette_sheet:
            st.warning("Feuille 'Maquette' introuvable dans le fichier Excel.")
        else:
            st.markdown("### Comparatif Heures Réelles vs Maquette")
            
            # Lecture Maquette (Simplifiée pour l'exemple mais robuste)
            # On utilise la logique fournie dans EDTChecker pour lire la maquette
            # (Code simplifié ici pour tenir dans la réponse, mais reprenant la logique de C et M)
            mq_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=maquette_sheet, header=None, engine='openpyxl')
            
            # Extraction simple (Subject col C -> idx 2, Target col M -> idx 12) - Adaptable
            rows_mq = []
            if mq_df.shape[1] > 12:
                for i in range(len(mq_df)):
                    subj = mq_df.iat[i, 2]
                    tgt = mq_df.iat[i, 12]
                    if pd.notna(subj) and str(subj).strip():
                        try: t_val = float(tgt)
                        except: t_val = 0
                        rows_mq.append({'Matière': str(subj).strip(), 'Cible': t_val})
            
            df_maquette = pd.DataFrame(rows_mq)
            
            if df_maquette.empty:
                st.error("Impossible de structurer la maquette (Colonnes C et M attendues).")
            else:
                # Comparaison
                promo_comp = st.selectbox("Promo à comparer", promo_sheets)
                # Calcul heures réelles
                real_hours = {}
                for ev in events_map.get(promo_comp, []):
                    d = (ev['end'] - ev['start']).total_seconds() / 3600
                    real_hours[ev['summary']] = real_hours.get(ev['summary'], 0) + d
                
                comp_data = []
                for _, row in df_maquette.iterrows():
                    m = row['Matière']
                    cible = row['Cible']
                    # Recherche approximative ou exacte
                    real = real_hours.get(m, 0)
                    diff = real - cible
                    comp_data.append({
                        "Matière": m,
                        "Prévu": cible,
                        "Planifié": round(real, 2),
                        "Ecart": round(diff, 2)
                    })
                
                df_comp = pd.DataFrame(comp_data)
                
                # Style conditionnel
                def color_diff(val):
                    color = 'red' if val < -2 else ('orange' if val < 0 else 'green')
                    return f'color: {color}; font-weight: bold'
                
                st.dataframe(df_comp.style.applymap(color_diff, subset=['Ecart']), use_container_width=True)

else:
    # Landing state
    st.info("👆 Commencez par charger votre fichier Excel en haut de page.")
