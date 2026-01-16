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
# 1. CONFIGURATION & DESIGN "CESI x APPLE"
# ==============================================================================
st.set_page_config(page_title="Modern Opus", page_icon="🎓", layout="wide")

CESI_YELLOW = "#f7e34f" 
CESI_BLACK = "#000000"
APPLE_GRAY = "#F5F5F7"
WHITE = "#FFFFFF"

st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
    
    html, body, [class*="css"] {{
        font-family: 'Inter', -apple-system, sans-serif;
        background-color: {APPLE_GRAY};
        color: {CESI_BLACK};
    }}
    
    .cesi-tag {{
        background-color: {CESI_YELLOW};
        color: {CESI_BLACK};
        font-weight: 800;
        padding: 2px 8px;
        text-transform: uppercase;
        font-size: 0.9rem;
        display: inline-block;
        margin-right: 10px;
    }}

    .cesi-title {{
        font-weight: 800;
        font-size: 2.5rem;
        color: {CESI_BLACK};
        position: relative;
        display: inline-block;
        margin-bottom: 10px;
    }}
    .cesi-title::after {{
        content: '';
        display: block;
        width: 60px;
        height: 6px;
        background-color: {CESI_YELLOW};
        margin-top: 5px;
    }}
    
    .main-header {{
        background-color: {WHITE};
        padding: 2rem 0;
        border-bottom: 1px solid #E5E5E5;
        margin-bottom: 2rem;
        text-align: center;
    }}
    
    .metric-card {{
        background-color: {WHITE};
        border-radius: 12px;
        padding: 1.5rem;
        box-shadow: 0 4px 12px rgba(0,0,0,0.03);
        text-align: center;
        border: 1px solid #EAEAEA;
    }}
    .metric-value {{ font-size: 2rem; font-weight: 800; }}
    .metric-label {{ font-size: 0.85rem; font-weight: 600; color: #86868b; text-transform: uppercase; }}
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 2. LOGIQUE MÉTIER
# ==============================================================================

def normalize_group_label(x):
    if x is None or (isinstance(x, float) and pd.isna(x)): return None
    s = str(x).strip()
    m = re.search(r'G\s*\.?\s*(\d+)', s, re.I)
    return f'G {m.group(1)}' if m else s

def is_time_like(x):
    if isinstance(x, (pd.Timestamp, datetime, time)): return True
    s = str(x).strip()
    return bool(re.match(r'^\d{1,2}[:hH]\d{2}', s)) if s else False

def to_time(x):
    if isinstance(x, time): return x
    try:
        s = str(x).replace('h', ':').replace('H', ':')
        return dtparser.parse(s).time()
    except: return None

def to_date(x):
    if isinstance(x, (date, datetime, pd.Timestamp)):
        return x.date() if hasattr(x, 'date') else x
    try: return dtparser.parse(str(x), dayfirst=True, fuzzy=True).date()
    except: return None

@st.cache_data
def parse_excel_fast(file_content, sheets):
    results = {}
    for sheet in sheets:
        try:
            promo = "P1" if "P1" in sheet.upper() else ("P2" if "P2" in sheet.upper() else sheet)
            df = pd.read_excel(io.BytesIO(file_content), sheet_name=sheet, header=None)
            
            h_rows = [i for i in range(len(df)) if re.match(r'^\s*H\s*\d+', str(df.iat[i,0]), re.I)]
            s_rows = [i for i in range(len(df)) if re.match(r'^\s*S\s*\d+', str(df.iat[i,0]), re.I)]
            
            final_events = []
            for r in h_rows:
                p_cand = [s for s in s_rows if s <= r]
                if not p_cand: continue
                p = max(p_cand)
                d_row, g_row = p+1, p+2
                
                for c in range(df.shape[1]):
                    d = to_date(df.iat[d_row, c])
                    if d is None: continue
                    
                    for col in [c, c+1]:
                        if col >= df.shape[1]: continue
                        summary = df.iat[r, col]
                        if pd.isna(summary) or not str(summary).strip(): continue
                        
                        times = []
                        for off in range(1, 12):
                            if r+off >= len(df): break
                            val = df.iat[r+off, col]
                            if is_time_like(val): times.append(to_time(val))
                        
                        if len(times) < 2: continue
                        
                        teachers = []
                        for off in range(2, 6):
                            if r+off >= len(df): break
                            val = df.iat[r+off, col]
                            if pd.notna(val) and not is_time_like(val) and isinstance(val, str):
                                teachers.append(val.strip())

                        group = normalize_group_label(df.iat[g_row, col]) if g_row < len(df) else None
                        
                        final_events.append({
                            'summary': str(summary).strip(),
                            'start': datetime.combine(d, times[0]),
                            'end': datetime.combine(d, times[-1]),
                            'teachers': sorted(list(set(teachers))),
                            'groups': [group] if group else [],
                            'promo': promo
                        })
            results[promo] = final_events
        except: results[sheet] = []
    return results

def generate_ics(events, include_prefix=False):
    cal = Calendar()
    cal.add('prodid', '-//Modern Opus//FR')
    cal.add('version', '2.0')
    tz = pytz.timezone('Europe/Paris')
    
    for ev in events:
        e = Event()
        prefix = f"[{ev['promo']} {' '.join(ev['groups'])}] " if include_prefix else ""
        e.add('summary', prefix + ev['summary'])
        e.add('dtstart', tz.localize(ev['start']))
        e.add('dtend', tz.localize(ev['end']))
        e.add('description', f"Enseignants: {', '.join(ev['teachers'])}\nGroupes: {', '.join(ev['groups'])}")
        e.add('uid', str(uuid.uuid4()))
        cal.add_component(e)
    return cal.to_ical()

# ==============================================================================
# 3. INTERFACE
# ==============================================================================

st.markdown('<div class="main-header"><div class="cesi-title">Modern Opus</div></div>', unsafe_allow_html=True)
uploaded = st.file_uploader("Déposer l'EDT Excel", type=['xlsx'])

if uploaded:
    bytes_data = uploaded.read()
    xls = pd.ExcelFile(io.BytesIO(bytes_data))
    edt_sheets = [s for s in xls.sheet_names if "EDT" in s.upper()]
    mq_sheet = next((s for s in xls.sheet_names if "MAQUETTE" in s.upper()), None)
    
    with st.spinner("Analyse du fichier..."):
        events_map = parse_excel_fast(bytes_data, edt_sheets)
    
    all_flat = [e for l in events_map.values() for e in l]
    all_teachers = sorted(list(set(t for e in all_flat for t in e['teachers'])))

    # Metrics
    m1, m2, m3 = st.columns(3)
    m1.metric("Séances total", len(all_flat))
    m2.metric("Promos", len(events_map))
    m3.metric("Enseignants", len(all_teachers))

    tab1, tab2, tab3, tab4, tab5 = st.tabs(["🗓️ Calendrier", "✉️ Mails", "📊 Récap", "📥 Exports", "📐 Maquette"])

    with tab1:
        sel_p = st.selectbox("Choisir la promo", list(events_map.keys()), key="p_cal")
        cal_evs = [{"title": e['summary'], "start": e['start'].isoformat(), "end": e['end'].isoformat(), "backgroundColor": CESI_YELLOW, "textColor": "#000"} for e in events_map[sel_p]]
        st_calendar(events=cal_evs, options={"initialView": "timeGridWeek", "locale": "fr", "slotMinTime": "08:00:00", "slotMaxTime": "19:00:00"})

    with tab2:
        t_sel = st.selectbox("Enseignant", all_teachers, key="t_mail")
        if t_sel:
            perso_evs = [e for e in all_flat if t_sel in e['teachers']]
            mail_txt = f"Bonjour,\n\nVoici le récapitulatif de vos interventions :\n\n"
            for ev in perso_evs:
                mail_txt += f"- {ev['start'].strftime('%d/%m')} ({ev['promo']}) : {ev['summary']} de {ev['start'].strftime('%H:%M')} à {ev['end'].strftime('%H:%M')}\n"
            st.text_area("Corps du mail", mail_txt, height=300)

    with tab3:
        p_stat = st.selectbox("Promo", list(events_map.keys()), key="p_stat")
        df_stat = pd.DataFrame([{"Matière": e['summary'], "Date": e['start'].date(), "Durée (h)": (e['end']-e['start']).total_seconds()/3600, "Groupes": ", ".join(e['groups'])} for e in events_map[p_stat]])
        st.dataframe(df_stat, width="stretch")

    with tab4:
        st.markdown('<div class="cesi-tag">Téléchargements</div>', unsafe_allow_html=True)
        col_a, col_b = st.columns(2)
        with col_a:
            st.write("#### Fichiers par Promo")
            for pr, evs in events_map.items():
                if evs:
                    st.download_button(f"📥 {pr}.ics", generate_ics(evs), f"{pr}.ics", key=f"dl_{pr}", mime="text/calendar")
        
        with col_b:
            st.write("#### Fichiers Enseignants")
            sel_multi = st.multiselect("Sélectionner profs", all_teachers, key="multi_exp")
            if sel_multi:
                filt_evs = [e for e in all_flat if any(t in e['teachers'] for t in sel_multi)]
                st.download_button("📥 Télécharger ICS groupé", generate_ics(filt_evs, True), "export_profs.ics", key="dl_profs", mime="text/calendar")

    with tab5:
        if mq_sheet:
            mq_data = pd.read_excel(io.BytesIO(bytes_data), sheet_name=mq_sheet)
            st.dataframe(mq_data, width="stretch")
        else:
            st.warning("Aucun onglet 'Maquette' trouvé.")

else:
    st.markdown("""
    <div style="text-align: center; margin-top: 50px;">
        <h2 style="color: #666;">En attente de fichier...</h2>
        <p>Glissez-déposez votre Excel pour commencer l'analyse.</p>
    </div>
    """, unsafe_allow_html=True)
