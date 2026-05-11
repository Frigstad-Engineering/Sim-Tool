"""
W90 Monte Carlo – Streamlit App
Converted from w90_monte_carlo_v6.ipynb
Sections 1-5 fully implemented.
"""

import streamlit as st
import os, io, math, time, uuid, glob, zipfile, shutil

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

try:
    import cdsapi
except ImportError:
    cdsapi = None

# ─────────────────────────────────────────────────────────────────────────────
#  PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="W90 Monte Carlo", page_icon="🌊", layout="wide")
st.title("🌊 W90 Monte Carlo – Offshore Wind Installation Simulator")
st.markdown("---")

# ─────────────────────────────────────────────────────────────────────────────
#  SECTION 1 – INPUTS
# ─────────────────────────────────────────────────────────────────────────────
with st.expander("⚙️  1. Project Inputs", expanded=True):
    col1, col2, col3 = st.columns(3)
    with col1:
        st.subheader("Project")
        N_WTG          = st.number_input("Number of turbines (N_WTG)", min_value=1, value=50, step=1)
        Simulations    = st.number_input("Simulations", min_value=10, max_value=5000, value=500, step=50)
        Sim_start_date = st.text_input("Sim start date (D-M-YYYY)", value="1-7-2027")
        API_key        = st.text_input("ERA5 API key", value="", type="password")
    with col2:
        st.subheader("Project Location (ERA5 0.25° grid)")
        lat       = st.number_input("Latitude",        value=35.25,  step=0.25, format="%.2f")
        lon       = st.number_input("Longitude",       value=140.75, step=0.25, format="%.2f")
        lat_barge = st.number_input("Barge Latitude",  value=56.50,  step=0.25, format="%.2f")
        lon_barge = st.number_input("Barge Longitude", value=-2.50,  step=0.25, format="%.2f")
    with col3:
        st.subheader("Vessel & Distance")
        dock_distance  = st.number_input("Dock transit distance (Nm)",  value=185, step=5)
        barge_distance = st.number_input("Barge transit distance (Nm)", value=115, step=5)
        W90_carry_cap  = st.number_input("W90 carry capacity",  value=8, step=1)
        JUV_carry_cap  = st.number_input("JUV carry capacity",  value=6, step=1)
        FIV_carry_cap  = st.number_input("FIV carry capacity",  value=5, step=1)

    col4, col5 = st.columns(2)
    with col4:
        st.subheader("Transit Speeds (kn)")
        W90_transit_speed  = st.number_input("W90",                  value=10, step=1)
        Comp_transit_speed = st.number_input("Competitor (JUV/FIV)", value=10, step=1)
        Fiv_transit_speed  = st.number_input("FIV",                  value=10, step=1)
    with col5:
        st.subheader("Operations")
        max_wave          = st.number_input("Max wave height (m)",          value=6,  step=1)
        intermediate_days = st.number_input("Intermediate days (MP → WTG)", value=14, step=1)
        W90_sim = st.checkbox("Include W90",   value=True)
        barge   = st.checkbox("Include Barge", value=True)
        JUV_sim = st.checkbox("Include JUV",   value=True)
        FIV_sim = st.checkbox("Include FIV",   value=True)

    st.subheader("Excel Sheet Names")
    c1, c2 = st.columns(2)
    with c1:
        W90_MP_sheet  = st.text_input("W90 MP sheet",  value="W90 MP")
        JUV_MP_sheet  = st.text_input("JUV MP sheet",  value="JUV MP")
        FIV_MP_sheet  = st.text_input("FIV MP sheet",  value="FIV MP")
    with c2:
        W90_WTG_sheet = st.text_input("W90 WTG sheet", value="W90 WTG")
        JUV_WTG_sheet = st.text_input("JUV WTG sheet", value="JUV WTG")
        FIV_WTG_sheet = st.text_input("FIV WTG sheet", value="FIV WTG")

htd = 1 / 24

# ─────────────────────────────────────────────────────────────────────────────
#  SECTION 1.3 – UPLOAD OPERATIONS EXCEL
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("---")
st.header("📂  1.3  Upload Operations Excel Sheet")
st.markdown(
    "Download the input template: "
    "[Input_for_simulations.xlsx](https://drive.google.com/uc?export=download&id=1E3qED0b3qkmqtP9cF462BpfN_GKq0iO2)"
)

uploaded_excel = st.file_uploader("Upload your operations Excel file", type=["xlsx", "xls"])
excel_bytes    = None
excel_path_obj = None

if uploaded_excel is not None:
    excel_bytes    = uploaded_excel.read()
    excel_file_obj = pd.ExcelFile(io.BytesIO(excel_bytes))
    data_sheets    = excel_file_obj.sheet_names[1:]
    excel_path_obj = io.BytesIO(excel_bytes)
    st.success(f"✅  Loaded {len(data_sheets)} data sheets: {data_sheets}")


# ─────────────────────────────────────────────────────────────────────────────
#  COLUMN-DETECTION CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────
_TIME_COLS        = ["valid_time", "time", "date", "datetime"]
_U_COLS           = ["10m_u_component_of_wind", "u10"]
_V_COLS           = ["10m_v_component_of_wind", "v10"]
_WAVE_MARKERS     = {"swh","mwd","mwp","mean_wave_direction","mean_wave_period",
                     "significant_height_of_combined_wind_waves_and_swell"}
_WIND_MARKERS     = {"u10","v10","10m_u_component_of_wind","10m_v_component_of_wind"}
_DROP_COLS        = {"latitude","longitude","lat","lon"}
_WAVE_HEIGHT_COLS = ["significant_height_of_combined_wind_waves_and_swell","swh"]

def _find_col(df, candidates, label):
    col = next((c for c in candidates if c in df.columns), None)
    if col is None:
        raise ValueError(f"Could not find {label} column. Found: {df.columns.tolist()}")
    return col

def _norm_marker(x):
    s = str(x).strip().upper()
    return "" if s in {"","FALSE","NAN","NONE"} else s

def _norm_bool(x):
    return str(x).strip().upper() == "TRUE"

def _excel_read(src, sheet_name):
    if hasattr(src, "seek"):
        src.seek(0)
    return pd.read_excel(src, sheet_name=sheet_name)


# ─────────────────────────────────────────────────────────────────────────────
#  WEATHER HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def classify_csv(fp):
    df   = pd.read_csv(fp)
    cols = set(c.lower() for c in df.columns)
    is_wind = bool(cols & _WIND_MARKERS)
    is_wave = bool(cols & _WAVE_MARKERS)
    if   is_wind and not is_wave: label = "wind"
    elif is_wave and not is_wind: label = "wave"
    elif is_wind and is_wave:     label = "mixed"
    else:                         label = "unknown"
    return label, df

def _base_clean(df):
    df       = df.copy()
    time_col = next((c for c in _TIME_COLS if c in df.columns), None)
    if time_col is None:
        raise ValueError(f"No time column. Columns: {df.columns.tolist()}")
    df[time_col] = pd.to_datetime(df[time_col], errors="coerce")
    df = df.drop(columns=list(_DROP_COLS & set(df.columns)))
    return df.rename(columns={time_col: "time"}) if time_col != "time" else df

def prepare_wind_dataframe(df):
    df = _base_clean(df)
    u  = _find_col(df, _U_COLS, "U-wind")
    v  = _find_col(df, _V_COLS, "V-wind")
    df[u] = pd.to_numeric(df[u], errors="coerce")
    df[v] = pd.to_numeric(df[v], errors="coerce")
    df["Windspeed"]          = np.hypot(df[u], df[v])
    df["Adjusted_Windspeed"] = df["Windspeed"] * 0.9
    return df

def prepare_wave_dataframe(df):
    return _base_clean(df)

def load_weather_from_zip_bytes(zip_bytes, required_types):
    extract_dir = f"/tmp/era5_{uuid.uuid4().hex}"
    os.makedirs(extract_dir, exist_ok=True)
    with zipfile.ZipFile(io.BytesIO(zip_bytes), "r") as zf:
        zf.extractall(extract_dir)
    csv_files = glob.glob(os.path.join(extract_dir, "*.csv"))
    found = {}
    for fp in csv_files:
        label, df_raw = classify_csv(fp)
        if label in required_types and label not in found:
            found[label] = df_raw
        elif label == "mixed":
            found.setdefault("wind", df_raw)
            found.setdefault("wave", df_raw)
    shutil.rmtree(extract_dir, ignore_errors=True)
    return found

@st.cache_data(show_spinner="Loading weather data …")
def load_all_weather(project_zip_bytes, barge_zip_bytes):
    dfs_proj  = load_weather_from_zip_bytes(project_zip_bytes, ["wind","wave"])
    df_pw     = prepare_wind_dataframe(dfs_proj["wind"])
    df_waves  = prepare_wave_dataframe(dfs_proj["wave"])
    df_proj   = pd.merge(df_pw, df_waves, on="time", how="inner")
    dfs_barge = load_weather_from_zip_bytes(barge_zip_bytes, ["wind"])
    df_barge  = prepare_wind_dataframe(dfs_barge["wind"])
    return df_pw, df_waves, df_proj, df_barge

def add_wind_direction(df):
    df = df.copy()
    u_col = next((c for c in _U_COLS if c in df.columns), None)
    v_col = next((c for c in _V_COLS if c in df.columns), None)
    if u_col and v_col:
        df[u_col] = pd.to_numeric(df[u_col], errors="coerce")
        df[v_col] = pd.to_numeric(df[v_col], errors="coerce")
        df["Wind_Direction"] = (270 - np.degrees(np.arctan2(df[v_col], df[u_col]))) % 360
    return df


# ─────────────────────────────────────────────────────────────────────────────
#  SECTION 2 – WEATHER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("---")
st.header("🌦️  2.  Weather Factor Analysis")
st.subheader("2.1  Upload ERA5 Weather ZIPs")

col_pw, col_bw = st.columns(2)
with col_pw:
    project_zip_file = st.file_uploader("Project location ZIP (wind + waves)", type="zip", key="proj_zip")
with col_bw:
    barge_zip_file   = st.file_uploader("Barge location ZIP (wind only)",       type="zip", key="barge_zip")

df_project_weather = None
df_barge_wind      = None

_proj_bytes  = project_zip_file.read()  if project_zip_file  else \
               (open("era5_project_location.zip","rb").read() if os.path.exists("era5_project_location.zip") else None)
_barge_bytes = barge_zip_file.read()    if barge_zip_file    else \
               (open("era5_barge_location.zip","rb").read()   if os.path.exists("era5_barge_location.zip")   else None)

if _proj_bytes and _barge_bytes:
    try:
        _, _, df_project_weather, df_barge_wind = load_all_weather(_proj_bytes, _barge_bytes)
        st.success(f"✅  Weather loaded — {len(df_project_weather):,} project rows, {len(df_barge_wind):,} barge rows")
    except Exception as e:
        st.error(f"Weather loading error: {e}")

# ── 2.3 Weather plots ─────────────────────────────────────────────────────────
if df_project_weather is not None:
    if st.button("📊  Generate weather statistics & plots"):
        df_project_weather = add_wind_direction(df_project_weather)
        dfw = df_project_weather.copy()
        def season_of(m):
            return "Winter" if m in [12,1,2] else "Spring" if m in [3,4,5] else "Summer" if m in [6,7,8] else "Autumn"
        dfw["Season"] = dfw["time"].dt.month.apply(season_of)

        def calc_wind(s, lbl):
            return {"Period":lbl,"Avg Wind (m/s)":round(s.mean(),2),"P75":round(s.quantile(.75),2),"P90":round(s.quantile(.90),2),"Max":round(s.max(),2)}
        ws = [calc_wind(dfw["Adjusted_Windspeed"],"All Data")]
        for ssn in ["Winter","Spring","Summer","Autumn"]:
            ws.append(calc_wind(dfw.loc[dfw["Season"]==ssn,"Adjusted_Windspeed"], ssn))
        st.subheader("Wind Speed Summary"); st.dataframe(pd.DataFrame(ws).set_index("Period"))

        wc = next((c for c in _WAVE_HEIGHT_COLS if c in dfw.columns), None)
        if wc:
            def calc_wave(s, lbl):
                return {"Period":lbl,"Avg Hs (m)":round(s.mean(),2),"P75":round(s.quantile(.75),2),"P90":round(s.quantile(.90),2),"Max":round(s.max(),2)}
            wvs = [calc_wave(dfw[wc],"All Data")]
            for ssn in ["Winter","Spring","Summer","Autumn"]:
                wvs.append(calc_wave(dfw.loc[dfw["Season"]==ssn, wc], ssn))
            st.subheader("Wave Height Summary"); st.dataframe(pd.DataFrame(wvs).set_index("Period"))

        c1, c2 = st.columns(2)
        with c1:
            fig, ax = plt.subplots(figsize=(7,5))
            s = dfw["Adjusted_Windspeed"].dropna()
            ax.hist(s, bins=50, density=True, alpha=0.7, edgecolor="black")
            mu, sg = s.mean(), s.std()
            x = np.linspace(s.min(), s.max(), 300)
            if sg > 0: ax.plot(x,(1/(sg*np.sqrt(2*np.pi)))*np.exp(-0.5*((x-mu)/sg)**2),lw=2)
            ax.set(title="Wind Speed Distribution", xlabel="Adjusted Wind Speed (m/s)", ylabel="Density")
            ax.grid(True, alpha=0.3); st.pyplot(fig)
        with c2:
            if wc:
                fig, ax = plt.subplots(figsize=(7,5))
                s = pd.to_numeric(dfw[wc], errors="coerce").dropna()
                ax.hist(s, bins=50, density=True, alpha=0.7, edgecolor="black")
                mu, sg = s.mean(), s.std()
                x = np.linspace(s.min(), s.max(), 300)
                if sg > 0: ax.plot(x,(1/(sg*np.sqrt(2*np.pi)))*np.exp(-0.5*((x-mu)/sg)**2),lw=2)
                ax.set(title="Wave Height Distribution", xlabel="Significant Wave Height (m)", ylabel="Density")
                ax.grid(True, alpha=0.3); st.pyplot(fig)

        if wc:
            fig, ax = plt.subplots(figsize=(8,6))
            hb = ax.hexbin(dfw["Adjusted_Windspeed"], pd.to_numeric(dfw[wc], errors="coerce"), gridsize=60, cmap="plasma", mincnt=1)
            fig.colorbar(hb, ax=ax, label="Point Density")
            ax.set(title="Wind Speed vs Wave Height", xlabel="Adjusted Wind Speed (m/s)", ylabel="Significant Wave Height (m)")
            ax.grid(True, alpha=0.3); st.pyplot(fig)

            st.subheader("Wave Exceedance Table")
            thresholds = np.arange(0.5, float(max_wave)+0.5, 0.5)
            total_n = len(dfw[wc].dropna())
            exc_rows = [{"Threshold (m)":round(t,1), "Exceedance %":round((pd.to_numeric(dfw[wc],errors="coerce")>t).sum()/total_n*100,2)} for t in thresholds]
            st.dataframe(pd.DataFrame(exc_rows).set_index("Threshold (m)"))


# ─────────────────────────────────────────────────────────────────────────────
#  SIMULATION FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

def build_mp_timeline(excel_src, sheet_name, carry_cap, total_units, htd=1/24):
    raw = _excel_read(excel_src, sheet_name)
    raw.columns = [str(c).strip() for c in raw.columns]
    desc_col = next(c for c in raw.columns if str(c).strip().startswith("Description"))
    load_flags = raw["Loading"].map(_norm_marker)
    inst_flags = raw["Installation"].map(_norm_marker)
    load_end_idx = int(raw.index[load_flags.eq("END")][0])
    load_phase   = raw.iloc[:load_end_idx+1].reset_index(drop=True)
    inst_phase   = raw.iloc[load_end_idx+1:].reset_index(drop=True)
    load_pf = load_phase["Loading"].map(_norm_marker)
    inst_pf = inst_phase["Installation"].map(_norm_marker)
    inst_dec = int(inst_phase.index[inst_pf.eq("START")][0]) if len(inst_phase.index[inst_pf.eq("START")])>0 else 0

    def plists(pdf, mser, pname):
        return {"seq":pd.to_numeric(pdf["N"],errors="coerce").fillna(0).astype(int).tolist(),
                "desc":pdf[desc_col].astype(str).tolist(),
                "dur":pd.to_numeric(pdf["Seq Dur (hrs)"],errors="coerce").fillna(0).astype(float).tolist(),
                "mark":mser.tolist(),"phase":[pname]*len(pdf),"wxr":[pname=="installation"]*len(pdf)}

    ld=plists(load_phase,load_pf,"loading"); id_=plists(inst_phase,inst_pf,"installation")
    inventory=0; units_left=int(total_units); rows=[]

    def cl():
        if inventory>0: return math.ceil(units_left/carry_cap)+1
        return math.ceil(units_left/carry_cap) if units_left>0 else 0

    def add(sq,ds,dr,ph,wr):
        rows.append({"Sequence":int(sq),"Description":ds,"Phase":ph,"Weather_Restricted":bool(wr),
                     "Seq_Duration_Days":float(dr)*htd,"Inventory":int(inventory),
                     "WTG_Left":int(units_left),"Cycles_Left":int(cl())})

    while units_left>0 or inventory>0:
        while inventory<carry_cap and units_left>0:
            i=0
            while i<len(ld["seq"]):
                add(ld["seq"][i],ld["desc"][i],ld["dur"][i],ld["phase"][i],ld["wxr"][i])
                if ld["mark"][i]=="START": inventory+=1; units_left-=1
                if ld["mark"][i]=="END": break
                i+=1
        while inventory>0:
            for i in range(len(id_["seq"])):
                add(id_["seq"][i],id_["desc"][i],id_["dur"][i],id_["phase"][i],id_["wxr"][i])
                if i==inst_dec: inventory-=1
    out=pd.DataFrame(rows); out.insert(0,"N",range(1,len(out)+1))
    return out


def simulate_weather_impacts_mp(timeline_df, weather_df, excel_src, sheet_name,
                                 transit_distance, transit_speed, simulations=1,
                                 eligible_indices=None, seed=42, max_wait_hours=5000, htd=1/24):
    transit_hours=float(transit_distance)/float(transit_speed)
    raw=_excel_read(excel_src, sheet_name); raw.columns=[str(c).strip() for c in raw.columns]
    op=raw[["N","Seq Dur (hrs)","Tr (hrs)","Transit","Oplim Waves (m)","Oplim Wind (m/s)"]].copy()
    for c in ["N","Seq Dur (hrs)","Tr (hrs)"]: op[c]=pd.to_numeric(op[c],errors="coerce")
    op["Transit"]=op["Transit"].map(_norm_bool)
    op["Oplim Waves (m)"]=pd.to_numeric(op["Oplim Waves (m)"],errors="coerce").fillna(np.inf)
    op["Oplim Wind (m/s)"]=pd.to_numeric(op["Oplim Wind (m/s)"],errors="coerce").fillna(np.inf)
    nt=~op["Transit"]
    op.loc[nt,"Seq Dur (hrs)"]=op.loc[nt,"Seq Dur (hrs)"].fillna(op.loc[nt,"Tr (hrs)"])
    op.loc[nt,"Tr (hrs)"]=op.loc[nt,"Tr (hrs)"].fillna(op.loc[nt,"Seq Dur (hrs)"])
    op.loc[op["Transit"],["Seq Dur (hrs)","Tr (hrs)"]]=transit_hours
    op=op.dropna(subset=["N","Seq Dur (hrs)","Tr (hrs)"]).drop_duplicates(subset=["N"]); op["N"]=op["N"].astype(int)
    lm=dict(zip(op["N"],np.minimum(op["Seq Dur (hrs)"],op["Tr (hrs)"])))
    hm=dict(zip(op["N"],np.maximum(op["Seq Dur (hrs)"],op["Tr (hrs)"])))
    wvm=dict(zip(op["N"],op["Oplim Waves (m)"])); wdm=dict(zip(op["N"],op["Oplim Wind (m/s)"]))
    tm=dict(zip(op["N"],op["Transit"]))
    base=timeline_df[["N","Sequence","Inventory","WTG_Left","Cycles_Left","Weather_Restricted","Phase"]].copy()
    base["Sequence"]=base["Sequence"].astype(int); base["Weather_Restricted"]=base["Weather_Restricted"].astype(bool)
    wave_col=_find_col(weather_df,_WAVE_HEIGHT_COLS,"wave height")
    w=weather_df[["time",wave_col,"Adjusted_Windspeed"]].copy()
    w["time"]=pd.to_datetime(w["time"],errors="coerce")
    for c in [wave_col,"Adjusted_Windspeed"]: w[c]=pd.to_numeric(w[c],errors="coerce")
    w=w.dropna().sort_values("time").drop_duplicates(subset="time").reset_index(drop=True)
    wt=w["time"].to_numpy(dtype="datetime64[ns]"); ww=w[wave_col].to_numpy(dtype=float); wnd=w["Adjusted_Windspeed"].to_numpy(dtype=float)
    eli=eligible_indices if eligible_indices is not None else np.arange(len(wt))
    rng=np.random.default_rng(seed); n=len(base)
    tot_m=np.zeros((n,simulations),dtype=float); dwn_m=np.zeros((n,simulations),dtype=float)
    seqs=base["Sequence"].to_numpy(dtype=int)
    sl=np.array([lm[s] for s in seqs],dtype=float); sh=np.array([hm[s] for s in seqs],dtype=float)
    swl=np.array([wvm[s] for s in seqs],dtype=float); swd=np.array([wdm[s] for s in seqs],dtype=float)
    sit=np.array([tm[s] for s in seqs],dtype=bool); swx=base["Weather_Restricted"].to_numpy(dtype=bool)

    def wait(start,wlim,wndlim):
        idx=int(np.searchsorted(wt,np.datetime64(pd.Timestamp(start)),side="left"))
        for h in range(max_wait_hours+1):
            i=idx+h
            if i>=len(wt): raise RuntimeError("Ran out of weather data.")
            if ww[i]<=wlim and wnd[i]<=wndlim: return h,i
        raise RuntimeError("Exceeded max_wait_hours.")

    summ=[]
    for si in range(simulations):
        ct=pd.Timestamp(wt[int(rng.choice(eli))]); ss=ct; ta=td=0.0
        for ri in range(n):
            if swx[ri]: dh,pi=wait(ct,swl[ri],swd[ri]); op_s=pd.Timestamp(wt[pi])
            else: dh,op_s=0.0,ct
            ah=transit_hours if sit[ri] else float(rng.uniform(sl[ri],sh[ri]))
            tot_m[ri,si]=(dh+ah)*htd; dwn_m[ri,si]=dh*htd; ta+=ah; td+=dh
            ct=op_s+pd.Timedelta(hours=ah)
        summ.append({"Simulation":f"Sim{si+1}","Start_Date":ss,"Finish_Date":ct,
                     "Total_Active_Days":ta*htd,"Total_Downtime_Days":td*htd,
                     "Total_Project_Days":(ta+td)*htd})
    sc=[f"Sim{i}" for i in range(1,simulations+1)]
    return (pd.concat([base,pd.DataFrame(tot_m,columns=sc,index=base.index)],axis=1),
            pd.concat([base,pd.DataFrame(dwn_m,columns=sc,index=base.index)],axis=1),
            pd.DataFrame(summ))


def get_seasonal_eligible_indices(weather_times, start_date_str, window_days=7, latest_start="2022-12-31 23:00:00"):
    target=pd.Timestamp(start_date_str); times_pd=pd.DatetimeIndex(weather_times); cutoff=pd.Timestamp(latest_start)
    doys=times_pd.day_of_year; tdoy=target.day_of_year
    circ=np.minimum(np.abs(doys-tdoy),365-np.abs(doys-tdoy))
    eligible=np.where((circ<=window_days)&(times_pd<=cutoff))[0]
    if len(eligible)==0: raise ValueError(f"No eligible starts near {target.strftime('%d %B')}.")
    return eligible


def build_w90_timeline(excel_src, sheet_name, carry_cap, total_wtg, htd=1/24,
                        dock_distance=None, transit_speed=None):
    raw=_excel_read(excel_src, sheet_name); raw.columns=[str(c).strip() for c in raw.columns]
    lf=raw["Loading"].map(_norm_marker); if_=raw["Installation"].map(_norm_marker)
    lei=int(lf[lf.eq("END")].index[0]); isi=int(if_[if_.eq("START")].index[0]); iei=int(if_[if_.eq("END")].index[0])
    lp=raw.iloc[:lei+1].reset_index(drop=True); ip=raw.iloc[isi:iei+1].reset_index(drop=True)
    lpf=lp["Loading"].map(_norm_marker); ipf=ip["Installation"].map(_norm_marker)
    isp=int(ipf[ipf.eq("START")].index[0])

    def plists(pdf, mser):
        ts=pdf["Transit"].map(_norm_bool)
        durs=[dock_distance/transit_speed if ts.iloc[i] else float(pd.to_numeric(pdf.iloc[i]["Seq Dur (hrs)"],errors="coerce") or 0.0) for i in range(len(pdf))]
        return {"seq":pd.to_numeric(pdf["N"],errors="coerce").fillna(0).astype(int).tolist(),
                "desc":pdf["Description"].astype(str).tolist(),"dur":durs,"mark":mser.tolist(),
                "p1":pdf["Parallell  1"].map(_norm_bool).tolist(),"p2":pdf["Parallell  2"].map(_norm_bool).tolist()}

    ld=plists(lp,lpf); id_=plists(ip,ipf)
    inventory=0; wtg_left=int(total_wtg); rows=[]

    def cl():
        if inventory>0: return math.ceil(wtg_left/carry_cap)+1
        return math.ceil(wtg_left/carry_cap) if wtg_left>0 else 0

    def add(sq,ds,dr,p1,p2):
        rows.append({"Sequence":int(sq),"Description":ds,"Seq_Duration_Hours":float(dr)*htd,
                     "Parallel_1":bool(p1),"Parallel_2":bool(p2),"Inventory":int(inventory),
                     "WTG_Left":int(wtg_left),"Cycles_Left":int(cl())})

    while wtg_left>0 or inventory>0:
        if inventory==0 and wtg_left>0:
            for i in range(len(ld["seq"])): add(ld["seq"][i],ld["desc"][i],ld["dur"][i],ld["p1"][i],ld["p2"][i])
            loaded=min(carry_cap,wtg_left); inventory=loaded; wtg_left-=loaded
        while inventory>0:
            i=0
            while i<len(id_["seq"]):
                add(id_["seq"][i],id_["desc"][i],id_["dur"][i],id_["p1"][i],id_["p2"][i])
                if id_["mark"][i]=="START": inventory-=1
                if id_["mark"][i]=="END":
                    if inventory>0: i=isp; continue
                    else: break
                i+=1

    df=pd.DataFrame(rows); collapsed=[]; i=0
    while i<len(df):
        is_par=bool(df.at[i,"Parallel_1"] or df.at[i,"Parallel_2"])
        if not is_par:
            collapsed.append({"Sequence":int(df.at[i,"Sequence"]),"Description":df.at[i,"Description"],
                               "Seq_Duration_Hours":float(df.at[i,"Seq_Duration_Hours"]),"Inventory":int(df.at[i,"Inventory"]),
                               "WTG_Left":int(df.at[i,"WTG_Left"]),"Cycles_Left":int(df.at[i,"Cycles_Left"]),
                               "Parallel_Block":False,"Parallel_Branch_1_Hours":0.0,"Parallel_Branch_2_Hours":0.0,
                               "Source_Sequences":str(int(df.at[i,"Sequence"]))}); i+=1
        else:
            block=[]; p1s=p2s=0.0
            while i<len(df) and bool(df.at[i,"Parallel_1"] or df.at[i,"Parallel_2"]):
                r=df.iloc[i]; block.append(r); d=float(r["Seq_Duration_Hours"])
                if r["Parallel_1"]: p1s+=d
                if r["Parallel_2"]: p2s+=d
                i+=1
            last=block[-1]
            collapsed.append({"Sequence":int(block[0]["Sequence"]),"Description":"Parallel block",
                               "Seq_Duration_Hours":max(p1s,p2s),"Inventory":int(last["Inventory"]),
                               "WTG_Left":int(last["WTG_Left"]),"Cycles_Left":int(last["Cycles_Left"]),
                               "Parallel_Block":True,"Parallel_Branch_1_Hours":p1s,"Parallel_Branch_2_Hours":p2s,
                               "Source_Sequences":", ".join(str(int(r["Sequence"])) for r in block)})
    out=pd.DataFrame(collapsed); out.insert(0,"N",range(1,len(out)+1))
    return out


def add_duration_simulations_w90(timeline_df, excel_src, sheet_name, simulations=1, seed=42,
                                   htd=1/24, barge_flag=False, dock_distance=None,
                                   barge_distance=None, transit_speed=None):
    raw=_excel_read(excel_src, sheet_name); raw.columns=[str(c).strip() for c in raw.columns]
    dur=raw[["N","Seq Dur (hrs)","Tr (hrs)","Transit","Parallell  1","Parallell  2"]].copy()
    dur["N"]=pd.to_numeric(dur["N"],errors="coerce")
    dur["Seq Dur (hrs)"]=pd.to_numeric(dur["Seq Dur (hrs)"],errors="coerce")
    dur["Tr (hrs)"]=pd.to_numeric(dur["Tr (hrs)"],errors="coerce")
    dur["Transit"]=dur["Transit"].map(_norm_bool); dur["Parallell  1"]=dur["Parallell  1"].map(_norm_bool); dur["Parallell  2"]=dur["Parallell  2"].map(_norm_bool)
    dur=dur.dropna(subset=["N"]).drop_duplicates(subset=["N"]); dur["N"]=dur["N"].astype(int)
    tm=dict(zip(dur["N"],dur["Transit"])); p1m=dict(zip(dur["N"],dur["Parallell  1"])); p2m=dict(zip(dur["N"],dur["Parallell  2"]))
    bl=dict(zip(dur["N"],dur[["Seq Dur (hrs)","Tr (hrs)"]].min(axis=1))); bh=dict(zip(dur["N"],dur[["Seq Dur (hrs)","Tr (hrs)"]].max(axis=1)))
    df=timeline_df.copy(); df=df.drop(columns=[c for c in df.columns if str(c).startswith("Sim")])
    rspecs=[]
    for _,row in df.iterrows():
        is_par=bool(row["Parallel_Block"])
        seqs=([int(s.strip()) for s in str(row["Source_Sequences"]).split(",") if s.strip()] if is_par else [int(row["Sequence"])])
        rspecs.append((is_par,seqs))

    def _sim(dist, soff=0):
        lm={s:(dist/transit_speed if tm.get(s) else float(bl[s])) for s in bl}
        hm_={s:(dist/transit_speed if tm.get(s) else float(bh[s])) for s in bh}
        rng=np.random.default_rng(None if seed is None else seed+soff)
        mat=np.empty((len(df),simulations),dtype=float)
        for ri,(is_par,seqs) in enumerate(rspecs):
            if not is_par:
                mat[ri,:]=rng.uniform(lm[seqs[0]],hm_[seqs[0]],size=simulations)*htd
            else:
                b1=np.zeros(simulations); b2=np.zeros(simulations)
                for s in seqs:
                    draws=rng.uniform(lm[s],hm_[s],size=simulations)
                    if p1m.get(s): b1+=draws
                    if p2m.get(s): b2+=draws
                mat[ri,:]=np.maximum(b1,b2)*htd
        return pd.concat([df.copy(),pd.DataFrame(mat,columns=[f"Sim{i}" for i in range(1,simulations+1)],index=df.index)],axis=1)

    dock_df=_sim(dock_distance,0)
    return (dock_df,_sim(barge_distance,1)) if barge_flag else dock_df


def build_sequential_wtg_timeline(excel_src, sheet_name, carry_cap, total_wtg, htd=1/24,
                                   dock_distance=None, transit_speed=None):
    raw=_excel_read(excel_src, sheet_name); raw.columns=[str(c).strip() for c in raw.columns]
    lf=raw["Loading"].map(_norm_marker); if_=raw["Installation"].map(_norm_marker)
    isi=int(if_[if_.eq("START")].index[0])
    iec=if_[if_.eq("END")].index; iei=int(iec[0]) if len(iec)>0 else int(raw.index[-1])
    lp=raw.iloc[:isi].reset_index(drop=True); ip=raw.iloc[isi:iei+1].reset_index(drop=True)
    lpf=lp["Loading"].map(_norm_marker); ipf=ip["Installation"].map(_norm_marker)

    def plists(pdf, mser):
        ts=pdf["Transit"].map(_norm_bool)
        durs=[dock_distance/transit_speed if ts.iloc[i] else float(pd.to_numeric(pdf.iloc[i]["Seq Dur (hrs)"],errors="coerce") or 0.0) for i in range(len(pdf))]
        return {"seq":pd.to_numeric(pdf["N"],errors="coerce").fillna(0).astype(int).tolist(),
                "desc":pdf["Description"].astype(str).tolist(),"dur":durs,"mark":mser.tolist()}

    ld=plists(lp,lpf); id_=plists(ip,ipf)
    inventory=0; wtg_left=int(total_wtg); rows=[]

    def cl():
        if inventory>0: return math.ceil(wtg_left/carry_cap)+1
        return math.ceil(wtg_left/carry_cap) if wtg_left>0 else 0

    def add(sq,ds,dr):
        rows.append({"Sequence":int(sq),"Description":ds,"Seq_Duration_Hours":float(dr)*htd,
                     "Inventory":int(inventory),"WTG_Left":int(wtg_left),"Cycles_Left":int(cl())})

    while wtg_left>0 or inventory>0:
        while inventory<carry_cap and wtg_left>0:
            i=0
            while i<len(ld["seq"]):
                add(ld["seq"][i],ld["desc"][i],ld["dur"][i])
                if ld["mark"][i]=="START": inventory+=1; wtg_left-=1
                if ld["mark"][i]=="END": break
                i+=1
        while inventory>0:
            i=0
            while i<len(id_["seq"]):
                add(id_["seq"][i],id_["desc"][i],id_["dur"][i])
                if id_["mark"][i]=="START": inventory-=1
                if id_["mark"][i]=="END": break
                i+=1
    out=pd.DataFrame(rows); out.insert(0,"N",range(1,len(out)+1))
    return out


def add_duration_simulations_sequential(timeline_df, excel_src, sheet_name, simulations=1,
                                         dock_distance=0, transit_speed=1, seed=42, htd=1/24):
    raw=_excel_read(excel_src, sheet_name); raw.columns=[str(c).strip() for c in raw.columns]
    dur=raw[["N","Seq Dur (hrs)","Tr (hrs)","Transit"]].copy()
    dur["N"]=pd.to_numeric(dur["N"],errors="coerce")
    dur["Seq Dur (hrs)"]=pd.to_numeric(dur["Seq Dur (hrs)"],errors="coerce")
    dur["Tr (hrs)"]=pd.to_numeric(dur["Tr (hrs)"],errors="coerce")
    dur["Transit"]=dur["Transit"].map(_norm_bool)
    dur=dur.dropna(subset=["N"]).drop_duplicates(subset=["N"]); dur["N"]=dur["N"].astype(int)
    td_days=(dock_distance/transit_speed)*htd
    lm={}; hm={}; fm={}
    for _,row in dur.iterrows():
        s=int(row["N"])
        if row["Transit"]: fm[s]=float(td_days)
        else: lm[s]=float(min(row["Seq Dur (hrs)"],row["Tr (hrs)"]))*htd; hm[s]=float(max(row["Seq Dur (hrs)"],row["Tr (hrs)"]))*htd
    df=timeline_df.copy(); df=df.drop(columns=[c for c in df.columns if str(c).startswith("Sim")])
    df["_l"]=df["Sequence"].map(lm); df["_h"]=df["Sequence"].map(hm); df["_f"]=df["Sequence"].map(fm)
    lows=df["_l"].to_numpy(dtype=float); highs=df["_h"].to_numpy(dtype=float); fixeds=df["_f"].to_numpy(dtype=float)
    tmask=~np.isnan(fixeds); ntmask=~tmask
    rng=np.random.default_rng(seed); mat=np.empty((len(df),simulations),dtype=float)
    mat[tmask,:]=fixeds[tmask,None]
    mat[ntmask,:]=rng.uniform(lows[ntmask,None],highs[ntmask,None],size=(ntmask.sum(),simulations))
    return pd.concat([df.drop(columns=["_l","_h","_f"]),
                      pd.DataFrame(mat,columns=[f"Sim{i}" for i in range(1,simulations+1)],index=df.index)],axis=1)


def simulate_weather_impacts_w90(timeline_df, project_weather_df, loading_weather_df,
                                   excel_src, sheet_name, transit_distance, transit_speed,
                                   simulations=1, eligible_indices=None, seed=42,
                                   max_wait_hours=5000, htd=1/24, mp_offset_days=None):
    transit_hours=float(transit_distance)/float(transit_speed)
    raw=_excel_read(excel_src, sheet_name); raw.columns=[str(c).strip() for c in raw.columns]
    op=raw[["N","Seq Dur (hrs)","Tr (hrs)","Transit","Loading","Installation",
            "Oplim Waves (m)","Oplim Wind (m/s)","Parallell  1","Parallell  2"]].copy()
    op["N"]=pd.to_numeric(op["N"],errors="coerce")
    for c in ["Seq Dur (hrs)","Tr (hrs)"]: op[c]=pd.to_numeric(op[c],errors="coerce")
    op["Transit"]=op["Transit"].map(_norm_bool); op["Loading"]=op["Loading"].map(_norm_marker); op["Installation"]=op["Installation"].map(_norm_marker)
    op["Oplim Waves (m)"]=pd.to_numeric(op["Oplim Waves (m)"],errors="coerce").fillna(np.inf)
    op["Oplim Wind (m/s)"]=pd.to_numeric(op["Oplim Wind (m/s)"],errors="coerce").fillna(np.inf)
    op["Parallell  1"]=op["Parallell  1"].map(_norm_bool); op["Parallell  2"]=op["Parallell  2"].map(_norm_bool)
    nt=~op["Transit"]
    op.loc[nt,"Seq Dur (hrs)"]=op.loc[nt,"Seq Dur (hrs)"].fillna(op.loc[nt,"Tr (hrs)"])
    op.loc[nt,"Tr (hrs)"]=op.loc[nt,"Tr (hrs)"].fillna(op.loc[nt,"Seq Dur (hrs)"])
    op.loc[op["Transit"],["Seq Dur (hrs)","Tr (hrs)"]]=transit_hours
    op=op.dropna(subset=["N","Seq Dur (hrs)","Tr (hrs)"]).drop_duplicates(subset=["N"]); op["N"]=op["N"].astype(int)
    le=op.index[op["Loading"].eq("END")]
    loading_seq_set=set(op.loc[:int(le[0]),"N"].tolist()) if len(le)>0 else set()
    lm=dict(zip(op["N"],np.minimum(op["Seq Dur (hrs)"],op["Tr (hrs)"])))
    hm=dict(zip(op["N"],np.maximum(op["Seq Dur (hrs)"],op["Tr (hrs)"])))
    wvm=dict(zip(op["N"],op["Oplim Waves (m)"])); wdm=dict(zip(op["N"],op["Oplim Wind (m/s)"]))
    p1m=dict(zip(op["N"],op["Parallell  1"])); p2m=dict(zip(op["N"],op["Parallell  2"]))
    tm=dict(zip(op["N"],op["Transit"]))
    base=timeline_df[["N","Sequence","Inventory","WTG_Left","Cycles_Left","Parallel_Block","Source_Sequences"]].copy()
    base["Sequence"]=pd.to_numeric(base["Sequence"],errors="coerce").astype(int)
    rspecs=[]
    for _,row in base.iterrows():
        is_par=bool(row["Parallel_Block"])
        seqs=([int(s.strip()) for s in str(row["Source_Sequences"]).split(",") if s.strip()] if is_par else [int(row["Sequence"])])
        pt="loading" if all(s in loading_seq_set for s in seqs) else "offshore"
        rspecs.append({"is_parallel":is_par,"seqs":seqs,
                       "wave_lim":float(min(wvm.get(s,np.inf) for s in seqs)),
                       "wind_lim":float(min(wdm.get(s,np.inf) for s in seqs)),"phase_type":pt})
    wave_col=_find_col(project_weather_df,_WAVE_HEIGHT_COLS,"wave height")
    def prep(df, cols):
        df=df[list(cols)].copy(); df["time"]=pd.to_datetime(df["time"],errors="coerce")
        for c in cols:
            if c!="time": df[c]=pd.to_numeric(df[c],errors="coerce")
        return df.dropna().sort_values("time").drop_duplicates(subset="time").reset_index(drop=True)
    wp=prep(project_weather_df,{"time",wave_col,"Adjusted_Windspeed"})
    wl=prep(loading_weather_df,{"time","Adjusted_Windspeed"})
    pt=wp["time"].to_numpy(dtype="datetime64[ns]"); pw=wp[wave_col].to_numpy(dtype=float); pnd=wp["Adjusted_Windspeed"].to_numpy(dtype=float)
    lt=wl["time"].to_numpy(dtype="datetime64[ns]"); lnd=wl["Adjusted_Windspeed"].to_numpy(dtype=float)
    eli=eligible_indices if eligible_indices is not None else np.arange(len(pt))
    rng=np.random.default_rng(seed); n=len(base)
    tot_m=np.zeros((n,simulations),dtype=float); dwn_m=np.zeros((n,simulations),dtype=float)

    def _wait(times,winds,waves,start,wlim,wndlim,check_w):
        idx=int(np.searchsorted(times,np.datetime64(pd.Timestamp(start)),side="left"))
        for h in range(max_wait_hours+1):
            i=idx+h
            if i>=len(times): raise RuntimeError("Ran out of weather data.")
            if winds[i]<=wndlim and (not check_w or waves[i]<=wlim): return h,i
        raise RuntimeError("Exceeded max_wait_hours.")

    def _draw(spec):
        if not spec["is_parallel"]:
            s=spec["seqs"][0]; return transit_hours if tm.get(s) else float(rng.uniform(lm[s],hm[s]))
        b1=b2=0.0
        for s in spec["seqs"]:
            d=transit_hours if tm.get(s) else float(rng.uniform(lm[s],hm[s]))
            if p1m.get(s): b1+=d
            if p2m.get(s): b2+=d
        return max(b1,b2)

    summ=[]
    for si in range(simulations):
        sidx=int(rng.choice(eli)); bs=pd.Timestamp(pt[sidx])
        offset=float(mp_offset_days[si]) if mp_offset_days is not None else 0.0
        ct=bs+pd.Timedelta(days=offset); ss=ct; ta=td=0.0
        for ri,spec in enumerate(rspecs):
            if spec["phase_type"]=="loading": dh,pi=_wait(lt,lnd,None,ct,np.inf,spec["wind_lim"],False); op_s=pd.Timestamp(lt[pi])
            else: dh,pi=_wait(pt,pnd,pw,ct,spec["wave_lim"],spec["wind_lim"],True); op_s=pd.Timestamp(pt[pi])
            ah=_draw(spec)
            tot_m[ri,si]=(dh+ah)*htd; dwn_m[ri,si]=dh*htd; ta+=ah; td+=dh
            ct=op_s+pd.Timedelta(hours=ah)
        summ.append({"Simulation":f"Sim{si+1}","WTG_Start_Date":ss,"Finish_Date":ct,"MP_Offset_Days":offset,
                     "Total_Active_Days":ta*htd,"Total_Downtime_Days":td*htd,"Total_Project_Days":(ta+td)*htd})
    sc=[f"Sim{i}" for i in range(1,simulations+1)]
    df_out=base[["N","Sequence","Inventory","WTG_Left","Cycles_Left"]].copy()
    return (pd.concat([df_out,pd.DataFrame(tot_m,columns=sc,index=df_out.index)],axis=1),
            pd.concat([df_out,pd.DataFrame(dwn_m,columns=sc,index=df_out.index)],axis=1),
            pd.DataFrame(summ))


def simulate_sequential_wtg_weather(timeline_df, weather_df, excel_src, sheet_name,
                                     transit_distance, transit_speed, simulations=1,
                                     eligible_indices=None, seed=42, max_wait_hours=5000,
                                     htd=1/24, mp_offset_days=None):
    transit_hours=float(transit_distance)/float(transit_speed)
    raw=_excel_read(excel_src, sheet_name); raw.columns=[str(c).strip() for c in raw.columns]
    op=raw[["N","Seq Dur (hrs)","Tr (hrs)","Transit","Oplim Waves (m)","Oplim Wind (m/s)"]].copy()
    op["N"]=pd.to_numeric(op["N"],errors="coerce")
    for c in ["Seq Dur (hrs)","Tr (hrs)"]: op[c]=pd.to_numeric(op[c],errors="coerce")
    op["Transit"]=op["Transit"].map(_norm_bool)
    op["Oplim Waves (m)"]=pd.to_numeric(op["Oplim Waves (m)"],errors="coerce").fillna(np.inf)
    op["Oplim Wind (m/s)"]=pd.to_numeric(op["Oplim Wind (m/s)"],errors="coerce").fillna(np.inf)
    nt=~op["Transit"]
    op.loc[nt,"Seq Dur (hrs)"]=op.loc[nt,"Seq Dur (hrs)"].fillna(op.loc[nt,"Tr (hrs)"])
    op.loc[nt,"Tr (hrs)"]=op.loc[nt,"Tr (hrs)"].fillna(op.loc[nt,"Seq Dur (hrs)"])
    op.loc[op["Transit"],["Seq Dur (hrs)","Tr (hrs)"]]=transit_hours
    op=op.dropna(subset=["N","Seq Dur (hrs)","Tr (hrs)"]).drop_duplicates(subset=["N"]); op["N"]=op["N"].astype(int)
    lm=dict(zip(op["N"],np.minimum(op["Seq Dur (hrs)"],op["Tr (hrs)"])))
    hm=dict(zip(op["N"],np.maximum(op["Seq Dur (hrs)"],op["Tr (hrs)"])))
    wvm=dict(zip(op["N"],op["Oplim Waves (m)"])); wdm=dict(zip(op["N"],op["Oplim Wind (m/s)"]))
    tm=dict(zip(op["N"],op["Transit"]))
    base=timeline_df[["N","Sequence","Inventory","WTG_Left","Cycles_Left"]].copy()
    base["Sequence"]=pd.to_numeric(base["Sequence"],errors="coerce").astype(int)
    wave_col=_find_col(weather_df,_WAVE_HEIGHT_COLS,"wave height")
    w=weather_df[["time",wave_col,"Adjusted_Windspeed"]].copy()
    w["time"]=pd.to_datetime(w["time"],errors="coerce")
    for c in [wave_col,"Adjusted_Windspeed"]: w[c]=pd.to_numeric(w[c],errors="coerce")
    w=w.dropna().sort_values("time").drop_duplicates(subset="time").reset_index(drop=True)
    wt=w["time"].to_numpy(dtype="datetime64[ns]"); ww_=w[wave_col].to_numpy(dtype=float); wnd=w["Adjusted_Windspeed"].to_numpy(dtype=float)
    eli=eligible_indices if eligible_indices is not None else np.arange(len(wt))
    rng=np.random.default_rng(seed); n=len(base)
    tot_m=np.zeros((n,simulations),dtype=float); dwn_m=np.zeros((n,simulations),dtype=float)
    seqs=base["Sequence"].to_numpy(dtype=int)
    sl=np.array([lm[s] for s in seqs],dtype=float); sh=np.array([hm[s] for s in seqs],dtype=float)
    swl=np.array([wvm[s] for s in seqs],dtype=float); swd=np.array([wdm[s] for s in seqs],dtype=float)
    sit=np.array([tm[s] for s in seqs],dtype=bool)

    def _wait(start,wlim,wndlim):
        idx=int(np.searchsorted(wt,np.datetime64(pd.Timestamp(start)),side="left"))
        for h in range(max_wait_hours+1):
            i=idx+h
            if i>=len(wt): raise RuntimeError("Ran out of weather data.")
            if ww_[i]<=wlim and wnd[i]<=wndlim: return h,i
        raise RuntimeError("Exceeded max_wait_hours.")

    summ=[]
    for si in range(simulations):
        sidx=int(rng.choice(eli)); bs=pd.Timestamp(wt[sidx])
        offset=float(mp_offset_days[si]) if mp_offset_days is not None else 0.0
        ct=bs+pd.Timedelta(days=offset); ss=ct; ta=td=0.0
        for ri in range(n):
            dh,pi=_wait(ct,swl[ri],swd[ri]); op_s=pd.Timestamp(wt[pi])
            ah=transit_hours if sit[ri] else float(rng.uniform(sl[ri],sh[ri]))
            tot_m[ri,si]=(dh+ah)*htd; dwn_m[ri,si]=dh*htd; ta+=ah; td+=dh
            ct=op_s+pd.Timedelta(hours=ah)
        summ.append({"Simulation":f"Sim{si+1}","WTG_Start_Date":ss,"Finish_Date":ct,"MP_Offset_Days":offset,
                     "Total_Active_Days":ta*htd,"Total_Downtime_Days":td*htd,"Total_Project_Days":(ta+td)*htd})
    sc=[f"Sim{i}" for i in range(1,simulations+1)]
    return (pd.concat([base,pd.DataFrame(tot_m,columns=sc,index=base.index)],axis=1),
            pd.concat([base,pd.DataFrame(dwn_m,columns=sc,index=base.index)],axis=1),
            pd.DataFrame(summ))


# ── Output / plotting helpers ─────────────────────────────────────────────────
def _sim_totals(df):
    sc=[c for c in df.columns if str(c).startswith("Sim")]
    return df[sc].apply(pd.to_numeric,errors="coerce").sum(axis=0).astype(float)

def build_project_summary(no_wx_df, wx_total_df, wx_down_df, label):
    nwt=_sim_totals(no_wx_df); wxt=_sim_totals(wx_total_df); wxd=_sim_totals(wx_down_df); net=wxt-wxd
    p0_perfect=float(nwt.min())
    def row(series, category, p0_override=None):
        v=pd.to_numeric(series,errors="coerce").dropna().to_numpy(dtype=float)
        r={"Project":label,"Category":category,"Average":round(float(np.mean(v)),2),
           "P0":round(p0_override if p0_override is not None else float(np.percentile(v,0)),2)}
        for p in [10,25,50,75,90,100]: r[f"P{p}"]=round(float(np.percentile(v,p)),2)
        return r
    return pd.DataFrame([row(wxt,"Operation Days",p0_perfect),row(wxd,"Downtime Days",0.0),row(net,"Net Operation",p0_perfect)])

def plot_cumulative(results_dict, title):
    _STYLE={"W90":{"c":"red","ls":"-"},"W90_dock":{"c":"red","ls":"-"},"W90_barge":{"c":"red","ls":":"},
            "barge":{"c":"red","ls":":"},"JUV":{"c":"blue","ls":"-"},"FIV":{"c":"orange","ls":"-"}}
    fig,ax=plt.subplots(figsize=(12,7)); p50_finals=[]
    for name,df in results_dict.items():
        sc=[c for c in df.columns if str(c).startswith("Sim")]
        cum=df[sc].cumsum(axis=0); p50=cum.median(axis=1)
        style=_STYLE.get(name,{"c":"grey","ls":"-"})
        ax.plot(np.arange(1,len(df)+1),cum,color="grey",alpha=0.15,linewidth=1)
        ax.plot(np.arange(1,len(df)+1),p50,label=f"{name} P50",color=style["c"],linestyle=style["ls"],linewidth=2.5)
        p50_finals.append(float(p50.iloc[-1]))
    if p50_finals and min(p50_finals)>0:
        ref=min(p50_finals)
        ax.secondary_yaxis("right",functions=(lambda y:y/ref*100,lambda y:y/100*ref)).set_ylabel("Relative Completion (%)")
    ax.set_title(title); ax.set_xlabel("Sequence Step"); ax.set_ylabel("Cumulative Duration (days)")
    ax.grid(True,alpha=0.3); ax.legend(); plt.tight_layout()
    return fig

def plot_bar_comparison(bar_rows):
    df=pd.DataFrame(bar_rows)
    fig,ax1=plt.subplots(figsize=(11,5.5))
    fig.patch.set_facecolor("#f2f2f2"); ax1.set_facecolor("#f2f2f2")
    x=np.arange(len(df)); w=0.34
    b1=ax1.bar(x-w/2,df["Net_P50"],   w,label="Net duration (P50)",   color="#1f3763")
    b2=ax1.bar(x+w/2,df["Total_P50"], w,label="Total duration (P50)", color="#6a9f3f")
    ax1.set_xticks(x); ax1.set_xticklabels(df["Vessel"]); ax1.set_ylabel("Duration (days)")
    ax1.set_title("Campaign Duration Comparison",fontweight="bold"); ax1.grid(axis="y",alpha=0.3)
    off=float(df["Total_P50"].max())*0.02
    for bars in (b1,b2):
        for b in bars: ax1.text(b.get_x()+b.get_width()/2,b.get_height()+off,f"{b.get_height():.1f}",ha="center",va="bottom",fontsize=9)
    ax2=ax1.twinx()
    ax2.plot(x,df["Downtime_Pct"],color="red",linewidth=2.2,label="Weather downtime (%)")
    ax2.set_ylabel("Weather downtime (%)")
    for xi,yi in zip(x,df["Downtime_Pct"]):
        ax2.text(xi,yi,f"{yi:.0f}%",ha="center",va="bottom",fontsize=9,
                 bbox=dict(boxstyle="square,pad=0.2",facecolor="white",edgecolor="none",alpha=0.8))
    h1,l1=ax1.get_legend_handles_labels(); h2,l2=ax2.get_legend_handles_labels()
    ax1.legend(h1+h2,l1+l2,loc="upper center",bbox_to_anchor=(0.5,-0.15),ncol=3,frameon=False)
    for a in (ax1,ax2):
        for sp in a.spines.values(): sp.set_visible(False)
    plt.tight_layout(); return fig


# ─────────────────────────────────────────────────────────────────────────────
#  SECTION 3-5 – RUN ALL SIMULATIONS
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("---")
st.header("🏗️  3–5.  Run All Simulations (MP + WTG)")

if not (uploaded_excel and df_project_weather is not None):
    st.info("Upload the operations Excel file and weather ZIPs above to enable simulations.")
else:
    if st.button("▶️  Run All Simulations", type="primary"):

        _MP_CONFIGS = {}
        if W90_sim: _MP_CONFIGS["W90"]   = (W90_MP_sheet,  W90_carry_cap)
        if barge:   _MP_CONFIGS["barge"] = (W90_MP_sheet,  W90_carry_cap)
        if JUV_sim: _MP_CONFIGS["JUV"]   = (JUV_MP_sheet,  JUV_carry_cap)
        if FIV_sim: _MP_CONFIGS["FIV"]   = (FIV_MP_sheet,  FIV_carry_cap)

        wt_arr = df_project_weather["time"].to_numpy(dtype="datetime64[ns]")
        try:
            seasonal_eligible = get_seasonal_eligible_indices(wt_arr, Sim_start_date)
            st.info(f"Seasonal eligible start timestamps: {len(seasonal_eligible):,}")
        except Exception as e:
            st.error(f"Seasonal index error: {e}"); st.stop()

        # ────────────────────────────────────────────────────────────────
        #  SECTION 3 – MP simulations
        # ────────────────────────────────────────────────────────────────
        st.subheader("3.  Foundation / Monopile Simulations")
        mp_weather_results = {}
        prog_mp = st.progress(0, text="Running MP simulations …")
        vessel_list = list(_MP_CONFIGS.keys())

        for vi,(vessel,(sheet,cap)) in enumerate(_MP_CONFIGS.items()):
            with st.spinner(f"MP – {vessel} …"):
                try:
                    tl = build_mp_timeline(io.BytesIO(excel_bytes), sheet, cap, N_WTG, htd)
                    tot,dwn,summ = simulate_weather_impacts_mp(
                        tl, df_project_weather, io.BytesIO(excel_bytes), sheet,
                        transit_distance=0.5, transit_speed=1.0,
                        simulations=int(Simulations), eligible_indices=seasonal_eligible, htd=htd)
                    mp_weather_results[vessel] = {"timeline":tl,"total":tot,"downtime":dwn,"summary":summ}
                    p50 = round(summ["Total_Project_Days"].median(),1)
                    st.success(f"✅  MP {vessel} — P50 total: {p50} days")
                except Exception as e:
                    st.error(f"MP {vessel} failed: {e}")
            prog_mp.progress((vi+1)/len(vessel_list))

        if not mp_weather_results:
            st.error("No MP results produced — stopping."); st.stop()

        st.pyplot(plot_cumulative({v:r["total"] for v,r in mp_weather_results.items()},
                                  "Cumulative MP Duration (With Weather)"))
        mp_rows = [{"Vessel":v,
                    "P50 Total (d)":round(r["summary"]["Total_Project_Days"].median(),1),
                    "P90 Total (d)":round(r["summary"]["Total_Project_Days"].quantile(.9),1),
                    "P50 Downtime (d)":round(r["summary"]["Total_Downtime_Days"].median(),1),
                    "P50 Active (d)":round(r["summary"]["Total_Active_Days"].median(),1)}
                   for v,r in mp_weather_results.items()]
        st.dataframe(pd.DataFrame(mp_rows).set_index("Vessel"))

        # MP→WTG offset per vessel
        mp_end_offsets = {}
        for vessel,result in mp_weather_results.items():
            td = pd.to_numeric(result["summary"]["Total_Project_Days"],errors="coerce").to_numpy(dtype=float)
            mp_end_offsets[vessel] = td + float(intermediate_days)

        # ────────────────────────────────────────────────────────────────
        #  SECTION 4 – WTG duration (no weather baseline)
        # ────────────────────────────────────────────────────────────────
        st.subheader("4.  WTG Installation – Timelines")
        wtg_duration_results = {}

        if W90_sim:
            with st.spinner("Building W90 WTG timeline …"):
                try:
                    tl = build_w90_timeline(io.BytesIO(excel_bytes), W90_WTG_sheet, W90_carry_cap,
                                            int(N_WTG), htd, float(dock_distance), float(W90_transit_speed))
                    dur = add_duration_simulations_w90(tl, io.BytesIO(excel_bytes), W90_WTG_sheet,
                                                       int(Simulations), 42, htd, barge,
                                                       float(dock_distance), float(barge_distance), float(W90_transit_speed))
                    wtg_duration_results["W90"] = {"timeline":tl,"duration":dur,"mp_offset_days":mp_end_offsets.get("W90")}
                    st.success("✅  W90 WTG timeline built")
                except Exception as e: st.error(f"W90 WTG build failed: {e}")

        for vessel, sheet, cap, spd in [("JUV",JUV_WTG_sheet,JUV_carry_cap,Comp_transit_speed),
                                         ("FIV",FIV_WTG_sheet,FIV_carry_cap,Fiv_transit_speed)]:
            flag = JUV_sim if vessel=="JUV" else FIV_sim
            if flag:
                with st.spinner(f"Building {vessel} WTG timeline …"):
                    try:
                        tl = build_sequential_wtg_timeline(io.BytesIO(excel_bytes), sheet, int(cap),
                                                           int(N_WTG), htd, float(dock_distance), float(spd))
                        dur = add_duration_simulations_sequential(tl, io.BytesIO(excel_bytes), sheet,
                                                                   int(Simulations), float(dock_distance), float(spd), 42, htd)
                        wtg_duration_results[vessel] = {"timeline":tl,"duration":dur,"mp_offset_days":mp_end_offsets.get(vessel)}
                        st.success(f"✅  {vessel} WTG timeline built")
                    except Exception as e: st.error(f"{vessel} WTG build failed: {e}")

        # ────────────────────────────────────────────────────────────────
        #  SECTION 5 – WTG weather simulations
        # ────────────────────────────────────────────────────────────────
        st.subheader("5.  WTG Installation – Weather Simulations")
        wtg_seasonal = get_seasonal_eligible_indices(wt_arr, Sim_start_date)
        wtg_weather_results = {}
        wtg_list = list(wtg_duration_results.keys())
        prog_wtg = st.progress(0, text="Running WTG weather simulations …")

        for vi,vessel in enumerate(wtg_list):
            res = wtg_duration_results[vessel]
            offset = res["mp_offset_days"]
            with st.spinner(f"WTG weather – {vessel} …"):
                try:
                    if vessel == "W90":
                        dur_data = res["duration"]
                        dock_df  = dur_data[0] if isinstance(dur_data,tuple) else dur_data
                        tot,dwn,summ = simulate_weather_impacts_w90(
                            dock_df, df_project_weather, df_barge_wind,
                            io.BytesIO(excel_bytes), W90_WTG_sheet,
                            float(dock_distance), float(W90_transit_speed), int(Simulations),
                            wtg_seasonal, 42, 5000, htd, offset)
                        wtg_weather_results["W90_dock"] = {"total":tot,"downtime":dwn,"summary":summ,"no_wx":dock_df}
                        if barge and isinstance(dur_data,tuple):
                            bdf=dur_data[1]
                            tot_b,dwn_b,summ_b = simulate_weather_impacts_w90(
                                bdf, df_project_weather, df_barge_wind,
                                io.BytesIO(excel_bytes), W90_WTG_sheet,
                                float(barge_distance), float(W90_transit_speed), int(Simulations),
                                wtg_seasonal, 42, 5000, htd, offset)
                            wtg_weather_results["W90_barge"] = {"total":tot_b,"downtime":dwn_b,"summary":summ_b,"no_wx":bdf}
                    else:
                        spd = float(Comp_transit_speed) if vessel=="JUV" else float(Fiv_transit_speed)
                        sheet = JUV_WTG_sheet if vessel=="JUV" else FIV_WTG_sheet
                        tot,dwn,summ = simulate_sequential_wtg_weather(
                            res["duration"], df_project_weather,
                            io.BytesIO(excel_bytes), sheet,
                            float(dock_distance), spd, int(Simulations),
                            wtg_seasonal, 42, 5000, htd, offset)
                        wtg_weather_results[vessel] = {"total":tot,"downtime":dwn,"summary":summ,"no_wx":res["duration"]}
                    p50 = round(summ["Total_Project_Days"].median(),1)
                    st.success(f"✅  WTG {vessel} weather — P50 total: {p50} days")
                except Exception as e:
                    st.error(f"WTG {vessel} weather failed: {e}")
            prog_wtg.progress((vi+1)/len(wtg_list))

        if not wtg_weather_results:
            st.warning("No WTG weather results produced."); st.stop()

        # ── Final outputs ─────────────────────────────────────────────────
        st.markdown("---")
        st.header("📊  Final Results")

        st.subheader("Cumulative WTG Campaign Duration (With Weather)")
        st.pyplot(plot_cumulative({v:r["total"] for v,r in wtg_weather_results.items()},
                                  "Cumulative WTG Duration (With Weather)"))

        st.subheader("WTG Project Duration Summary")
        wtg_rows = []
        for v,r in wtg_weather_results.items():
            s=r["summary"]
            tot_p50=s["Total_Project_Days"].median(); dwn_p50=s["Total_Downtime_Days"].median()
            wtg_rows.append({"Vessel":v,
                "P50 Total (d)":   round(tot_p50,1),
                "P90 Total (d)":   round(s["Total_Project_Days"].quantile(.9),1),
                "P50 Downtime (d)":round(dwn_p50,1),
                "Downtime %":      round(dwn_p50/tot_p50*100 if tot_p50>0 else 0,1),
                "P50 Active (d)":  round(s["Total_Active_Days"].median(),1)})
        st.dataframe(pd.DataFrame(wtg_rows).set_index("Vessel"))

        st.subheader("Full Project Summary (P0 = no-weather baseline)")
        full_frames = [build_project_summary(r["no_wx"],r["total"],r["downtime"],v)
                       for v,r in wtg_weather_results.items()]
        st.dataframe(pd.concat(full_frames,ignore_index=True))

        st.subheader("Campaign Duration Bar Chart")
        bar_rows = []
        for v,r in wtg_weather_results.items():
            s=r["summary"]
            tp=float(s["Total_Project_Days"].median()); dp=float(s["Total_Downtime_Days"].median())
            bar_rows.append({"Vessel":v,"Total_P50":round(tp,1),"Net_P50":round(tp-dp,1),
                             "Downtime_Pct":round(dp/tp*100 if tp>0 else 0,1)})
        st.pyplot(plot_bar_comparison(bar_rows))

        st.success("🎉  All simulations complete!")

# ─────────────────────────────────────────────────────────────────────────────
#  FOOTER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption("W90 Monte Carlo v6 · Streamlit conversion")
