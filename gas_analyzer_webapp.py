import streamlit as st
import pandas as pd
import math
from dataclasses import dataclass
import io
from datetime import datetime

# 1. SETUP & DATA
st.set_page_config(page_title="Gas Analyzer Pro", page_icon="🔥", layout="wide")

@dataclass
class Component:
    name: str
    formula: str
    mw: float
    lhv_mass: float # MJ/kg (ISO 6976 @ 15°C)
    hhv_mass: float # MJ/kg (ISO 6976 @ 15°C)
    o2_stoic: float # Moles of O2 per mole of fuel

COMPONENTS = {
    'Methane': Component('Methane', 'CH4', 16.043, 50.009, 55.503, 2.0),
    'Ethane':  Component('Ethane', 'C2H6', 30.070, 47.794, 51.901, 3.5),
    'Propane': Component('Propane', 'C3H8', 44.097, 46.357, 50.366, 5.0),
    'n-Butane': Component('n-Butane', 'C4H10', 58.123, 45.752, 49.512, 6.5),
    'i-Butane': Component('i-Butane', 'C4H10', 58.123, 45.614, 49.375, 6.5),
    'Hydrogen': Component('Hydrogen', 'H2', 2.016, 119.95, 141.86, 0.5),
    'CO':       Component('Carbon Monoxide', 'CO', 28.010, 10.103, 10.103, 0.5),
    'H2S':      Component('Hydrogen Sulfide', 'H2S', 34.081, 15.208, 16.532, 1.5),
    'Nitrogen': Component('Nitrogen', 'N2', 28.013, 0.0, 0.0, 0.0),
    'CO2':      Component('Carbon Dioxide', 'CO2', 44.010, 0.0, 0.0, 0.0),
}

MW_AIR = 28.9645
MOLAR_VOLUME_15C = 23.6443 # m3/kmol @ 15°C

# 2. CALCULATION ENGINE (ISO 6976 Math)
def calculate_properties(comp_percent):
    total_raw = sum(comp_percent.values())
    if total_raw == 0: return None
    comp = {k: (v / total_raw) for k, v in comp_percent.items()}
    
    # Molar Mass
    mw_mix = sum(comp[n] * COMPONENTS[n].mw for n in comp)
    sg = mw_mix / MW_AIR
    
    # SI Units (Metric)
    dens_si = mw_mix / MOLAR_VOLUME_15C
    lhv_m_si = sum((comp[n] * COMPONENTS[n].mw / mw_mix) * COMPONENTS[n].lhv_mass for n in comp)
    hhv_m_si = sum((comp[n] * COMPONENTS[n].mw / mw_mix) * COMPONENTS[n].hhv_mass for n in comp)
    lhv_v_si = lhv_m_si * dens_si
    hhv_v_si = hhv_m_si * dens_si
    wi_l_si = lhv_v_si / math.sqrt(sg)
    wi_h_si = hhv_v_si / math.sqrt(sg)
    
    # US Customary Units (Imperial)
    # Conversion: MJ/m3 to Btu/scf (26.839), MJ/kg to Btu/lb (429.9)
    dens_us = dens_si * 0.06242796
    lhv_m_us = lhv_m_si * 429.92
    hhv_m_us = hhv_m_si * 429.92
    lhv_v_us = lhv_v_si * 26.839
    hhv_v_us = hhv_v_si * 26.839
    wi_l_us = wi_l_si * 26.839
    wi_h_us = wi_h_si * 26.839

    # Component-Specific Flags
    h2_pct = comp.get('Hydrogen', 0) * 100
    co2_n2_pct = (comp.get('Carbon Dioxide', 0) + comp.get('Nitrogen', 0)) * 100
    h2s_ppm = comp.get('H2S', 0) * 1000000
    
    # Stoichiometric Air-Fuel Ratio
    o2_req = sum(comp[n] * COMPONENTS[n].o2_stoic for n in comp)
    afr = (o2_req / 0.20947) * (MW_AIR / mw_mix)

    return {
        "composition": comp, "mw": mw_mix, "sg": sg,
        "dens_si": dens_si, "dens_us": dens_us,
        "lhv_m_si": lhv_m_si, "lhv_m_us": lhv_m_us,
        "lhv_v_si": lhv_v_si, "lhv_v_us": lhv_v_us,
        "hhv_m_si": hhv_m_si, "hhv_m_us": hhv_m_us,
        "hhv_v_si": hhv_v_si, "hhv_v_us": hhv_v_us,
        "wi_l_si": wi_l_si, "wi_l_us": wi_l_us,
        "wi_h_si": wi_h_si, "wi_h_us": wi_h_us,
        "h2": h2_pct, "co2_n2": co2_n2_pct, "h2s": h2s_ppm,
        "mn": 100 - (comp.get('Ethane', 0)*100*0.5), # Simple approx
        "afr": afr, "aft_c": 1500, "aft_f": 2732 # Placeholders
    }

# 3. INITIALIZE STATE
if 'results' not in st.session_state: st.session_state.results = {}
if 'use_si' not in st.session_state: st.session_state.use_si = True

# 4. UI - SIDEBAR & INPUTS
st.sidebar.title("Configuration")
st.session_state.use_si = st.sidebar.toggle("Use SI Units", value=True)

st.title("Gas Turbine Fuel Analyzer")
tabs = st.tabs(["Input", "Results"])

with tabs[0]:
    col1, col2 = st.columns(2)
    comp_input = {}
    items = list(COMPONENTS.items())
    for i, (name, obj) in enumerate(items):
        with col1 if i < len(items)/2 else col2:
            comp_input[name] = st.number_input(f"{name} (mol%)", 0.0, 100.0, 0.0, step=0.1, key=f"in_{name}")
    
    # FIXED INDENTATION FOR BUTTON
    if st.button("CALCULATE PROPERTIES", type="primary", use_container_width=True):
        res = calculate_properties(comp_input)
        if res:
            st.session_state.results = res
            st.success("Calculated successfully!")

# 5. UI - RESULTS
with tabs[1]:
    if not st.session_state.results:
        st.info("Please calculate first.")
    else:
        r = st.session_state.results
        si = st.session_state.use_si
        
        # Display table mapping exactly to the dict keys
        res_data = [
            ["Molecular Weight", f"{r['mw']:.3f}", "g/mol"],
            ["Specific Gravity", f"{r['sg']:.4f}", "-"],
            ["Density", f"{r['dens_si' if si else 'dens_us']:.4f}", "kg/m3" if si else "lb/ft3"],
            ["LHV (Vol)", f"{r['lhv_v_si' if si else 'lhv_v_us']:.2f}", "MJ/m3" if si else "Btu/scf"],
            ["Wobbe Index", f"{r['wi_l_si' if si else 'wi_l_us']:.2f}", "MJ/m3" if si else "Btu/scf"],
            ["Air/Fuel Ratio", f"{r['afr']:.2f}", "kg/kg"]
        ]
        st.table(pd.DataFrame(res_data, columns=["Property", "Value", "Unit"]))
