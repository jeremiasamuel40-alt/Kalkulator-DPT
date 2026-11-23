import io
import math
from datetime import datetime
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import streamlit as st

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as XLImage

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

# ==============================================================================
# KONFIGURASI HALAMAN
# ==============================================================================
st.set_page_config(
    page_title="Kalkulator DPT - SNI 8460 & SNI 2847",
    page_icon="🧱",
    layout="wide"
)

st.title("🧱 Kalkulator Analisis & Desain Dinding Penahan Tanah (DPT)")
st.caption("Standar: SNI 8460:2017 (Geoteknik), SNI 2847:2019 (Beton Bertulang), SNI 1727:2020 (Pembebanan)")

# ==============================================================================
# SIDEBAR INPUT PARAMETER
# ==============================================================================
with st.sidebar:
    st.header("1. Identitas Proyek")
    project_name = st.text_input("Nama Proyek", "Pembangunan Kawasan Industri & Perimeter Wall")
    wall_type_name = st.text_input("Kode / Tipe DPT", "DPT-1 (H_exp = 2.51 m)")
    engineer_name = st.text_input("Perencana Struktur", "Jeremia Samuel, S.T.")

    st.header("2. Geometri Penampang (m)")
    col_g1, col_g2 = st.columns(2)
    with col_g1:
        H_exp = st.number_input("Tinggi Ekspos (H_exp)", min_value=0.2, max_value=12.0, value=2.51, step=0.05)
        t_top = st.number_input("Tebal Puncak Stem", min_value=0.15, max_value=1.5, value=0.26, step=0.01)
        B_toe = st.number_input("Panjang Toe (0 = Tipe L)", min_value=0.0, max_value=5.0, value=0.00, step=0.05)
        t_slab_base = st.number_input("Tebal Slab Pangkal", min_value=0.15, max_value=1.5, value=0.35, step=0.05)
    with col_g2:
        D_f = st.number_input("Kedalaman Tertanam (Df)", min_value=0.3, max_value=5.0, value=1.00, step=0.05)
        t_bot = st.number_input("Tebal Dasar Stem", min_value=0.15, max_value=2.0, value=0.35, step=0.01)
        B_heel = st.number_input("Panjang Heel", min_value=0.2, max_value=8.0, value=1.65, step=0.05)
        t_slab_tip = st.number_input("Tebal Slab Ujung", min_value=0.15, max_value=1.5, value=0.25, step=0.05)

    use_shear_key = st.checkbox("Gunakan Shear Key (Kunci Geser)", value=True)
    if use_shear_key:
        col_sk1, col_sk2 = st.columns(2)
        with col_sk1:
            D_key = st.number_input("Kedalaman Key (m)", min_value=0.15, max_value=2.0, value=0.40, step=0.05)
        with col_sk2:
            B_key = st.number_input("Lebar Key (m)", min_value=0.15, max_value=1.5, value=0.30, step=0.05)
    else:
        D_key = 0.0
        B_key = 0.0

    st.header("3. Pagar Atas & Beban Lalu Lintas")
    use_fence = st.checkbox("Ada Pagar Precast di Puncak DPT", value=True)
    if use_fence:
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            H_panel = st.number_input("Tinggi Panel Pagar (m)", min_value=0.5, max_value=4.0, value=2.00, step=0.1)
            q_wind = st.number_input("Tekanan Angin (kPa)", min_value=0.0, max_value=3.0, value=0.60, step=0.05)
        with col_f2:
            H_wire = st.number_input("Tinggi Kawat Duri (m)", min_value=0.0, max_value=1.5, value=0.50, step=0.1)
            d_pocket = st.number_input("Kedalaman Pocket (m)", min_value=0.2, max_value=1.0, value=0.50, step=0.05)
    else:
        H_panel, H_wire, q_wind, d_pocket = 0.0, 0.0, 0.0, 0.0

    q_surch = st.number_input("Beban Surcharge Jalan (kPa)", min_value=0.0, max_value=50.0, value=10.0, step=1.0)

    st.header("4. Parameter Tanah & Gempa")
    col_s1, col_s2 = st.columns(2)
    with col_s1:
        gamma_1 = st.number_input("γ Tanah Urug (kN/m³)", 14.0, 22.0, 18.0, 0.5)
        phi_1 = st.number_input("φ Tanah Urug (°)", 15.0, 45.0, 30.0, 1.0)
        c_1 = st.number_input("c Tanah Urug (kPa)", 0.0, 50.0, 0.0, 1.0)
        q_all = st.number_input("Daya Dukung Izin q_all (kPa)", 30.0, 600.0, 150.0, 10.0)
        mu_base = st.number_input("Koef. Gesek Dasar (μ)", 0.20, 0.75, 0.45, 0.05)
    with col_s2:
        gamma_2 = st.number_input("γ Tanah Dasar (kN/m³)", 14.0, 22.0, 18.5, 0.5)
        phi_2 = st.number_input("φ Tanah Dasar (°)", 15.0, 45.0, 32.0, 1.0)
        c_2 = st.number_input("c Tanah Dasar (kPa)", 0.0, 100.0, 5.0, 1.0)
        SF_bc_target = st.number_input("SF Daya Dukung Target", 2.0, 4.0, 3.0, 0.5)
        k_h = st.number_input("Koef. Gempa Horizontal (kh)", 0.00, 0.50, 0.15, 0.01)

    use_passive_front = st.checkbox("Perhitungkan Tekanan Pasif Depan (Df)", value=True)
    R_pasif = st.slider("Faktor Reduksi Pasif Depan", 0.0, 1.0, 0.50, 0.05) if use_passive_front else 0.0

    st.header("5. Material & Konfigurasi Tulangan")
    col_m1, col_m2 = st.columns(2)
    with col_m1:
        fc = st.number_input("Mutu Beton fc' (MPa)", 17.5, 50.0, 25.0, 2.5)
        cc_wall = st.number_input("Selimut Dinding (mm)", 30, 75, 40, 5)
        D_wall = st.selectbox("Dia. Tul. Dinding (mm)", [10, 13, 16, 19, 22, 25], index=2)
        D_heel = st.selectbox("Dia. Tul. Heel (mm)", [10, 13, 16, 19, 22, 25], index=2)
        D_toe = st.selectbox("Dia. Tul. Toe (mm)", [10, 13, 16, 19, 22, 25], index=1)
    with col_m2:
        fy = st.number_input("Mutu Baja fy (MPa)", 240.0, 550.0, 420.0, 20.0)
        cc_slab = st.number_input("Selimut Slab Bawah (mm)", 40, 100, 75, 5)
        s_wall = st.number_input("Spasi Tul. Dinding (mm)", 75, 300, 150, 25)
        s_heel = st.number_input("Spasi Tul. Heel (mm)", 75, 300, 150, 25)
        s_toe = st.number_input("Spasi Tul. Toe (mm)", 75, 300, 150, 25)

# ==============================================================================
# MESIN PERHITUNGAN GEOTEKNIK & STRUKTUR
# ==============================================================================
def run_calculations():
    gamma_c = 24.0  # kN/m3 berat jenis beton bertulang
    B = B_toe + t_bot + B_heel
    H_tot = H_exp + D_f + t_slab_base
    H_stem = H_exp + D_f

    # 1. Koefisien Tekanan Tanah (Rankine & Mononobe-Okabe)
    phi1_rad = math.radians(phi_1)
    phi2_rad = math.radians(phi_2)
    Ka = math.tan(math.radians(45.0 - phi_1 / 2.0)) ** 2
    Kp = math.tan(math.radians(45.0 + phi_2 / 2.0)) ** 2

    # Mononobe-Okabe (SNI 8460:2017)
    delta_mo = (2.0 / 3.0) * phi1_rad
    psi = math.atan(k_h)
    seismic_valid = phi1_rad > psi
    if k_h > 0 and seismic_valid:
        num = math.cos(phi1_rad - psi) ** 2
        sq_term = math.sqrt((math.sin(phi1_rad + delta_mo) * math.sin(phi1_rad - psi)) / math.cos(delta_mo + psi))
        den = math.cos(psi) * math.cos(delta_mo + psi) * ((1.0 + sq_term) ** 2)
        Kae = num / den
        dKae = max(0.0, Kae - Ka)
    else:
        Kae = Ka
        dKae = 0.0

    # 2. Beban Vertikal & Momen Penahan terhadap Ujung Depan (x = 0 di ujung Toe)
    # Rincian komponen gaya per 1.0 m lebar DPT: (Nama, Gaya V [kN/m], Lengan x [m])
    v_items = []
    # W1: Stem Persegi
    w1 = t_top * H_stem * gamma_c
    x1 = B_toe + (t_bot - t_top) + t_top / 2.0
    v_items.append(("W1 - Dinding (Stem Persegi)", w1, x1))

    # W2: Stem Segitiga (jika tapered di sisi depan)
    w2 = 0.5 * (t_bot - t_top) * H_stem * gamma_c
    x2 = B_toe + (2.0 / 3.0) * (t_bot - t_top)
    v_items.append(("W2 - Dinding (Stem Miring)", w2, x2))

    # W3: Base Slab Persegi
    w3 = B * t_slab_tip * gamma_c
    x3 = B / 2.0
    v_items.append(("W3 - Pelat Dasar (Persegi)", w3, x3))

    # W4: Base Slab Tapered (di bawah heel)
    dt_slab = max(0.0, t_slab_base - t_slab_tip)
    w4 = 0.5 * B_heel * dt_slab * gamma_c
    x4 = B_toe + t_bot + (1.0 / 3.0) * B_heel
    v_items.append(("W4 - Pelat Dasar (Miring Heel)", w4, x4))

    # W5: Shear Key
    w5 = B_key * D_key * gamma_c if use_shear_key else 0.0
    x5 = B_toe + t_bot - B_key / 2.0
    v_items.append(("W5 - Shear Key Beton", w5, x5))

    # W6: Tanah Urug di atas Heel (persegi + segitiga akibat kemiringan slab atas)
    w6a = B_heel * H_stem * gamma_1
    x6a = B_toe + t_bot + B_heel / 2.0
    w6b = 0.5 * B_heel * dt_slab * gamma_1
    x6b = B_toe + t_bot + (2.0 / 3.0) * B_heel
    w6 = w6a + w6b
    x6 = (w6a * x6a + w6b * x6b) / w6 if w6 > 0 else x6a
    v_items.append(("W6 - Tanah Urug di atas Heel", w6, x6))

    # W7: Beban Surcharge di atas Heel
    w7 = q_surch * B_heel
    x7 = B_toe + t_bot + B_heel / 2.0
    v_items.append(("W7 - Beban Surcharge di atas Heel", w7, x7))

    # W8: Tanah Penutup di atas Toe
    w8 = B_toe * D_f * gamma_2
    x8 = B_toe / 2.0
    v_items.append(("W8 - Tanah Penutup di atas Toe", w8, x8))

    # W9: Berat Sendiri Pagar Precast (estimasi panel 50mm + kolom)
    w9 = (H_panel * 0.05 * gamma_c + 0.35) if use_fence else 0.0
    x9 = B_toe + (t_bot - t_top) + t_top / 2.0
    v_items.append(("W9 - Berat Pagar Precast Atas", w9, x9))

    df_v = pd.DataFrame(v_items, columns=["Komponen", "Gaya_V (kN/m)", "Lengan_x (m)"])
    df_v["Momen_Penahan (kNm/m)"] = df_v["Gaya_V (kN/m)"] * df_v["Lengan_x (m)"]

    sum_V_with_q = df_v["Gaya_V (kN/m)"].sum()
    sum_MR_with_q = df_v["Momen_Penahan (kNm/m)"].sum()

    # Untuk cek guling/geser konservatif, surcharge di atas heel sering diabaikan pada gaya penahan:
    sum_V_no_q = sum_V_with_q - w7
    sum_MR_no_q = sum_MR_with_q - (w7 * x7)

    # 3. Gaya Horizontal & Momen Guling (y = 0 di dasar slab)
    Pa = 0.5 * Ka * gamma_1 * (H_tot ** 2)
    ya = H_tot / 3.0

    Pq = Ka * q_surch * H_tot
    yq = H_tot / 2.0

    H_fence_tot = (H_panel + H_wire) if use_fence else 0.0
    Pw = q_wind * H_fence_tot
    yw = H_tot + H_fence_tot / 2.0

    # Dinamik gempa
    dPae = 0.5 * dKae * gamma_1 * (H_tot ** 2)
    ydae = 0.6 * H_tot
    W_conc_tot = w1 + w2 + w3 + w4 + w5 + w9
    Pir = k_h * W_conc_tot
    yir = 0.45 * H_tot  # pendekatan titik berat beton

    # Tahanan Geser Dasar & Pasif
    # Jika memakai shear key, sebagian bidang geser bergeser ke tanah-tanah (tan(phi_2))
    mu_eff = math.tan(phi2_rad) if use_shear_key else mu_base
    Pp_front = R_pasif * 0.5 * Kp * gamma_2 * (D_f ** 2)
    Pp_key = (0.5 * Kp * gamma_2 * ((D_f + D_key) ** 2 - (D_f ** 2)) + 2.0 * c_2 * math.sqrt(Kp) * D_key) if use_shear_key else 0.0
    adh_base = 0.5 * c_2 * B

    # --- KASUS 1: STATIS + ANGIN ---
    Fh_stat = Pa + Pq + Pw
    Mo_stat = Pa * ya + Pq * yq + Pw * yw
    Fr_stat = mu_eff * sum_V_no_q + adh_base + Pp_front + Pp_key

    SF_ot_stat = sum_MR_no_q / Mo_stat if Mo_stat > 0 else 99.0
    SF_sl_stat = Fr_stat / Fh_stat if Fh_stat > 0 else 99.0

    # Daya dukung menggunakan beban vertikal total (termasuk surcharge) agar q_max maksimum
    M_net_stat = sum_MR_with_q - Mo_stat
    xR_stat = M_net_stat / sum_V_with_q
    e_stat = (B / 2.0) - xR_stat
    if e_stat <= B / 6.0:
        q_max_stat = (sum_V_with_q / B) * (1.0 + 6.0 * e_stat / B)
        q_min_stat = (sum_V_with_q / B) * (1.0 - 6.0 * e_stat / B)
        uplift_stat = False
    else:
        q_max_stat = (2.0 * sum_V_with_q) / (3.0 * max(0.05, xR_stat))
        q_min_stat = 0.0
        uplift_stat = True

    q_ult = q_all * SF_bc_target
    SF_bc_stat = q_ult / q_max_stat if q_max_stat > 0 else 99.0
    SF_gov_stat = min(SF_ot_stat, SF_sl_stat, SF_bc_stat)

    # --- KASUS 2: GEMPA (MONONOBE-OKABE, TANPA ANGIN EKSTREM) ---
    Fh_seis = Pa + Pq + dPae + Pir
    Mo_seis = Pa * ya + Pq * yq + dPae * ydae + Pir * yir
    Fr_seis = mu_eff * sum_V_no_q + adh_base + Pp_front + Pp_key

    SF_ot_seis = sum_MR_no_q / Mo_seis if Mo_seis > 0 else 99.0
    SF_sl_seis = Fr_seis / Fh_seis if Fh_seis > 0 else 99.0

    M_net_seis = sum_MR_with_q - Mo_seis
    xR_seis = M_net_seis / sum_V_with_q
    e_seis = (B / 2.0) - xR_seis
    if e_seis <= B / 6.0:
        q_max_seis = (sum_V_with_q / B) * (1.0 + 6.0 * e_seis / B)
        q_min_seis = (sum_V_with_q / B) * (1.0 - 6.0 * e_seis / B)
    else:
        q_max_seis = (2.0 * sum_V_with_q) / (3.0 * max(0.05, xR_seis))
        q_min_seis = 0.0

    SF_bc_seis = q_ult / q_max_seis if q_max_seis > 0 else 99.0
    SF_gov_seis = min(SF_ot_seis, SF_sl_seis, SF_bc_seis)

    # ==========================================================================
    # 4. PERHITUNGAN PENULANGAN BETON BERTULANG (SNI 2847:2019)
    # ==========================================================================
    phi_flex = 0.90
    phi_shear = 0.75
    rho_min = 0.0018 if fy >= 420.0 else 0.0020

    def check_section_rebar(Mu, Vu, h_mm, cc_mm, db_mm, sp_mm):
        b = 1000.0  # per 1 meter lebar
        d = max(50.0, h_mm - cc_mm - 0.5 * db_mm)
        # Lentur
        Rn = (Mu * 1e6) / (phi_flex * b * (d ** 2))
        term = 1.0 - (2.0 * Rn) / (0.85 * fc)
        if term <= 0:
            rho_req = 0.05  # section too thin
        else:
            rho_req = (0.85 * fc / fy) * (1.0 - math.sqrt(term))
        rho_design = max(rho_req, rho_min)
        As_req = rho_design * b * d
        As_prov = (1000.0 / sp_mm) * (0.25 * math.pi * (db_mm ** 2))
        a = (As_prov * fy) / (0.85 * fc * b)
        phi_Mn = phi_flex * As_prov * fy * (d - a / 2.0) / 1e6  # kNm/m
        dc_flex = Mu / phi_Mn if phi_Mn > 0 else 99.0

        # Geser 1 arah tanpa sengkang (SNI 2847:2019 Pasal 22.5.5.1)
        phi_Vc = phi_shear * 0.17 * 1.0 * math.sqrt(fc) * b * d / 1000.0  # kN/m
        dc_shear = Vu / phi_Vc if phi_Vc > 0 else 99.0

        return {
            "h_mm": h_mm, "d_mm": d, "Mu": Mu, "Vu": Vu,
            "As_req": As_req, "As_prov": As_prov, "phi_Mn": phi_Mn,
            "dc_flex": dc_flex, "phi_Vc": phi_Vc, "dc_shear": dc_shear,
            "status": "OK" if (dc_flex <= 1.0 and dc_shear <= 1.0 and term > 0) else "NOT OK"
        }

    # A. Dinding (Stem) - Penampang kritis di atas base slab (tinggi = H_stem)
    # Kombinasi 1: 1.6 H + 1.0 W
    Pa_stem = 0.5 * Ka * gamma_1 * (H_stem ** 2)
    Pq_stem = Ka * q_surch * H_stem
    Mu_stem_1 = 1.6 * (Pa_stem * (H_stem / 3.0) + Pq_stem * (H_stem / 2.0)) + 1.0 * (Pw * (H_stem + H_fence_tot / 2.0))
    # Kombinasi 2: 1.0 H + 1.0 E
    dPae_stem = 0.5 * dKae * gamma_1 * (H_stem ** 2)
    Pir_stem = k_h * (w1 + w2 + w9)
    Mu_stem_2 = 1.0 * (Pa_stem * (H_stem / 3.0) + Pq_stem * (H_stem / 2.0)) + 1.0 * (dPae_stem * (0.6 * H_stem) + Pir_stem * (0.5 * H_stem))
    Mu_stem = max(Mu_stem_1, Mu_stem_2)

    d_stem_m = (t_bot * 1000.0 - cc_wall - 0.5 * D_wall) / 1000.0
    H_v_stem = max(0.1, H_stem - d_stem_m)
    Vu_stem_1 = 1.6 * (0.5 * Ka * gamma_1 * (H_v_stem ** 2) + Ka * q_surch * H_v_stem) + 1.0 * Pw
    Vu_stem_2 = 1.0 * (0.5 * Ka * gamma_1 * (H_v_stem ** 2) + Ka * q_surch * H_v_stem + 0.5 * dKae * gamma_1 * (H_v_stem ** 2) + Pir_stem)
    Vu_stem = max(Vu_stem_1, Vu_stem_2)

    rebar_stem = check_section_rebar(Mu_stem, Vu_stem, t_bot * 1000.0, cc_wall, D_wall, s_wall)

    # Kontrol Lokal Pocket Pagar di Puncak Stem (jika ada pagar)
    if use_fence and Pw > 0:
        Mu_pocket = 1.6 * Pw * (H_fence_tot / 2.0 + d_pocket / 2.0)
        t_pocket_wall = max(50.0, (t_top * 1000.0 - 170.0) / 2.0)  # asumsi lebar pocket 170 mm seperti gambar
        rebar_pocket = check_section_rebar(Mu_pocket, 1.6 * Pw, t_pocket_wall, 25.0, 10, 150)
    else:
        rebar_pocket = None

    # B. Slab Belakang (Heel) - Penampang kritis di muka dalam Stem (panjang = B_heel)
    # Beban ke bawah (tarik serat atas): 1.2*(tanah + slab heel) + 1.6*(surcharge) - 0.9*(reaksi tanah ke atas)
    W_slab_heel = (B_heel * t_slab_tip + 0.5 * B_heel * dt_slab) * gamma_c
    x_slab_heel = (B_heel * t_slab_tip * gamma_c * (B_heel / 2.0) + 0.5 * B_heel * dt_slab * gamma_c * (B_heel / 3.0)) / W_slab_heel
    x_soil_heel = x6 - (B_toe + t_bot)  # jarak dari muka dalam stem

    # Reaksi tanah ke atas di bawah heel (ambil distribusi statis minimum agar konservatif)
    q_heel_tip = q_min_stat
    q_heel_face = q_min_stat + (q_max_stat - q_min_stat) * (B_heel / B)
    F_up_heel = 0.5 * (q_heel_tip + q_heel_face) * B_heel
    x_up_heel = (B_heel / 3.0) * (2.0 * q_heel_tip + q_heel_face) / max(1e-6, (q_heel_tip + q_heel_face))

    Vu_heel = max(0.0, 1.2 * (w6 + W_slab_heel) + 1.6 * w7 - 0.9 * F_up_heel)
    Mu_heel = max(0.0, 1.2 * (w6 * x_soil_heel + W_slab_heel * x_slab_heel) + 1.6 * (w7 * (B_heel / 2.0)) - 0.9 * (F_up_heel * x_up_heel))
    rebar_heel = check_section_rebar(Mu_heel, Vu_heel, t_slab_base * 1000.0, cc_slab, D_heel, s_heel)

    # C. Slab Depan (Toe) - Penampang kritis di muka luar Stem (panjang = B_toe)
    if B_toe > 0.05:
        q_max_env = max(q_max_stat, q_max_seis)
        q_min_env = min(q_min_stat, q_min_seis)
        q_toe_face = q_max_env - (q_max_env - q_min_env) * (B_toe / B)
        # Beban ke atas (tarik serat bawah): 1.6*tekanan tanah - 0.9*(berat slab toe + tanah atas toe)
        W_slab_toe = B_toe * t_slab_base * gamma_c
        F_up_toe = 0.5 * (q_max_env + q_toe_face) * B_toe
        x_up_toe = (B_toe / 3.0) * (2.0 * q_max_env + q_toe_face) / max(1e-6, (q_max_env + q_toe_face))

        Mu_toe = max(0.0, 1.4 * (F_up_toe * x_up_toe) - 0.9 * ((W_slab_toe + w8) * (B_toe / 2.0)))
        d_toe_m = (t_slab_base * 1000.0 - cc_slab - 0.5 * D_toe) / 1000.0
        L_v_toe = max(0.0, B_toe - d_toe_m)
        Vu_toe = max(0.0, 1.4 * (0.5 * (q_max_env + q_toe_face) * L_v_toe) - 0.9 * ((W_slab_toe + w8) * (L_v_toe / B_toe)))
        rebar_toe = check_section_rebar(Mu_toe, Vu_toe, t_slab_base * 1000.0, cc_slab, D_toe, s_toe)
    else:
        rebar_toe = None

    return {
        "B": B, "H_tot": H_tot, "H_stem": H_stem, "Ka": Ka, "Kp": Kp, "Kae": Kae, "dKae": dKae,
        "seismic_valid": seismic_valid, "df_v": df_v,
        "sum_V_with_q": sum_V_with_q, "sum_MR_with_q": sum_MR_with_q,
        "sum_V_no_q": sum_V_no_q, "sum_MR_no_q": sum_MR_no_q,
        "Pa": Pa, "Pq": Pq, "Pw": Pw, "dPae": dPae, "Pir": Pir,
        "Pp_front": Pp_front, "Pp_key": Pp_key, "mu_eff": mu_eff,
        "Fh_stat": Fh_stat, "Mo_stat": Mo_stat, "Fr_stat": Fr_stat,
        "SF_ot_stat": SF_ot_stat, "SF_sl_stat": SF_sl_stat, "SF_bc_stat": SF_bc_stat,
        "SF_gov_stat": SF_gov_stat, "e_stat": e_stat, "q_max_stat": q_max_stat, "q_min_stat": q_min_stat,
        "uplift_stat": uplift_stat,
        "Fh_seis": Fh_seis, "Mo_seis": Mo_seis, "Fr_seis": Fr_seis,
        "SF_ot_seis": SF_ot_seis, "SF_sl_seis": SF_sl_seis, "SF_bc_seis": SF_bc_seis,
        "SF_gov_seis": SF_gov_seis, "e_seis": e_seis, "q_max_seis": q_max_seis, "q_min_seis": q_min_seis,
        "rebar_stem": rebar_stem, "rebar_heel": rebar_heel, "rebar_toe": rebar_toe,
        "rebar_pocket": rebar_pocket
    }

res = run_calculations()

# ==============================================================================
# FUNGSI GENERATE SKETSA PENAMPANG & DIAGRAM TEKANAN
# ==============================================================================
def create_cross_section_plot(res):
    fig, ax = plt.subplots(figsize=(8.5, 6.2), dpi=160)
    B = res["B"]
    H_tot = res["H_tot"]
    H_stem = res["H_stem"]

    # Koordinat Poligon Beton DPT (x=0 di ujung Toe)
    poly_pts = [
        (0.0, 0.0),
        (B, 0.0),
        (B, t_slab_tip),
        (B_toe + t_bot, t_slab_base),
        (B_toe + t_bot, H_tot),
        (B_toe + (t_bot - t_top), H_tot),
        (B_toe, t_slab_base),
        (0.0, t_slab_base),
    ]
    wall_poly = patches.Polygon(poly_pts, closed=True, facecolor="#D3D7DC", edgecolor="#1F2937", linewidth=1.8, label="Beton DPT")
    ax.add_patch(wall_poly)

    # Shear Key
    if use_shear_key and D_key > 0:
        sk_x0 = B_toe + t_bot - B_key
        sk_rect = patches.Rectangle((sk_x0, -D_key), B_key, D_key, facecolor="#B0B7C0", edgecolor="#1F2937", linewidth=1.5)
        ax.add_patch(sk_rect)

    # Pagar Precast di Puncak
    if use_fence and H_panel > 0:
        fx0 = B_toe + (t_bot - t_top) + (t_top - 0.17) / 2.0
        fence_rect = patches.Rectangle((fx0, H_tot), 0.17, H_panel, facecolor="#E5E7EB", edgecolor="#374151", linewidth=1.3, hatch="//")
        ax.add_patch(fence_rect)
        # Kawat duri
        if H_wire > 0:
            ax.plot([fx0 + 0.085, fx0 + 0.085], [H_tot + H_panel, H_tot + H_panel + H_wire], color="#4B5563", lw=1.5, ls="--")
            ax.add_patch(patches.Ellipse((fx0 + 0.085, H_tot + H_panel + H_wire / 2.0), 0.22, H_wire, fill=False, edgecolor="#4B5563", lw=1.2))

    # Garis Tanah Belakang (Backfill) & Depan (Outer)
    ax.plot([B_toe + t_bot, B + 1.2], [H_tot, H_tot], color="#78350F", lw=2.2)
    ax.fill_between([B_toe + t_bot, B + 1.2], t_slab_base, H_tot, color="#FDE68A", alpha=0.35, label="Tanah Urug (Backfill)")

    front_ground_y = t_slab_base + D_f
    ax.plot([-1.0, B_toe], [front_ground_y, front_ground_y], color="#065F46", lw=2.2)
    ax.fill_between([-1.0, B_toe], 0.0, front_ground_y, color="#A7F3D0", alpha=0.35, label="Tanah Depan (Df)")

    # Diagram Tekanan Tanah Aktif (Segitiga di sisi kanan)
    scale_p = 0.8 / max(10.0, (res["Ka"] * gamma_1 * H_tot))
    p_base = res["Ka"] * gamma_1 * H_tot * scale_p
    x_diag = B + 0.25
    ax.plot([x_diag, x_diag, x_diag + p_base, x_diag], [0, H_tot, 0, 0], color="#DC2626", lw=1.5)
    ax.fill([x_diag, x_diag + p_base, x_diag], [H_tot, 0, 0], color="#FCA5A5", alpha=0.45)
    ax.annotate("", xy=(x_diag, H_tot / 3.0), xytext=(x_diag + p_base * 0.75, H_tot / 3.0),
                arrowprops=dict(arrowstyle="->", color="#DC2626", lw=1.8))
    ax.text(x_diag + p_base * 0.35, H_tot / 3.0 + 0.12, f"Pa = {res['Pa']:.1f} kN/m", color="#991B1B", fontsize=8.5, fontweight="bold")

    # Diagram Tegangan Tanah Dasar (Bearing Pressure di bawah slab)
    scale_q = 0.65 / max(50.0, res["q_max_stat"])
    q1_h = res["q_max_stat"] * scale_q
    q2_h = res["q_min_stat"] * scale_q
    y0_q = -D_key - 0.15
    ax.plot([0, B, B, 0, 0], [y0_q, y0_q, y0_q - q2_h, y0_q - q1_h, y0_q], color="#2563EB", lw=1.5)
    ax.fill([0, B, B, 0], [y0_q, y0_q, y0_q - q2_h, y0_q - q1_h], color="#93C5FD", alpha=0.45)
    ax.text(0.02, y0_q - q1_h - 0.18, f"q_max = {res['q_max_stat']:.1f} kPa", color="#1E40AF", fontsize=8.5, fontweight="bold")
    ax.text(B - 0.65, y0_q - q2_h - 0.18, f"q_min = {res['q_min_stat']:.1f} kPa", color="#1E40AF", fontsize=8.5)

    # Anotasi Dimensi Utama & Tulangan
    ax.text(B / 2.0, y0_q - max(q1_h, q2_h) - 0.42, f"Lebar Pondasi Total B = {B:.2f} m (Toe={B_toe:.2f}m, Heel={B_heel:.2f}m)",
            ha="center", fontsize=9, fontweight="bold", color="#111827")
    ax.text(B_toe + t_bot + 0.08, H_tot * 0.65,
            f"Tul. Dinding:\nD{D_wall}-{s_wall}\n(D/C={res['rebar_stem']['dc_flex']:.2f})",
            fontsize=8, color="#1F2937", bbox=dict(boxstyle="round,pad=0.25", fc="white", ec="#9CA3AF"))
    ax.text(B_toe + t_bot + B_heel * 0.45, t_slab_base + 0.15,
            f"Tul. Heel: D{D_heel}-{s_heel} (Atas)", fontsize=8, color="#1F2937",
            bbox=dict(boxstyle="round,pad=0.2", fc="white", ec="#9CA3AF"))

    ax.set_xlim(-1.1, B + 1.5)
    ax.set_ylim(y0_q - max(q1_h, q2_h) - 0.65, H_tot + (H_panel + H_wire if use_fence else 0) + 0.45)
    ax.set_aspect("equal")
    ax.grid(True, linestyle=":", alpha=0.5)
    ax.set_title(f"Sketsa Penampang & Distribusi Gaya — {wall_type_name}", fontsize=11, fontweight="bold")
    ax.set_xlabel("Lebar (m)")
    ax.set_ylabel("Elevasi Relatif (m)")
    plt.tight_layout()
    return fig

fig_cross = create_cross_section_plot(res)

# ==============================================================================
# TAMPILAN UTAMA (DASHBOARD ALA PLAXIS + DETAIL PERHITUNGAN)
# ==============================================================================
if not res["seismic_valid"]:
    st.error("⚠️ Peringatan Gempa: Sudut geser dalam tanah urug (φ1) lebih kecil dari sudut inersia gempa ψ = arctan(kh). Harap periksa parameter tanah atau kh.")

st.subheader("1. Ringkasan Safety Factor (Global Stability Summary)")

col_sf1, col_sf2 = st.columns(2)
with col_sf1:
    st.markdown("#### 🔹 Kondisi 1: Beban Statis Layan + Angin Pagar")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Governing SF", f"{res['SF_gov_stat']:.2f}", "OK" if res["SF_gov_stat"] >= 1.5 else "CHECK")
    m2.metric("SF Guling (≥ 2.0)", f"{res['SF_ot_stat']:.2f}", "OK" if res["SF_ot_stat"] >= 2.0 else "NOT OK")
    m3.metric("SF Geser (≥ 1.5)", f"{res['SF_sl_stat']:.2f}", "OK" if res["SF_sl_stat"] >= 1.5 else "NOT OK")
    m4.metric(f"SF Bearing (≥ {SF_bc_target:.1f})", f"{res['SF_bc_stat']:.2f}", "OK" if res["SF_bc_stat"] >= SF_bc_target else "NOT OK")

with col_sf2:
    st.markdown("#### 🔸 Kondisi 2: Beban Gempa (Mononobe-Okabe)")
    s1, s2, s3, s4 = st.columns(4)
    s1.metric("Governing SF", f"{res['SF_gov_seis']:.2f}", "OK" if res["SF_gov_seis"] >= 1.15 else "CHECK")
    s2.metric("SF Guling (≥ 1.5)", f"{res['SF_ot_seis']:.2f}", "OK" if res["SF_ot_seis"] >= 1.5 else "NOT OK")
    s3.metric("SF Geser (≥ 1.15)", f"{res['SF_sl_seis']:.2f}", "OK" if res["SF_sl_seis"] >= 1.15 else "NOT OK")
    s4.metric("SF Bearing (≥ 2.0)", f"{res['SF_bc_seis']:.2f}", "OK" if res["SF_bc_seis"] >= 2.0 else "NOT OK")

# Tabel Detail Stabilitas
df_sf = pd.DataFrame([
    {
        "Kriteria Kontrol": "1. Stabilitas Guling (Overturning)",
        "Nilai Statis": f"{res['SF_ot_stat']:.2f}",
        "Syarat Statis": "≥ 2.00",
        "Status Statis": "OK" if res["SF_ot_stat"] >= 2.0 else "NOT OK",
        "Nilai Gempa": f"{res['SF_ot_seis']:.2f}",
        "Syarat Gempa": "≥ 1.50",
        "Status Gempa": "OK" if res["SF_ot_seis"] >= 1.5 else "NOT OK",
    },
    {
        "Kriteria Kontrol": "2. Stabilitas Geser (Sliding)",
        "Nilai Statis": f"{res['SF_sl_stat']:.2f}",
        "Syarat Statis": "≥ 1.50",
        "Status Statis": "OK" if res["SF_sl_stat"] >= 1.5 else "NOT OK",
        "Nilai Gempa": f"{res['SF_sl_seis']:.2f}",
        "Syarat Gempa": "≥ 1.15",
        "Status Gempa": "OK" if res["SF_sl_seis"] >= 1.15 else "NOT OK",
    },
    {
        "Kriteria Kontrol": "3. Daya Dukung Tanah (Bearing)",
        "Nilai Statis": f"{res['SF_bc_stat']:.2f} (q_max={res['q_max_stat']:.1f} kPa)",
        "Syarat Statis": f"≥ {SF_bc_target:.2f} (q_all={q_all:.0f} kPa)",
        "Status Statis": "OK" if res["SF_bc_stat"] >= SF_bc_target else "NOT OK",
        "Nilai Gempa": f"{res['SF_bc_seis']:.2f} (q_max={res['q_max_seis']:.1f} kPa)",
        "Syarat Gempa": "≥ 2.00",
        "Status Gempa": "OK" if res["SF_bc_seis"] >= 2.0 else "NOT OK",
    },
    {
        "Kriteria Kontrol": "4. Eksentrisitas Dasar (Kern B/6)",
        "Nilai Statis": f"e = {res['e_stat']:.3f} m",
        "Syarat Statis": f"≤ {res['B']/6.0:.3f} m",
        "Status Statis": "OK" if res["e_stat"] <= res["B"] / 6.0 else "NOT OK (Uplift)",
        "Nilai Gempa": f"e = {res['e_seis']:.3f} m",
        "Syarat Gempa": f"≤ {res['B']/4.0:.3f} m",
        "Status Gempa": "OK" if res["e_seis"] <= res["B"] / 4.0 else "NOT OK",
    }
])
st.dataframe(df_sf, use_container_width=True, hide_index=True)

st.divider()

col_left, col_right = st.columns([1.15, 1.0])
with col_left:
    st.subheader("2. Visualisasi Penampang & Diagram Gaya")
    st.pyplot(fig_cross)

with col_right:
    st.subheader("3. Rincian Gaya Vertikal & Momen Penahan")
    st.dataframe(res["df_v"].style.format({"Gaya_V (kN/m)": "{:.2f}", "Lengan_x (m)": "{:.3f}", "Momen_Penahan (kNm/m)": "{:.2f}"}),
                 use_container_width=True, hide_index=True)
    st.info(
        f"**Koefisien Tanah:** Ka = {res['Ka']:.3f} | Kp = {res['Kp']:.3f} | Kae (Gempa) = {res['Kae']:.3f}\n\n"
        f"**Gaya Dorong Statis:** Pa = {res['Pa']:.2f} kN/m | Pq = {res['Pq']:.2f} kN/m | Pw (Angin) = {res['Pw']:.2f} kN/m\n\n"
        f"**Tahanan Geser:** μ_eff = {res['mu_eff']:.3f} | Pasif Depan = {res['Pp_front']:.2f} kN/m | Pasif Shear Key = {res['Pp_key']:.2f} kN/m"
    )

st.divider()

# ==============================================================================
# TAMPILAN PENULANGAN BETON BERTULANG (SNI 2847:2019)
# ==============================================================================
st.subheader("4. Desain Penulangan Dinding & Pelat Bawah (SNI 2847:2019)")

rebar_rows = []
def add_rebar_row(elem_name, pos_name, rb, dia, sp):
    if rb is None:
        rebar_rows.append({
            "Elemen Struktur": elem_name, "Sisi Tarik": pos_name, "Tebal h (mm)": "-",
            "Mu (kNm/m)": "-", "As Perlu (mm²)": "-", "Tulangan Terpasang": "Tidak Ada (B_toe = 0)",
            "As Pasang (mm²)": "-", "D/C Lentur": "-", "Vu (kN/m)": "-", "φVc (kN/m)": "-",
            "D/C Geser": "-", "Status": "N/A"
        })
    else:
        rebar_rows.append({
            "Elemen Struktur": elem_name,
            "Sisi Tarik": pos_name,
            "Tebal h (mm)": f"{rb['h_mm']:.0f}",
            "Mu (kNm/m)": f"{rb['Mu']:.2f}",
            "As Perlu (mm²)": f"{rb['As_req']:.0f}",
            "Tulangan Terpasang": f"D{dia}-{sp}",
            "As Pasang (mm²)": f"{rb['As_prov']:.0f}",
            "D/C Lentur": f"{rb['dc_flex']:.2f}",
            "Vu (kN/m)": f"{rb['Vu']:.2f}",
            "φVc (kN/m)": f"{rb['phi_Vc']:.2f}",
            "D/C Geser": f"{rb['dc_shear']:.2f}",
            "Status": rb["status"]
        })

add_rebar_row("Dinding Vertikal (Stem)", "Sisi Dalam (Backfill)", res["rebar_stem"], D_wall, s_wall)
add_rebar_row("Slab Belakang (Heel)", "Serat Atas (Top)", res["rebar_heel"], D_heel, s_heel)
add_rebar_row("Slab Depan (Toe)", "Serat Bawah (Bottom)", res["rebar_toe"], D_toe, s_toe)
if res["rebar_pocket"] is not None:
    add_rebar_row("Dinding Pocket Pagar (500mm)", "Lokal Dinding Pocket", res["rebar_pocket"], 10, 150)

df_rebar = pd.DataFrame(rebar_rows)
st.dataframe(df_rebar, use_container_width=True, hide_index=True)
st.caption("Catatan: Tulangan susut & suhu arah memanjang (distribusi) dipasang minimal ρ = 0.0018 × b × h (SNI 2847:2019 Pasal 24.4.3.2) pada kedua sisi dinding dan pelat.")

# ==============================================================================
# EXPORT LAPORAN WORD (.DOCX) & EXCEL (.XLSX) BERUMUS DINAMIS
# ==============================================================================
st.divider()
st.subheader("5. Export Hasil Perhitungan untuk Laporan")

def generate_word_report(res, fig):
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(10.5)

    # Header Judul
    title = doc.add_heading("LAPORAN PERHITUNGAN GEOTEKNIK & STRUKTUR\nDINDING PENAHAN TANAH (DPT)", level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p_meta = doc.add_paragraph()
    p_meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_meta.add_run(f"Proyek: {project_name}\n").bold = True
    p_meta.add_run(f"Tipe Struktur: {wall_type_name} | Perencana: {engineer_name} | Tanggal: {datetime.now().strftime('%d %B %Y')}")

    doc.add_heading("1. Kriteria Desain & Referensi Standar", level=2)
    doc.add_paragraph(
        "Analisis stabilitas eksternal dan desain struktur beton bertulang Dinding Penahan Tanah mengacu pada:\n"
        "• SNI 8460:2017 — Persyaratan Perancangan Geoteknik (Stabilitas guling, geser, daya dukung, dan gempa Mononobe-Okabe).\n"
        "• SNI 2847:2019 — Persyaratan Beton Struktural untuk Bangunan Gedung (Desain lentur, geser satu arah, dan tulangan minimum).\n"
        "• SNI 1727:2020 — Beban Desain Minimum dan Kriteria Terkait untuk Bangunan Gedung dan Struktur Lain."
    )

    doc.add_heading("2. Parameter Input Geometri, Tanah & Material", level=2)
    t_in = doc.add_table(rows=1, cols=4)
    t_in.style = "Table Grid"
    hdr = t_in.rows[0].cells
    hdr[0].text = "Parameter Geometri & Beban"
    hdr[1].text = "Nilai"
    hdr[2].text = "Parameter Tanah & Material"
    hdr[3].text = "Nilai"

    in_pairs = [
        ("Tinggi Ekspos (H_exp)", f"{H_exp:.2f} m", "Tanah Urug (γ1 / φ1 / c1)", f"{gamma_1} kN/m³ / {phi_1}° / {c_1} kPa"),
        ("Kedalaman Tertanam (Df)", f"{D_f:.2f} m", "Tanah Dasar (γ2 / φ2 / c2)", f"{gamma_2} kN/m³ / {phi_2}° / {c_2} kPa"),
        ("Tebal Stem (Atas / Bawah)", f"{t_top:.2f} m / {t_bot:.2f} m", "Daya Dukung Izin (q_all)", f"{q_all:.1f} kPa (Target SF = {SF_bc_target:.1f})"),
        ("Panjang Toe / Heel", f"{B_toe:.2f} m / {B_heel:.2f} m", "Koef. Gesek Dasar (μ)", f"{mu_base:.2f} (Efektif = {res['mu_eff']:.3f})"),
        ("Lebar Pondasi Total (B)", f"{res['B']:.2f} m", "Koef. Gempa Horizontal (kh)", f"{k_h:.2f}"),
        ("Tebal Slab (Pangkal / Ujung)", f"{t_slab_base:.2f} m / {t_slab_tip:.2f} m", "Mutu Beton (fc') / Baja (fy)", f"{fc:.1f} MPa / {fy:.0f} MPa"),
        ("Shear Key (D_key × B_key)", f"{D_key:.2f} m × {B_key:.2f} m" if use_shear_key else "Tidak Digunakan", "Selimut Beton (Dinding / Slab)", f"{cc_wall} mm / {cc_slab} mm"),
        ("Pagar Atas & Beban Surcharge", f"H={H_panel+H_wire:.2f}m, q={q_surch:.1f} kPa", "Koef. Tekanan (Ka / Kp / Kae)", f"{res['Ka']:.3f} / {res['Kp']:.3f} / {res['Kae']:.3f}")
    ]
    for r_data in in_pairs:
        row = t_in.add_row().cells
        for idx, val in enumerate(r_data):
            row[idx].text = val

    doc.add_heading("3. Sketsa Penampang & Distribusi Tekanan", level=2)
    img_buf = io.BytesIO()
    fig.savefig(img_buf, format="png", dpi=180, bbox_inches="tight")
    img_buf.seek(0)
    doc.add_picture(img_buf, width=Inches(6.0))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_heading("4. Ringkasan Safety Factor (Stabilitas Eksternal)", level=2)
    t_sf = doc.add_table(rows=1, cols=7)
    t_sf.style = "Table Grid"
    for idx, col_name in enumerate(df_sf.columns):
        t_sf.rows[0].cells[idx].text = col_name
    for _, r in df_sf.iterrows():
        row = t_sf.add_row().cells
        for idx, col_name in enumerate(df_sf.columns):
            row[idx].text = str(r[col_name])

    doc.add_paragraph(
        f"\nGoverning Safety Factor Statis = {res['SF_gov_stat']:.2f} | "
        f"Governing Safety Factor Gempa = {res['SF_gov_seis']:.2f}."
    ).bold = True

    doc.add_heading("5. Rincian Gaya Vertikal & Momen Penahan", level=2)
    t_v = doc.add_table(rows=1, cols=4)
    t_v.style = "Table Grid"
    for idx, col_name in enumerate(res["df_v"].columns):
        t_v.rows[0].cells[idx].text = col_name
    for _, r in res["df_v"].iterrows():
        row = t_v.add_row().cells
        row[0].text = str(r["Komponen"])
        row[1].text = f"{r['Gaya_V (kN/m)']:.2f}"
        row[2].text = f"{r['Lengan_x (m)']:.3f}"
        row[3].text = f"{r['Momen_Penahan (kNm/m)']:.2f}"

    doc.add_heading("6. Rekapitulasi Desain Penulangan Beton Bertulang (SNI 2847:2019)", level=2)
    t_rb = doc.add_table(rows=1, cols=9)
    t_rb.style = "Table Grid"
    rb_cols = ["Elemen Struktur", "Sisi Tarik", "h (mm)", "Mu (kNm)", "As Perlu", "Tulangan", "D/C Lentur", "D/C Geser", "Status"]
    for idx, c_name in enumerate(rb_cols):
        t_rb.rows[0].cells[idx].text = c_name
    for _, r in df_rebar.iterrows():
        row = t_rb.add_row().cells
        row[0].text = str(r["Elemen Struktur"])
        row[1].text = str(r["Sisi Tarik"])
        row[2].text = str(r["Tebal h (mm)"])
        row[3].text = str(r["Mu (kNm/m)"])
        row[4].text = str(r["As Perlu (mm²)"])
        row[5].text = str(r["Tulangan Terpasang"])
        row[6].text = str(r["D/C Lentur"])
        row[7].text = str(r["D/C Geser"])
        row[8].text = str(r["Status"])

    out_docx = io.BytesIO()
    doc.save(out_docx)
    out_docx.seek(0)
    return out_docx

def generate_excel_report(res, fig):
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "1_Input_&_Summary"

    # Styling
    font_title = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    fill_header = PatternFill("solid", fgColor="1E3A8A")
    fill_sub = PatternFill("solid", fgColor="DBEAFE")
    fill_input = PatternFill("solid", fgColor="FEF9C3")
    font_bold = Font(name="Calibri", size=10, bold=True)
    thin_side = Side(style="thin", color="9CA3AF")
    border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    ws1.merge_cells("A1:E1")
    ws1["A1"] = f"PERHITUNGAN DINDING PENAHAN TANAH (DPT) — {wall_type_name}"
    ws1["A1"].font = font_title
    ws1["A1"].fill = fill_header
    ws1["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 28

    ws1["A3"] = "Nama Proyek:"
    ws1["B3"] = project_name
    ws1["A4"] = "Perencana:"
    ws1["B4"] = engineer_name

    # Input Table (Rows 6 to 34)
    ws1["A6"] = "PARAMETER INPUT"
    ws1["B6"] = "SIMBOL"
    ws1["C6"] = "NILAI"
    ws1["D6"] = "SATUAN"
    for col_let in ["A", "B", "C", "D"]:
        ws1[f"{col_let}6"].font = font_bold
        ws1[f"{col_let}6"].fill = fill_sub
        ws1[f"{col_let}6"].border = border_all

    inputs_list = [
        ("Tinggi Ekspos Tanah Depan", "H_exp", H_exp, "m"),                 # Row 7
        ("Kedalaman Tertanam Depan", "Df", D_f, "m"),                       # Row 8
        ("Tebal Puncak Stem", "t_top", t_top, "m"),                         # Row 9
        ("Tebal Dasar Stem", "t_bot", t_bot, "m"),                          # Row 10
        ("Panjang Toe (0 = Tipe L)", "B_toe", B_toe, "m"),                  # Row 11
        ("Panjang Heel", "B_heel", B_heel, "m"),                            # Row 12
        ("Tebal Slab Pangkal", "t_slab_base", t_slab_base, "m"),            # Row 13
        ("Tebal Slab Ujung", "t_slab_tip", t_slab_tip, "m"),                # Row 14
        ("Kedalaman Shear Key", "D_key", D_key, "m"),                       # Row 15
        ("Lebar Shear Key", "B_key", B_key, "m"),                           # Row 16
        ("Lebar Pondasi Total (Auto)", "B", "=C11+C10+C12", "m"),           # Row 17
        ("Tinggi Total DPT (Auto)", "H_tot", "=C7+C8+C13", "m"),            # Row 18
        ("Tinggi Total Pagar Atas", "H_fence", H_panel + H_wire, "m"),      # Row 19
        ("Tekanan Angin Pagar", "q_wind", q_wind, "kPa"),                   # Row 20
        ("Beban Surcharge Jalan", "q_surch", q_surch, "kPa"),               # Row 21
        ("Berat Isi Tanah Urug", "gamma_1", gamma_1, "kN/m3"),              # Row 22
        ("Sudut Geser Dalam Urug", "phi_1", phi_1, "deg"),                  # Row 23
        ("Berat Isi Tanah Dasar", "gamma_2", gamma_2, "kN/m3"),             # Row 24
        ("Sudut Geser Dalam Dasar", "phi_2", phi_2, "deg"),                # Row 25
        ("Kohesi Tanah Dasar", "c_2", c_2, "kPa"),                          # Row 26
        ("Daya Dukung Izin Tanah", "q_all", q_all, "kPa"),                  # Row 27
        ("SF Daya Dukung Target", "SF_bc", SF_bc_target, "-"),              # Row 28
        ("Koefisien Gesek Dasar", "mu", mu_base, "-"),                      # Row 29
        ("Faktor Reduksi Pasif Depan", "R_pasif", R_pasif, "-"),            # Row 30
        ("Koefisien Gempa Horizontal", "k_h", k_h, "-"),                    # Row 31
        ("Mutu Beton", "fc'", fc, "MPa"),                                   # Row 32
        ("Mutu Baja Tulangan", "fy", fy, "MPa"),                            # Row 33
        ("Selimut Beton Dinding", "cc_wall", cc_wall, "mm"),                # Row 34
        ("Selimut Beton Slab", "cc_slab", cc_slab, "mm"),                   # Row 35
        ("Diameter Tulangan Dinding", "D_wall", D_wall, "mm"),              # Row 36
        ("Spasi Tulangan Dinding", "s_wall", s_wall, "mm"),                 # Row 37
        ("Diameter Tulangan Heel", "D_heel", D_heel, "mm"),                 # Row 38
        ("Spasi Tulangan Heel", "s_heel", s_heel, "mm"),                    # Row 39
        ("Diameter Tulangan Toe", "D_toe", D_toe, "mm"),                    # Row 40
        ("Spasi Tulangan Toe", "s_toe", s_toe, "mm"),                       # Row 41
    ]

    for idx, (label, sym, val, unit) in enumerate(inputs_list, start=7):
        ws1[f"A{idx}"] = label
        ws1[f"B{idx}"] = sym
        ws1[f"C{idx}"] = val
        ws1[f"D{idx}"] = unit
        for col_let in ["A", "B", "C", "D"]:
            ws1[f"{col_let}{idx}"].border = border_all
        if not (isinstance(val, str) and val.startswith("=")):
            ws1[f"C{idx}"].fill = fill_input

    # Summary Table di Kolom F-J
    ws1["F6"] = "KONTROL STABILITAS (SNI 8460:2017)"
    ws1["G6"] = "STATIS"
    ws1["H6"] = "SYARAT"
    ws1["I6"] = "STATUS"
    for col_let in ["F", "G", "H", "I"]:
        ws1[f"{col_let}6"].font = font_bold
        ws1[f"{col_let}6"].fill = fill_sub
        ws1[f"{col_let}6"].border = border_all

    sf_excel_rows = [
        ("SF Guling (Overturning)", "='2_Stabilitas_&_Penulangan'!C22", 2.0, '=IF(G7>=H7,"OK","NOT OK")'),
        ("SF Geser (Sliding)", "='2_Stabilitas_&_Penulangan'!C23", 1.5, '=IF(G8>=H8,"OK","NOT OK")'),
        ("SF Daya Dukung (Bearing)", "='2_Stabilitas_&_Penulangan'!C27", "=C28", '=IF(G9>=H9,"OK","NOT OK")'),
        ("Eksentrisitas Dasar e (m)", "='2_Stabilitas_&_Penulangan'!C25", "=C17/6", '=IF(G10<=H10,"OK","NOT OK")'),
        ("SF Guling Gempa", "='2_Stabilitas_&_Penulangan'!C31", 1.5, '=IF(G11>=H11,"OK","NOT OK")'),
        ("SF Geser Gempa", "='2_Stabilitas_&_Penulangan'!C32", 1.15, '=IF(G12>=H12,"OK","NOT OK")'),
    ]
    for idx, (lbl, f_val, f_req, f_stat) in enumerate(sf_excel_rows, start=7):
        ws1[f"F{idx}"] = lbl
        ws1[f"G{idx}"] = f_val
        ws1[f"H{idx}"] = f_req
        ws1[f"I{idx}"] = f_stat
        ws1[f"G{idx}"].number_format = "0.00"
        ws1[f"H{idx}"].number_format = "0.00"
        for col_let in ["F", "G", "H", "I"]:
            ws1[f"{col_let}{idx}"].border = border_all

    # Embed Gambar ke Excel
    img_buf_xl = io.BytesIO()
    fig.savefig(img_buf_xl, format="png", dpi=130, bbox_inches="tight")
    img_buf_xl.seek(0)
    xl_img = XLImage(img_buf_xl)
    xl_img.width = 480
    xl_img.height = 350
    ws1.add_image(xl_img, "F15")

    # Sheet 2: Rumus Dinamis Stabilitas & Penulangan
    ws2 = wb.create_sheet("2_Stabilitas_&_Penulangan")
    ws2["A1"] = "PERHITUNGAN DETAIL STABILITAS & PENULANGAN (RUMUS DINAMIS)"
    ws2["A1"].font = Font(name="Calibri", size=12, bold=True)

    calc_formulas = [
        ("A. KOEFISIEN TEKANAN TANAH", "", ""),
        ("Koefisien Tekanan Aktif (Ka)", "=TAN(RADIANS(45-'1_Input_&_Summary'!C23/2))^2", "-"),             # Row 4
        ("Koefisien Tekanan Pasif (Kp)", "=TAN(RADIANS(45+'1_Input_&_Summary'!C25/2))^2", "-"),              # Row 5
        ("B. GAYA VERTIKAL & MOMEN PENAHAN", "", ""),
        ("Berat Sendiri Beton DPT Total (W_conc)", "=(('1_Input_&_Summary'!C9+'1_Input_&_Summary'!C10)/2*('1_Input_&_Summary'!C7+'1_Input_&_Summary'!C8) + '1_Input_&_Summary'!C17*'1_Input_&_Summary'!C14 + 0.5*'1_Input_&_Summary'!C12*('1_Input_&_Summary'!C13-'1_Input_&_Summary'!C14) + '1_Input_&_Summary'!C15*'1_Input_&_Summary'!C16)*24", "kN/m"), # Row 7
        ("Berat Tanah Urug di atas Heel (W_soil)", "=('1_Input_&_Summary'!C12*('1_Input_&_Summary'!C7+'1_Input_&_Summary'!C8) + 0.5*'1_Input_&_Summary'!C12*('1_Input_&_Summary'!C13-'1_Input_&_Summary'!C14))*'1_Input_&_Summary'!C22", "kN/m"), # Row 8
        ("Berat Tanah di atas Toe (W_toe_soil)", "='1_Input_&_Summary'!C11*'1_Input_&_Summary'!C8*'1_Input_&_Summary'!C24", "kN/m"), # Row 9
        ("Total Gaya Vertikal Tanpa Surcharge (ΣV_no_q)", "=C7+C8+C9", "kN/m"),                               # Row 10
        ("Total Gaya Vertikal Dengan Surcharge (ΣV_tot)", "=C10+'1_Input_&_Summary'!C21*'1_Input_&_Summary'!C12", "kN/m"), # Row 11
        ("Total Momen Penahan Tanpa Surcharge (ΣMR_no_q)", f"={res['sum_MR_no_q']:.2f}", "kNm/m"),           # Row 12
        ("Total Momen Penahan Dengan Surcharge (ΣMR_tot)", "=C12+('1_Input_&_Summary'!C21*'1_Input_&_Summary'!C12)*('1_Input_&_Summary'!C11+'1_Input_&_Summary'!C10+'1_Input_&_Summary'!C12/2)", "kNm/m"), # Row 13
        ("C. GAYA HORIZONTAL & STABILITAS STATIS", "", ""),
        ("Tekanan Aktif Tanah (Pa)", "=0.5*C4*'1_Input_&_Summary'!C22*('1_Input_&_Summary'!C18^2)", "kN/m"),  # Row 15
        ("Tekanan Lateral Surcharge (Pq)", "=C4*'1_Input_&_Summary'!C21*'1_Input_&_Summary'!C18", "kN/m"),     # Row 16
        ("Gaya Angin Pagar Atas (Pw)", "='1_Input_&_Summary'!C20*'1_Input_&_Summary'!C19", "kN/m"),           # Row 17
        ("Total Gaya Dorong Statis (ΣFh_stat)", "=C15+C16+C17", "kN/m"),                                      # Row 18
        ("Total Momen Guling Statis (ΣMo_stat)", "=C15*('1_Input_&_Summary'!C18/3)+C16*('1_Input_&_Summary'!C18/2)+C17*('1_Input_&_Summary'!C18+'1_Input_&_Summary'!C19/2)", "kNm/m"), # Row 19
        ("Tahanan Pasif Depan + Shear Key (Pp_tot)", "='1_Input_&_Summary'!C30*0.5*C5*'1_Input_&_Summary'!C24*('1_Input_&_Summary'!C8^2) + IF('1_Input_&_Summary'!C15>0, 0.5*C5*'1_Input_&_Summary'!C24*(('1_Input_&_Summary'!C8+'1_Input_&_Summary'!C15)^2-'1_Input_&_Summary'!C8^2)+2*'1_Input_&_Summary'!C26*SQRT(C5)*'1_Input_&_Summary'!C15, 0)", "kN/m"), # Row 20
        ("Total Gaya Tahan Geser (ΣFr)", "=IF('1_Input_&_Summary'!C15>0,TAN(RADIANS('1_Input_&_Summary'!C25)),'1_Input_&_Summary'!C29)*C10 + 0.5*'1_Input_&_Summary'!C26*'1_Input_&_Summary'!C17 + C20", "kN/m"), # Row 21
        ("SF Guling Statis", "=C12/C19", "-"),                                                                # Row 22
        ("SF Geser Statis", "=C21/C18", "-"),                                                                 # Row 23
        ("Jarak Resultan dari Ujung Depan (xR)", "=(C13-C19)/C11", "m"),                                      # Row 24
        ("Eksentrisitas Dasar (e)", "=('1_Input_&_Summary'!C17/2)-C24", "m"),                                 # Row 25
        ("Tegangan Tanah Maksimum (q_max)", "=IF(C25<='1_Input_&_Summary'!C17/6, (C11/'1_Input_&_Summary'!C17)*(1+6*C25/'1_Input_&_Summary'!C17), (2*C11)/(3*C24))", "kPa"), # Row 26
        ("SF Daya Dukung Statis", "=('1_Input_&_Summary'!C27*'1_Input_&_Summary'!C28)/C26", "-"),             # Row 27
        ("D. STABILITAS GEMPA (MONONOBE-OKABE)", "", ""),
        ("Total Gaya Dorong Gempa (ΣFh_seis)", f"={res['Fh_seis']:.2f}", "kN/m"),                             # Row 29
        ("Total Momen Guling Gempa (ΣMo_seis)", f"={res['Mo_seis']:.2f}", "kNm/m"),                           # Row 30
        ("SF Guling Gempa", "=C12/C30", "-"),                                                                 # Row 31
        ("SF Geser Gempa", "=C21/C29", "-"),                                                                 # Row 32
        ("E. KONTROL PENULANGAN LENTUR & GESER (SNI 2847:2019)", "", ""),
        ("Dinding Stem - Momen Terfaktor Mu", f"={res['rebar_stem']['Mu']:.2f}", "kNm/m"),                    # Row 34
        ("Dinding Stem - As Terpasang", "=(1000/'1_Input_&_Summary'!C37)*(0.25*PI()*('1_Input_&_Summary'!C36^2))", "mm2/m"), # Row 35
        ("Dinding Stem - Kapasitas Lentur φMn", "=0.9*C35*'1_Input_&_Summary'!C33*(('1_Input_&_Summary'!C10*1000-'1_Input_&_Summary'!C34-0.5*'1_Input_&_Summary'!C36) - (C35*'1_Input_&_Summary'!C33)/(1.7*'1_Input_&_Summary'!C32*1000))/1000000", "kNm/m"), # Row 36
        ("Dinding Stem - Rasio D/C Lentur", "=C34/C36", "-"),                                                 # Row 37
        ("Heel Slab - Momen Terfaktor Mu", f"={res['rebar_heel']['Mu']:.2f}", "kNm/m"),                       # Row 38
        ("Heel Slab - As Terpasang", "=(1000/'1_Input_&_Summary'!C39)*(0.25*PI()*('1_Input_&_Summary'!C38^2))", "mm2/m"), # Row 39
        ("Heel Slab - Kapasitas Lentur φMn", "=0.9*C39*'1_Input_&_Summary'!C33*(('1_Input_&_Summary'!C13*1000-'1_Input_&_Summary'!C35-0.5*'1_Input_&_Summary'!C38) - (C39*'1_Input_&_Summary'!C33)/(1.7*'1_Input_&_Summary'!C32*1000))/1000000", "kNm/m"), # Row 40
        ("Heel Slab - Rasio D/C Lentur", "=C38/C40", "-"),                                                    # Row 41
    ]

    for idx, (desc, form, unit) in enumerate(calc_formulas, start=3):
        ws2[f"A{idx}"] = desc
        ws2[f"C{idx}"] = form
        ws2[f"D{idx}"] = unit
        if form == "":
            ws2[f"A{idx}"].font = font_bold
            ws2[f"A{idx}"].fill = fill_sub
        else:
            ws2[f"C{idx}"].number_format = "0.00"
            for c_l in ["A", "C", "D"]:
                ws2[f"{c_l}{idx}"].border = border_all

    for ws in [ws1, ws2]:
        for col in ws.columns:
            max_len = 0
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            for cell in col:
                if cell.value and not str(cell.value).startswith("="):
                    max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = max(14, min(max_len + 4, 46))

    out_xlsx = io.BytesIO()
    wb.save(out_xlsx)
    out_xlsx.seek(0)
    return out_xlsx

col_dl1, col_dl2 = st.columns(2)
with col_dl1:
    docx_data = generate_word_report(res, fig_cross)
    st.download_button(
        label="📄 Download Laporan Perhitungan (Word .docx)",
        data=docx_data,
        file_name=f"Laporan_DPT_{wall_type_name.split()[0]}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )

with col_dl2:
    xlsx_data = generate_excel_report(res, fig_cross)
    st.download_button(
        label="📊 Download Calculation Sheet (Excel .xlsx)",
        data=xlsx_data,
        file_name=f"Perhitungan_DPT_{wall_type_name.split()[0]}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )