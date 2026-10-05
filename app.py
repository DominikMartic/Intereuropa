import io
import os
import re
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Kontrola Intereuropa Računa", page_icon="📦", layout="wide"
)

st.title("📦 Sustav za Kontrolu i Analizu Računa - Intereuropa")
st.write(
    "Učitaj Excel specifikaciju računa Intereurope. Sustav automatski"
    " prepoznaje zone prema poštanskim brojevima, kontrolira rokove tranzita i"
    " omogućuje pregled stavki."
)


# Određivanje zone prema poštanskom broju primatelja (Prilog 1 cjeniku Intereuropa)
def odredi_intereuropa_zonu(pbr_primatelja):
    try:
        pbr = int(pbr_primatelja)
    except:
        return "Zona 2"

    island_pbr = [
        20230, 20240, 20242, 20243, 20244, 20245, 20246, 20247, 20248, 20250,
        20260, 20263, 20264, 20267, 20269, 20270, 20271, 20272, 20273, 20274,
        20275, 21400, 21403, 21404, 21405, 21410, 21412, 21413, 21420, 21423,
        21424, 21425, 21426, 21450, 21454, 21460, 21462, 21463, 21465, 21466,
        21467, 21468, 21469, 21480, 21483, 21485, 22240, 22242, 22243, 22244,
        23212, 23234, 23249, 23250, 23251, 23262, 23263, 23264, 23271, 23272,
        23273, 23274, 23275, 51280, 51281, 51500, 51511, 51512, 51513, 51514,
        51515, 51516, 51517, 51521, 51522, 51523, 51550, 51551, 51554, 51555,
        51556, 51557, 51559, 51564, 53291, 53294, 53296,
    ]
    if pbr in island_pbr:
        return "Zona 3 (Otoci)"

    prefix = pbr // 1000
    if prefix == 53:
        return "Zona 3"
    elif prefix in [10, 44, 47, 49]:
        return "Zona 1"
    else:
        return "Zona 2"


def izracunaj_radne_dane(datum_slanja, datum_dostave):
    try:
        d1 = pd.to_datetime(datum_slanja, errors="coerce")
        d2 = pd.to_datetime(datum_dostave, errors="coerce")
        if pd.isna(d1) or pd.isna(d2):
            return None
        b_days = (
            pd.bdate_range(start=d1.normalize(), end=d2.normalize()).shape[0] - 1
        )
        return max(0, b_days)
    except:
        return None


def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Izvjestaj")
    return output.getvalue()


uploaded_file = st.file_uploader(
    "Učitaj Excel specifikaciju računa Intereurope", type=["xlsx", "csv"]
)

if uploaded_file is not None:
    if uploaded_file.name.endswith(".xlsx"):
        df = pd.read_excel(uploaded_file, sheet_name=0, header=1)
    else:
        df = pd.read_csv(uploaded_file)

    st.success("Specifikacija uspješno učitana!")

    if st.button("Pokreni kontrolu i analizu"):
        rezultati = []

        for idx, row in df.iterrows():
            usluga = row.get("Usluga", 0)
            usluga_naziv = row.get("Usluga-naziv", "")
            pbr_primatelja = row.get("Prim.-pošt.br.", 10000)
            masa = float(row.get("B.teža", row.get("Vred.osn./količ.", 0.0)))
            naplaceni_iznos = float(row.get("Iznos (bezPDV)", 0.0))

            d_slanja = row.get("Odlazak", None)
            d_dostave = row.get("Dostava", None)
            tranzit_dani = izracunaj_radne_dane(d_slanja, d_dostave)

            zona = odredi_intereuropa_zonu(pbr_primatelja)

            red = {
                "RedniBroj": idx + 1,
                "Broj Računa": row.get("Br.rač.", ""),
                "Narudžba": row.get("Nar.", ""),
                "Usluga": usluga_naziv,
                "Primatelj": row.get("Primatelj", ""),
                "Mjesto": row.get("Prim-mjesto", ""),
                "Poštanski broj": pbr_primatelja,
                "Zona": zona,
                "Masa (kg)": masa,
                "Slanje": d_slanja,
                "Dostava": d_dostave,
                "Tranzit (dana)": (
                    tranzit_dani if tranzit_dani is not None else -1
                ),
                "Iznos (€ bez PDV)": round(naplaceni_iznos, 2),
            }
            rezultati.append(red)

        res_df = pd.DataFrame(rezultati)

        tab1, tab2, tab3 = st.tabs([
            "📊 1. Pregled svih stavki računa",
            "⛽ 2. Pregled dodataka za gorivo i usluga",
            "⏱️ 3. Analiza tranzita pošiljaka",
        ])

        with tab1:
            st.subheader("Sve stavke specifikacije")
            st.dataframe(res_df, use_container_width=True)
            st.download_button(
                "📥 Preuzmi Excel (Sve stavke)",
                to_excel(res_df),
                "intereuropa_sve_stavke.xlsx",
                (
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
            )

        with tab2:
            st.subheader("Pregled goriva i dodatnih usluga")
            dodatne = res_df[~res_df["Usluga"].str.contains("EXPRESS", case=False, na=False)]
            st.dataframe(dodatne, use_container_width=True)
            st.download_button(
                "📥 Preuzmi Excel (Dodatne usluge)",
                to_excel(dodatne),
                "intereuropa_dodatne_usluge.xlsx",
                (
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
            )

        with tab3:
            st.subheader("Analiza radnih dana tranzita")
            express_tranzit = res_df[res_df["Usluga"].str.contains("EXPRESS", case=False, na=False)]
            st.dataframe(
                express_tranzit[[
                    "Narudžba",
                    "Primatelj",
                    "Mjesto",
                    "Zona",
                    "Slanje",
                    "Dostava",
                    "Tranzit (dana)",
                ]],
                use_container_width=True,
            )
