import streamlit as st
import requests
import io
import json
import re
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from datetime import date

st.set_page_config(page_title="Generator Polityki Rachunkowosci", page_icon="📋", layout="wide")

ENTITY_FORM_LABELS = ["Sp. z o.o.", "Spolka akcyjna", "Spolka cywilna", "Spolka jawna",
                       "Spolka komandytowa", "Spolka kom.-akcyjna", "JDG", "Fundacja", "Stowarzyszenie"]
ENTITY_FORM_KEYS = ["sp_zoo", "sa", "sc", "sj", "sk", "ska", "jdg", "fundacja", "stowarzyszenie"]
ENTITY_FORM_FULL = {
    "sp_zoo": "Spolka z ograniczona odpowiedzialnoscia",
    "sa": "Spolka akcyjna", "sc": "Spolka cywilna", "sj": "Spolka jawna",
    "sk": "Spolka komandytowa", "ska": "Spolka komandytowo-akcyjna",
    "jdg": "Jednoosobowa dzialalnosc gospodarcza",
    "fundacja": "Fundacja", "stowarzyszenie": "Stowarzyszenie",
}
STEP_NAMES = ["Dane jednostki", "Ksiegi i Plan Kont", "Metody wyceny", "Koszty i RZiS",
              "Waluty obce", "Ochrona danych", "Polityki dodatkowe", "Podglad i eksport"]
ALL_CUR = ["EUR", "USD", "GBP", "CHF", "CZK", "SEK", "NOK", "DKK", "JPY", "CNY"]

if "step" not in st.session_state:
    st.session_state.step = 0
for k, v in dict(d_name="", d_form=0, d_nip="", d_krs="", d_regon="", d_addr="",
                  d_fys="01-01", d_fye="12-31", d_small=False, d_micro=False,
                  d_zpk="Wzorcowy plan kont", d_sn="", d_sv="", d_sp="",
                  d_dep="Metoda liniowa", d_thr=10000, d_iv="Cena nabycia", d_id="FIFO",
                  d_cm="Tylko Zespol 4 (uklad rodzajowy)", d_pl="Wariant porownawczy",
                  d_pc="Pelny koszt wytworzenia", d_oh="Klucz przychodowy",
                  d_fxs="Kurs sredni NBP", d_fxd="FIFO", d_hfx=False, d_cur=["EUR", "USD"],
                  d_dp="Elektroniczna i fizyczna", d_ay=5, d_bk="Codziennie", d_ac=True,
                  d_rp="", d_rev="Zasada memorialowa", d_ls="Wg przepisow bilansowych",
                  d_prov=True, d_dt=True, d_cf="Metoda posrednia",
                  d_adate=date.today(), d_edate=date.today(), d_ab="").items():
    if k not in st.session_state:
        st.session_state[k] = v

def G(k):
    return st.session_state.get(k, "")

# ══════════════════════════════════════════════════════
# KRS API
# ══════════════════════════════════════════════════════

def fetch_krs_by_krs_nr(krs_nr):
    krs_clean = re.sub(r"[^0-9]", "", krs_nr).zfill(10)
    headers = {"Accept": "application/json", "User-Agent": "Mozilla/5.0 (compatible; PolitikaRachunkowosci/1.0)"}
    url = f"https://api-krs.ms.gov.pl/api/krs/OdpisAktualny/{krs_clean}"
    try:
        r = requests.get(url, params={"rejestr": "P", "format": "json"}, headers=headers, timeout=20)
        if r.status_code == 200:
            return _parse_odpis(r.json(), krs_clean)
        r2 = requests.get(url, params={"rejestr": "S", "format": "json"}, headers=headers, timeout=20)
        if r2.status_code == 200:
            return _parse_odpis(r2.json(), krs_clean)
    except requests.exceptions.ConnectionError:
        raise ConnectionError("Brak polaczenia z API KRS")
    except requests.exceptions.Timeout:
        raise TimeoutError("API KRS nie odpowiada")
    except Exception as e:
        raise RuntimeError(f"Blad API KRS: {e}")
    return None

def _parse_odpis(data, krs_nr=""):
    try:
        odpis = data.get("odpis", data)
        naglowek = odpis.get("naglowekA", {})
        dane = odpis.get("dane", {})
        dzial1 = dane.get("dzial1", {})
        dane_p = dzial1.get("danePodmiotu", {})
        nazwa = dane_p.get("nazwa", "")
        ident = dane_p.get("identyfikatory", {})
        nip_val = ident.get("nip", "")
        regon_raw = ident.get("regon", "")
        regon_val = regon_raw[:9] if regon_raw else ""
        forma = dane_p.get("formaPrawna", "")
        siedz_blok = dzial1.get("siedzibaIAdres", {})
        adres = siedz_blok.get("adres", {})
        ulica = adres.get("ulica", "")
        nr_domu = adres.get("nrDomu", "")
        nr_lok = adres.get("nrLokalu", "")
        kod = adres.get("kodPocztowy", "")
        miasto = adres.get("miejscowosc", "")
        siedziba = f"{ulica} {nr_domu}".strip()
        if nr_lok: siedziba += f"/{nr_lok}"
        if kod and miasto: siedziba += f", {kod} {miasto}"
        krs_val = naglowek.get("numerKRS", krs_nr)
        fl = forma.lower() if isinstance(forma, str) else ""
        forma_key = ("sp_zoo" if "ograniczon" in fl else "ska" if "komandytowo-akcyjn" in fl else
                     "sk" if "komandytow" in fl else "sa" if "akcyjn" in fl else
                     "sj" if "jawn" in fl else "fundacja" if "fundacj" in fl else
                     "stowarzyszenie" if "stowarzysz" in fl else "")
        dzial2 = dane.get("dzial2", {})
        sklad = dzial2.get("reprezentacja", {}).get("sklad", [])
        rep = ""
        if sklad:
            o = sklad[0]
            no = o.get("nazwisko", {})
            io2 = o.get("imiona", {})
            nz = no.get("nazwiskoICzlon", "") if isinstance(no, dict) else str(no)
            im = io2.get("imie", "") if isinstance(io2, dict) else str(io2)
            fn = o.get("funkcjaWOrganie", o.get("funkcja", ""))
            rep = f"{im} {nz}".strip()
            if fn: rep += f" - {fn}"
        return {"nazwa": nazwa, "siedziba": siedziba, "nip": nip_val, "krs": krs_val,
                "regon": regon_val, "forma_key": forma_key, "forma_prawna": forma, "rep": rep}
    except Exception:
        return None


# ══════════════════════════════════════════════════════
# ZPK GENERATOR — LOGIKA PLANU KONT 2026
# ══════════════════════════════════════════════════════

def generate_zpk(branza, typ_cit, wariant_rzis, skala, obsluga_aut, podmioty_powiazane):
    """Generuje Zakladowy Plan Kont na podstawie parametrow."""
    konta = []

    def add(kod, nazwa, typ, atr_pod, ksef=""):
        konta.append({"Kod_Konta": kod, "Nazwa_Konta": nazwa, "Typ": typ,
                       "Atrybut_Podatkowy": atr_pod, "Znacznik_KSeF": ksef})

    tp = ".TP" if podmioty_powiazane else ""

    # ── ZESPOL 0: Aktywa trwale ──
    add("010", "Srodki trwale", "Bilansowe", "-")
    add("011", "Wartosci niematerialne i prawne", "Bilansowe", "-")
    add("013", "Srodki trwale w budowie", "Bilansowe", "-")
    add("020", "Wartosci niematerialne i prawne - WNiP", "Bilansowe", "-")
    add("030", "Dlugoterminowe aktywa finansowe", "Bilansowe", "-")
    add("070", "Umorzenie srodkow trwalych", "Bilansowe", "-")
    add("071", "Umorzenie WNiP", "Bilansowe", "-")
    add("080", "Srodki trwale w budowie", "Bilansowe", "-")

    # ── ZESPOL 1: Srodki pieniezne ──
    add("100", "Kasa", "Bilansowe", "-")
    add("130", "Rachunki bankowe", "Bilansowe", "-")
    add("131", "Rachunek bankowy - biezacy PLN", "Bilansowe", "-")
    add("132", "Rachunek bankowy - walutowy", "Bilansowe", "-")
    add("135", "Rachunek VAT (split payment)", "Bilansowe", "-", "VAT_SPP")
    add("139", "Srodki pieniezne w drodze", "Bilansowe", "-")
    add("140", "Krotkoterminowe aktywa finansowe", "Bilansowe", "-")

    # ── ZESPOL 2: Rozrachunki ──
    add("200", "Rozrachunki z odbiorcami", "Bilansowe", "-", "FA_NAL")
    add("201", "Rozrachunki z dostawcami", "Bilansowe", "-", "FA_ZOB")
    add("220", "Rozrachunki publicznoprawne", "Bilansowe", "-")
    add("221", "Rozrachunki z US - VAT nalezny", "Bilansowe", "-", "VAT_NAL")
    add("222", "Rozrachunki z US - VAT naliczony", "Bilansowe", "-", "VAT_NAL")
    add("223", "Rozrachunki z US - CIT", "Bilansowe", "CIT")
    add("225", "Rozrachunki z US - PIT (pracownicy)", "Bilansowe", "-")
    add("229", "Rozrachunki z ZUS", "Bilansowe", "-")
    add("230", "Rozrachunki z pracownikami - wynagrodzenia", "Bilansowe", "-")
    add("234", "Rozrachunki z pracownikami - inne", "Bilansowe", "-")
    add("240", "Pozostale rozrachunki", "Bilansowe", "-")
    add("245", "Rozrachunki z wlascicielami/wspolnikami", "Bilansowe", "-")
    add("290", "Odpisy aktualizujace naleznosci", "Bilansowe", "NKUP")

    if podmioty_powiazane:
        add("200-TP", "Rozrachunki z odbiorcami - podmioty powiazane", "Bilansowe", "-", "FA_NAL_TP")
        add("201-TP", "Rozrachunki z dostawcami - podmioty powiazane", "Bilansowe", "-", "FA_ZOB_TP")

    # ── ZESPOL 3: Materialy i towary ──
    if branza in ["Produkcja", "Hybryda"]:
        add("310", "Materialy", "Bilansowe", "-")
        add("311", "Materialy na skladzie", "Bilansowe", "-")
        add("340", "Odchylenia od cen ewidencyjnych materialow", "Bilansowe", "-")

    if branza in ["Handel", "Hybryda"]:
        add("330", "Towary", "Bilansowe", "-")
        add("340", "Odchylenia od cen ewidencyjnych towarow", "Bilansowe", "-")

    add("300", "Rozliczenie zakupu", "Bilansowe", "-")

    # ── ZESPOL 4: Koszty rodzajowe ──
    if obsluga_aut:
        add("400", "Amortyzacja", "Wynikowe", "KUP")
        add("400-01", "Amortyzacja - KUP", "Wynikowe", "KUP")
        add("400-02", "Amortyzacja - NKUP (nadwyzka ponad limit)", "Wynikowe", "NKUP")
        add("401", "Zuzycie materialow i energii", "Wynikowe", "KUP")
        add("402", "Uslugi obce", "Wynikowe", "KUP")
        add("402-01", "Uslugi obce - KUP", "Wynikowe", "KUP")
        add("402-02", "Uslugi obce - NKUP (nadwyzka limit samochod)", "Wynikowe", "NKUP")
        add("403", "Podatki i oplaty", "Wynikowe", "KUP")
        add("404", "Wynagrodzenia", "Wynikowe", "KUP")
        add("405", "Ubezpieczenia spoleczne i inne swiadczenia", "Wynikowe", "KUP")
        add("409", "Pozostale koszty rodzajowe", "Wynikowe", "KUP")
    else:
        add("400", "Amortyzacja", "Wynikowe", "KUP")
        add("401", "Zuzycie materialow i energii", "Wynikowe", "KUP")
        add("402", "Uslugi obce", "Wynikowe", "KUP")
        add("403", "Podatki i oplaty", "Wynikowe", "KUP")
        add("404", "Wynagrodzenia", "Wynikowe", "KUP")
        add("405", "Ubezpieczenia spoleczne i inne swiadczenia", "Wynikowe", "KUP")
        add("409", "Pozostale koszty rodzajowe", "Wynikowe", "KUP")

    if podmioty_powiazane:
        add(f"402{tp}", "Uslugi obce - podmioty powiazane", "Wynikowe", "KUP")
        add(f"404{tp}", "Wynagrodzenia - podmioty powiazane", "Wynikowe", "KUP")

    # ── ZESPOL 5: Koszty wg typow dzialalnosci (kalkulacyjny) ──
    if wariant_rzis == "Kalkulacyjny" or branza in ["Produkcja", "Hybryda"]:
        add("501", "Koszty produkcji podstawowej", "Wynikowe", "KUP")
        add("520", "Koszty wydzialow", "Wynikowe", "KUP")
        add("527", "Koszty sprzedazy", "Wynikowe", "KUP")
        add("550", "Koszty ogolnego zarzadu", "Wynikowe", "KUP")
        add("580", "Rozliczenie kosztow dzialalnosci", "Wynikowe", "-")

        if branza in ["Produkcja", "Hybryda"]:
            add("530", "Koszty dzialalnosci pomocniczej", "Wynikowe", "KUP")

    # ── ZESPOL 6: Produkty i rozliczenia ──
    if branza in ["Produkcja", "Hybryda"]:
        add("601", "Wyroby gotowe", "Bilansowe", "-")
        add("602", "Polprodukty i produkcja w toku", "Bilansowe", "-")
        add("620", "Odchylenia od cen ewidencyjnych produktow", "Bilansowe", "-")

    add("640", "Rozliczenia miedzyokresowe kosztow czynne", "Bilansowe", "-")
    add("641", "Rozliczenia miedzyokresowe kosztow bierne", "Bilansowe", "-")

    # ── ZESPOL 7: Przychody ──
    add(f"700{tp}", "Przychody ze sprzedazy produktow", "Wynikowe", "Przychody_Op", "FA_PRZYCH")
    add(f"701{tp}", "Przychody ze sprzedazy uslug", "Wynikowe", "Przychody_Op", "FA_PRZYCH")

    if branza in ["Produkcja", "Hybryda"]:
        add("711", "Koszt wlasny sprzedazy produktow", "Wynikowe", "KUP")

    if branza in ["Handel", "Hybryda"]:
        add(f"730{tp}", "Przychody ze sprzedazy towarow", "Wynikowe", "Przychody_Op", "FA_PRZYCH")
        add("731", "Wartosc sprzedanych towarow w cenach zakupu", "Wynikowe", "KUP")

    add("740", "Przychody ze sprzedazy materialow", "Wynikowe", "Przychody_Op")
    add("741", "Wartosc sprzedanych materialow", "Wynikowe", "KUP")
    add("760", "Pozostale przychody operacyjne", "Wynikowe", "Przychody_Op")
    add("761", "Pozostale koszty operacyjne", "Wynikowe", "KUP")
    add("750", "Przychody finansowe", "Wynikowe", "Przychody_Kap")
    add("751", "Koszty finansowe", "Wynikowe", "KUP")
    add("770", "Zyski nadzwyczajne", "Wynikowe", "Przychody_Op")
    add("771", "Straty nadzwyczajne", "Wynikowe", "KUP")
    add("790", "Obroty wewnetrzne", "Wynikowe", "-")
    add("791", "Koszt obrotow wewnetrznych", "Wynikowe", "-")

    # ── ZESPOL 8: Kapital, rezerwy, wynik ──
    add("800", "Kapital zakladowy", "Bilansowe", "-")
    add("801", "Kapital zapasowy", "Bilansowe", "-")
    add("802", "Kapital rezerwowy", "Bilansowe", "-")
    add("803", "Kapital z aktualizacji wyceny", "Bilansowe", "-")
    add("810", "Zyski/straty z lat ubieglych", "Bilansowe", "-")
    add("820", "Rozliczenie wyniku finansowego", "Bilansowe", "-")

    if typ_cit == "Estonski":
        add("821", "Ukryte zyski (CIT estonski)", "Wynikowe", "NKUP")
        add("822", "Wydatki niezwiazane z dzialalnoscia (CIT estonski)", "Wynikowe", "NKUP")
        add("823", "Dochod z tyt. wydatkow niezwiazanych z dzialalnoscia", "Wynikowe", "NKUP")
        add("824", "Dochod z tyt. zmiany wartosci skladnikow majatku", "Wynikowe", "NKUP")

    add("840", "Rezerwy i rozliczenia miedzyokresowe przychodow", "Bilansowe", "-")
    add("841", "Rezerwa z tytulu odroczonego podatku dochodowego", "Bilansowe", "-")
    add("845", "Dotacje i subwencje", "Bilansowe", "-")
    add("850", "Fundusze specjalne (ZFSS)", "Bilansowe", "-")
    add("860", "Wynik finansowy", "Wynikowe", "-")
    add("870", "Obowiazkowe obciazenia wyniku finansowego - CIT", "Wynikowe", "CIT")

    return konta


def zpk_to_xlsx(konta):
    """Konwertuje liste kont na plik XLSX."""
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    except ImportError:
        return None

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Zakladowy Plan Kont"

    # Naglowki
    headers = ["Kod_Konta", "Nazwa_Konta", "Typ", "Atrybut_Podatkowy", "Znacznik_KSeF"]
    hfill = PatternFill(start_color="1B2A4A", end_color="1B2A4A", fill_type="solid")
    hfont = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    thin = Side(style="thin", color="B0B0B0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = hfont
        cell.fill = hfill
        cell.alignment = Alignment(horizontal="center")
        cell.border = border

    # Dane
    alt_fill = PatternFill(start_color="F2F6FA", end_color="F2F6FA", fill_type="solid")
    dfont = Font(name="Calibri", size=10)

    for i, konto in enumerate(konta, 2):
        vals = [konto["Kod_Konta"], konto["Nazwa_Konta"], konto["Typ"],
                konto["Atrybut_Podatkowy"], konto["Znacznik_KSeF"]]
        for col, v in enumerate(vals, 1):
            cell = ws.cell(row=i, column=col, value=v)
            cell.font = dfont
            cell.border = border
            if i % 2 == 0:
                cell.fill = alt_fill

    # Szerokosci kolumn
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 55
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["E"].width = 18

    # Autofiltr
    ws.auto_filter.ref = f"A1:E{len(konta)+1}"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ══════════════════════════════════════════════════════
# SIDEBAR — DANE JEDNOSTKI + KRS
# ══════════════════════════════════════════════════════

with st.sidebar:
    st.header("Dane jednostki")
    krs_input = st.text_input("Numer KRS spolki", placeholder="0000640431")
    if st.button("Pobierz dane z KRS", use_container_width=True):
        if krs_input:
            with st.spinner("Pobieranie z API KRS..."):
                try:
                    krs_data = fetch_krs_by_krs_nr(krs_input)
                    if krs_data:
                        st.session_state["krs_data"] = krs_data
                        st.success("Dane pobrane z KRS!")
                    else:
                        st.error("Nie znaleziono.")
                except Exception as e:
                    st.error(f"Blad: {e}")
        else:
            st.warning("Wpisz numer KRS.")

    krs = st.session_state.get("krs_data", {})
    st.session_state.d_name = st.text_input("Nazwa spolki", value=krs.get("nazwa", G("d_name")))
    st.session_state.d_addr = st.text_input("Siedziba", value=krs.get("siedziba", G("d_addr")))
    st.session_state.d_nip = st.text_input("NIP", value=krs.get("nip", G("d_nip")))
    st.session_state.d_krs = st.text_input("Nr KRS", value=krs.get("krs", G("d_krs")))
    st.session_state.d_regon = st.text_input("REGON", value=krs.get("regon", G("d_regon")))
    if krs.get("forma_key") and krs["forma_key"] in ENTITY_FORM_KEYS:
        dfi = ENTITY_FORM_KEYS.index(krs["forma_key"])
    else:
        dfi = G("d_form") if isinstance(G("d_form"), int) else 0
    fv = st.selectbox("Forma prawna", ENTITY_FORM_LABELS, index=dfi)
    st.session_state.d_form = ENTITY_FORM_LABELS.index(fv) if fv in ENTITY_FORM_LABELS else 0
    if krs.get("rep") and not G("d_ab"):
        st.session_state.d_ab = krs["rep"]
    st.divider()
    st.subheader("Rok obrotowy")
    st.session_state.d_fys = st.text_input("Poczatek (MM-DD)", value=G("d_fys"))
    st.session_state.d_fye = st.text_input("Koniec (MM-DD)", value=G("d_fye"))
    st.session_state.d_small = st.checkbox("Jednostka mala (art. 3 ust. 1c)", value=G("d_small"))
    st.session_state.d_micro = st.checkbox("Jednostka mikro (art. 3 ust. 1a)", value=G("d_micro"))


# ══════════════════════════════════════════════════════
# DOCX GENERATION
# ══════════════════════════════════════════════════════

def gen_docx():
    doc = Document()
    sec = doc.sections[0]; sec.page_width = Cm(21); sec.page_height = Cm(29.7)
    sec.top_margin = Cm(2.5); sec.bottom_margin = Cm(2.5); sec.left_margin = Cm(2.5); sec.right_margin = Cm(2)
    ns = doc.styles["Normal"]; ns.font.name = "Arial"; ns.font.size = Pt(11)
    ns.paragraph_format.space_after = Pt(6); ns.paragraph_format.line_spacing = 1.15
    for lv, (sz, cl) in {0: (16, "1A3C5E"), 1: (13, "2B5E8C"), 2: (11, "3B6B4F")}.items():
        h = doc.styles[f"Heading {lv+1}"]; h.font.name = "Arial"; h.font.size = Pt(sz)
        h.font.bold = True; h.font.color.rgb = RGBColor.from_string(cl)
        h.paragraph_format.space_before = Pt(18 if lv == 0 else 12); h.paragraph_format.space_after = Pt(8)
    hp = sec.header.paragraphs[0] if sec.header.paragraphs else sec.header.add_paragraph()
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    hr = hp.add_run(f"Polityka Rachunkowosci - {G('d_name')}"); hr.font.size = Pt(8)
    hr.font.color.rgb = RGBColor(153, 153, 153); hr.font.italic = True
    fp = sec.footer.paragraphs[0] if sec.footer.paragraphs else sec.footer.add_paragraph()
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rf = fp.add_run("Strona "); rf.font.size = Pt(8); rf.font.color.rgb = RGBColor(153, 153, 153)
    rp = fp.add_run(); rp.font.size = Pt(8)
    f1 = OxmlElement("w:fldChar"); f1.set(qn("w:fldCharType"), "begin")
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = " PAGE "
    f2 = OxmlElement("w:fldChar"); f2.set(qn("w:fldCharType"), "end")
    rp._r.append(f1); rp._r.append(it); rp._r.append(f2)

    def P(t, b=False):
        pp = doc.add_paragraph(); r = pp.add_run(t); r.bold = b; return pp
    def PC(t, sz=11, b=False, i=False, c=None):
        pp = doc.add_paragraph(); pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = pp.add_run(t); r.font.size = Pt(sz); r.bold = b; r.font.italic = i
        if c: r.font.color.rgb = RGBColor.from_string(c)

    efi = G("d_form"); efk = ENTITY_FORM_KEYS[efi] if isinstance(efi, int) and efi < len(ENTITY_FORM_KEYS) else ""
    efl = ENTITY_FORM_FULL.get(efk, "")
    ad = G("d_adate"); ed = G("d_edate")
    ads = ad.strftime("%d.%m.%Y") if isinstance(ad, date) else str(ad)
    eds = ed.strftime("%d.%m.%Y") if isinstance(ed, date) else str(ed)
    thr = f"{G('d_thr'):,}".replace(",", " ")

    for _ in range(4): doc.add_paragraph()
    PC("POLITYKA RACHUNKOWOSCI", 24, True)
    PC(G("d_name") or "[nazwa jednostki]", 16)
    doc.add_paragraph()
    PC("Na podstawie Ustawy z dnia 29 wrzesnia 1994 r. o rachunkowosci\n(Dz.U. z 2023 r. poz. 120 ze zm.)", 11, False, True, "666666")
    PC(f"Obowiazuje od: {eds}", 11)
    doc.add_page_break()

    doc.add_heading("I. Postanowienia ogolne", level=1)
    P('1. Polityka Rachunkowosci opracowana na podstawie Ustawy z dnia 29.09.1994 r. o rachunkowosci oraz KSR.')
    kp = f", KRS: {G('d_krs')}" if G("d_krs") else ""
    P(f"2. Jednostka: {G('d_name') or '[nazwa]'}, forma: {efl or '[forma]'}, NIP: {G('d_nip') or '[NIP]'}, REGON: {G('d_regon') or '[REGON]'}{kp}, siedziba: {G('d_addr') or '[adres]'}.")
    fys = "1 stycznia" if G("d_fys") == "01-01" else G("d_fys")
    fye = "31 grudnia" if G("d_fye") == "12-31" else G("d_fye")
    P(f"3. Rok obrotowy: od {fys} do {fye}.")
    P("4. Ksiegi w jezyku polskim, waluta PLN.")
    if G("d_small"): P("5. Jednostka mala - uproszczenia (art. 3 ust. 1c UoR).")
    elif G("d_micro"): P("5. Jednostka mikro - uproszczenia (art. 3 ust. 1a UoR).")
    else: P("5. Pelne zasady rachunkowosci.")

    doc.add_heading("II. Zakladowy Plan Kont i ksiegi rachunkowe", level=1)
    zpk = "wzorcowy plan kont" if "Wzorcowy" in G("d_zpk") else "indywidualny plan kont (wygenerowany na podstawie parametrow jednostki)"
    P(f"1. ZPK oparty o {zpk} - Zalacznik nr 1.")
    P("2. Ksiegi: dziennik, konta ksiegi glownej, konta ksiag pomocniczych, zestawienie obrotow i sald.")
    sf = G("d_sn") or "[program]"
    if G("d_sv"): sf += f", wersja: {G('d_sv')}"
    if G("d_sp"): sf += f", producent: {G('d_sp')}"
    P(f"3. System informatyczny: {sf}.")
    P("4. Opis systemu - Zalacznik nr 2.")

    doc.add_heading("III. Metody wyceny aktywow i pasywow", level=1)
    dm = {"Metoda liniowa": "liniowa", "Metoda degresywna": "degresywna", "Jednorazowa": "jednorazowo"}
    P(f"1. ST powyzej {thr} PLN - amortyzacja {dm.get(G('d_dep'), G('d_dep'))}.")
    P(f"2. Ponizej {thr} PLN - jednorazowy odpis w koszty.")
    P("3. WNiP - metoda liniowa.")
    ivm = {"Cena nabycia": "cen nabycia", "Koszt wytworzenia": "kosztu wytworzenia", "Cena rynkowa": "wartosci rynkowej"}
    P(f"4. Zapasy wg {ivm.get(G('d_iv'), G('d_iv'))}.")
    idm = {"FIFO": "FIFO", "LIFO": "LIFO", "Srednia wazona": "sredniej wazonej", "Szczegolowa identyfikacja": "szczegolowej identyfikacji"}
    P(f"5. Rozchod zapasow: {idm.get(G('d_id'), G('d_id'))}.")
    P("6. Naleznosci w kwocie wymaganej zaplaty. Odpisy aktualizujace wg zasady ostroznosci.")
    P("7. Zobowiazania w kwocie wymagajacej zaplaty. Rezerwy na prawdopodobne zobowiazania.")

    doc.add_heading("IV. Ewidencja kosztow i RZiS", level=1)
    cmm = {"Tylko Zespol 4 (uklad rodzajowy)": "wylacznie w Zespole 4",
           "Tylko Zespol 5 (uklad kalkulacyjny)": "wylacznie w Zespole 5",
           "Zespol 4 + 5 (oba uklady)": "rownolegle w Zespole 4 i 5"}
    P(f"1. Koszty {cmm.get(G('d_cm'), G('d_cm'))}.")
    plbl = "porownawczym" if "porownawczy" in G("d_pl") else "kalkulacyjnym"
    P(f"2. RZiS wariant {plbl}.")

    doc.add_heading("V. Operacje walutowe", level=1)
    fxm = {"Kurs sredni NBP": "sredni NBP", "Kurs kupna banku": "kupna banku", "Kurs sprzedazy banku": "sprzedazy banku"}
    P(f"1. Kurs: {fxm.get(G('d_fxs'), G('d_fxs'))}.")
    P("2. Dzien bilansowy - kurs sredni NBP (art. 30 ust. 1).")

    doc.add_heading("VI. Ochrona danych", level=1)
    P(f"1. Archiwizacja: {G('d_ay')} lat (art. 74 UoR).")
    bkm = {"Codziennie": "codzienna", "Co tydzien": "tygodniowa", "Co miesiac": "miesieczna"}
    P(f"2. Kopie zapasowe: {bkm.get(G('d_bk'), G('d_bk'))}.")

    doc.add_heading("VII. Zasady dodatkowe", level=1)
    P("1. Przychody: zasada memorialowa." if "memorialowa" in G("d_rev") else "1. Przychody: zasada kasowa.")
    P("2. Leasing: klasyfikacja bilansowa." if "bilansow" in G("d_ls") else "2. Leasing: klasyfikacja podatkowa.")

    doc.add_heading("VIII. Postanowienia koncowe", level=1)
    P(f"1. Wchodzi w zycie: {eds}.")
    P(f"2. Zatwierdzil(a): {G('d_ab') or 'kierownik jednostki'}.")
    P("3. Zalaczniki: ZPK, opis systemu, system ochrony danych.")

    buf = io.BytesIO(); doc.save(buf); buf.seek(0); return buf


# ══════════════════════════════════════════════════════
# WIZARD STEPS
# ══════════════════════════════════════════════════════

def step_0():
    st.subheader("Krok 1: Sprawdz dane jednostki")
    st.info("Dane jednostki wypelnij w **panelu bocznym po lewej**. Mozesz pobrac je z KRS.")
    krs = st.session_state.get("krs_data", {})
    if krs.get("nazwa"):
        st.success(f"Dane z KRS: **{krs['nazwa']}**")
    st.write(f"**Nazwa:** {G('d_name') or '-'}")
    st.write(f"**NIP:** {G('d_nip') or '-'} | **KRS:** {G('d_krs') or '-'} | **REGON:** {G('d_regon') or '-'}")
    st.write(f"**Adres:** {G('d_addr') or '-'}")


def step_1():
    st.subheader("Krok 2: Ksiegi rachunkowe i Plan Kont")

    st.session_state.d_zpk = st.radio("Zakladowy Plan Kont",
        ["Wzorcowy plan kont", "Wygeneruj plan kont na podstawie parametrow"], key="wzpk")

    if "Wygeneruj" in st.session_state.d_zpk:
        st.markdown("---")
        st.markdown("### Generator Zakladowego Planu Kont (ZPK) 2026")
        st.caption("Odpowiedz na pytania - system wygeneruje ZPK z uwzglednieniem JPK_CIT i KSeF.")

        c1, c2 = st.columns(2)
        with c1:
            zpk_branza = st.selectbox("Branza", ["Uslugi", "Handel", "Produkcja", "Hybryda"], key="zpk_br")
            zpk_cit = st.selectbox("Typ CIT", ["Klasyczny", "Estonski"], key="zpk_cit")
            zpk_rzis = st.selectbox("Wariant RZiS", ["Porownawczy", "Kalkulacyjny"], key="zpk_rzis")
        with c2:
            zpk_skala = st.selectbox("Skala podatnika", ["Maly", "Duzy"], key="zpk_sk")
            zpk_aut = st.selectbox("Analityka KUP/NKUP (samochody, limity)", ["Tak", "Nie"], key="zpk_aut")
            zpk_tp = st.selectbox("Podmioty powiazane (TP)", ["Nie", "Tak"], key="zpk_tp")

        if st.button("Generuj Plan Kont", use_container_width=True, type="primary", key="gen_zpk"):
            konta = generate_zpk(zpk_branza, zpk_cit, zpk_rzis, zpk_skala,
                                  zpk_aut == "Tak", zpk_tp == "Tak")
            st.session_state["zpk_konta"] = konta
            st.success(f"Wygenerowano {len(konta)} kont!")

        if "zpk_konta" in st.session_state:
            konta = st.session_state["zpk_konta"]
            st.markdown(f"**Wygenerowany plan: {len(konta)} kont**")

            # Podglad
            import pandas as pd
            df = pd.DataFrame(konta)
            st.dataframe(df, use_container_width=True, height=400)

            # Eksport XLSX
            xlsx_buf = zpk_to_xlsx(konta)
            if xlsx_buf:
                st.download_button("Pobierz ZPK jako XLSX", xlsx_buf,
                    f"ZPK_{(G('d_name') or 'spolka').replace(' ','_')}.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)

    st.markdown("---")
    st.markdown("**System informatyczny**")
    st.session_state.d_sn = st.text_input("Oprogramowanie", value=G("d_sn"), key="wsn", placeholder="np. Symfonia, Enova365")
    c1, c2 = st.columns(2)
    with c1:
        st.session_state.d_sv = st.text_input("Wersja", value=G("d_sv"), key="wsv")
    with c2:
        st.session_state.d_sp = st.text_input("Producent", value=G("d_sp"), key="wsp")


def step_2():
    st.subheader("Krok 3: Metody wyceny")
    st.session_state.d_dep = st.radio("Amortyzacja ST", ["Metoda liniowa", "Metoda degresywna", "Jednorazowa"], key="wdep")
    st.session_state.d_thr = st.slider("Prog ST (PLN)", 3500, 30000, G("d_thr"), 500, key="wthr")
    st.session_state.d_iv = st.radio("Wycena zapasow", ["Cena nabycia", "Koszt wytworzenia", "Cena rynkowa"], key="wiv")
    st.info("**Art. 34 ust. 4 UoR** - wybierz metode rozchodu i stosuj konsekwentnie.")
    st.session_state.d_id = st.radio("Rozchod zapasow", ["FIFO", "LIFO", "Srednia wazona", "Szczegolowa identyfikacja"], key="wid")


def step_3():
    st.subheader("Krok 4: Koszty i RZiS")
    st.session_state.d_cm = st.radio("Model kosztow", ["Tylko Zespol 4 (uklad rodzajowy)", "Tylko Zespol 5 (uklad kalkulacyjny)", "Zespol 4 + 5 (oba uklady)"], key="wcm")
    cm = st.session_state.d_cm
    if "Zespol 4" in cm and "5" not in cm:
        st.session_state.d_pl = "Wariant porownawczy"; st.info("RZiS: **porownawczy** (auto)")
    elif "Zespol 5" in cm and "4" not in cm:
        st.session_state.d_pl = "Wariant kalkulacyjny"; st.info("RZiS: **kalkulacyjny** (auto)")
    else:
        st.session_state.d_pl = st.radio("Wariant RZiS", ["Wariant porownawczy", "Wariant kalkulacyjny"], key="wpl")
    if "Zespol 5" in cm or "4 + 5" in cm:
        st.session_state.d_pc = st.radio("Kalkulacja kosztu", ["Pelny koszt wytworzenia", "Zmienny koszt wytworzenia"], key="wpc")
        st.session_state.d_oh = st.radio("Klucz kosztow posrednich", ["Klucz przychodowy", "Klucz kosztowy", "Bezposrednie przypisanie"], key="woh")


def step_4():
    st.subheader("Krok 5: Waluty obce")
    st.session_state.d_fxs = st.radio("Kurs walutowy", ["Kurs sredni NBP", "Kurs kupna banku", "Kurs sprzedazy banku"], key="wfxs")
    st.session_state.d_hfx = st.checkbox("Rachunki walutowe", value=G("d_hfx"), key="whfx")
    if st.session_state.d_hfx:
        st.session_state.d_fxd = st.radio("Rozchod waluty", ["FIFO", "LIFO", "Srednia wazona"], key="wfxd")
        st.session_state.d_cur = st.multiselect("Waluty", ALL_CUR, default=G("d_cur"), key="wcur")


def step_5():
    st.subheader("Krok 6: Ochrona danych")
    st.session_state.d_dp = st.radio("Metoda ochrony", ["Elektroniczna i fizyczna", "Wylacznie elektroniczna", "Wylacznie fizyczna"], key="wdp")
    st.session_state.d_ay = st.slider("Archiwizacja (lata)", 5, 15, G("d_ay"), key="way")
    st.session_state.d_bk = st.radio("Kopie zapasowe", ["Codziennie", "Co tydzien", "Co miesiac"], key="wbk")
    st.session_state.d_ac = st.checkbox("Kontrola dostepu z haslami", value=G("d_ac"), key="wac")
    st.session_state.d_rp = st.text_input("Osoba odpowiedzialna", value=G("d_rp"), key="wrp")


def step_6():
    st.subheader("Krok 7: Polityki dodatkowe")
    st.session_state.d_rev = st.radio("Przychody", ["Zasada memorialowa", "Zasada kasowa"], key="wrev")
    st.session_state.d_ls = st.radio("Leasing", ["Wg przepisow bilansowych", "Wg przepisow podatkowych"], key="wls")
    st.session_state.d_prov = st.checkbox("Rezerwy (art. 35d)", value=G("d_prov"), key="wprov")
    st.session_state.d_dt = st.checkbox("Podatek odroczony", value=G("d_dt"), key="wdt")
    if not (G("d_small") or G("d_micro")):
        st.session_state.d_cf = st.radio("Przeplywy pieniezne", ["Metoda posrednia", "Metoda bezposrednia"], key="wcf")
    st.markdown("**Zatwierdzenie**")
    c1, c2 = st.columns(2)
    with c1: st.session_state.d_adate = st.date_input("Data zatwierdzenia", value=G("d_adate"), key="wad")
    with c2: st.session_state.d_edate = st.date_input("Data wejscia w zycie", value=G("d_edate"), key="wed")
    st.session_state.d_ab = st.text_input("Zatwierdzil(a)", value=G("d_ab"), key="wab", placeholder="Imie i nazwisko")


def step_7():
    st.subheader("Krok 8: Eksport DOCX")
    buf = gen_docx()
    fn = f"Polityka_Rachunkowosci_{(G('d_name') or 'jednostka').replace(' ', '_')}.docx"
    st.download_button("Pobierz Polityke Rachunkowosci (DOCX)", buf, fn,
                       "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                       use_container_width=True, type="primary")

    if "zpk_konta" in st.session_state:
        xlsx_buf = zpk_to_xlsx(st.session_state["zpk_konta"])
        if xlsx_buf:
            st.download_button("Pobierz Zakladowy Plan Kont (XLSX)", xlsx_buf,
                f"ZPK_{(G('d_name') or 'spolka').replace(' ', '_')}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)

    st.success("Dokumenty gotowe do pobrania!")
    st.divider()
    efi = G("d_form")
    efl = ENTITY_FORM_LABELS[efi] if isinstance(efi, int) and efi < len(ENTITY_FORM_LABELS) else ""
    with st.expander("Podglad danych", expanded=True):
        st.write(f"**{G('d_name') or '-'}** ({efl})")
        st.write(f"NIP: {G('d_nip') or '-'} | KRS: {G('d_krs') or '-'} | REGON: {G('d_regon') or '-'}")
        st.write(f"Koszty: {G('d_cm')} | RZiS: {G('d_pl')}")
        st.write(f"Zatwierdzil(a): {G('d_ab') or '-'}")


# ══════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════

STEPS = [step_0, step_1, step_2, step_3, step_4, step_5, step_6, step_7]
st.title("Generator Polityki Rachunkowosci")
st.caption("Zgodna z Ustawa o Rachunkowosci (art. 10 UoR) | Stan prawny 2026")

prog = st.session_state.step / max(len(STEPS) - 1, 1)
st.progress(prog, text=f"**{STEP_NAMES[st.session_state.step]}** ({st.session_state.step+1}/{len(STEPS)})")

STEPS[st.session_state.step]()

st.divider()
c1, c2, c3 = st.columns([1, 2, 1])
with c1:
    if st.session_state.step > 0 and st.button("Wstecz", use_container_width=True, key="bk"):
        st.session_state.step -= 1; st.rerun()
with c2:
    st.markdown(f"<p style='text-align:center;color:#999;margin-top:8px'>{st.session_state.step+1} / {len(STEPS)}</p>", unsafe_allow_html=True)
with c3:
    if st.session_state.step < len(STEPS) - 1 and st.button("Dalej", use_container_width=True, type="primary", key="fw"):
        st.session_state.step += 1; st.rerun()
