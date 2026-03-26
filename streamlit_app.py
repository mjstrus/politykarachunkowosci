import streamlit as st
import requests
import io
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from datetime import date

st.set_page_config(page_title="Generator Polityki Rachunkowosci", page_icon="\U0001F4CB", layout="centered")

ENTITY_FORM_LABELS = ["Sp. z o.o.", "Spolka akcyjna", "Spolka cywilna", "Spolka jawna", "Spolka komandytowa", "Spolka kom.-akcyjna", "JDG", "Fundacja", "Stowarzyszenie"]
ENTITY_FORM_KEYS = ["sp_zoo", "sa", "sc", "sj", "sk", "ska", "jdg", "fundacja", "stowarzyszenie"]
ENTITY_FORM_FULL = {"sp_zoo":"Spolka z ograniczona odpowiedzialnoscia","sa":"Spolka akcyjna","sc":"Spolka cywilna","sj":"Spolka jawna","sk":"Spolka komandytowa","ska":"Spolka komandytowo-akcyjna","jdg":"Jednoosobowa dzialalnosc gospodarcza","fundacja":"Fundacja","stowarzyszenie":"Stowarzyszenie"}
STEP_NAMES = ["Dane jednostki","Ksiegi rachunkowe","Metody wyceny","Koszty i RZiS","Waluty obce","Ochrona danych","Polityki dodatkowe","Podglad i eksport"]
ALL_CUR = ["EUR","USD","GBP","CHF","CZK","SEK","NOK","DKK","JPY","CNY"]

DEFS = dict(step=0,d_name="",d_form=0,d_nip="",d_krs="",d_regon="",d_addr="",d_fys="01-01",d_fye="12-31",d_small=False,d_micro=False,d_zpk="Wzorcowy plan kont",d_sn="",d_sv="",d_sp="",d_dep="Metoda liniowa",d_thr=10000,d_iv="Cena nabycia",d_id="FIFO",d_cm="Tylko Zespol 4 (uklad rodzajowy)",d_pl="Wariant porownawczy",d_pc="Pelny koszt wytworzenia",d_oh="Klucz przychodowy",d_fxs="Kurs sredni NBP",d_fxd="FIFO",d_hfx=False,d_cur=["EUR","USD"],d_dp="Elektroniczna i fizyczna",d_ay=5,d_bk="Codziennie",d_ac=True,d_rp="",d_rev="Zasada memorialowa",d_ls="Wg przepisow bilansowych",d_prov=True,d_dt=True,d_cf="Metoda posrednia",d_adate=date.today(),d_edate=date.today(),d_ab="")
for k,v in DEFS.items():
    if k not in st.session_state:
        st.session_state[k]=v

def G(k): return st.session_state.get(k, DEFS.get(k,""))

# ── KRS API ──
def fetch_krs(nr):
    nr=nr.strip().replace("-","").replace(" ","")
    if not nr.isdigit(): return {"error":"Numer KRS musi zawierac tylko cyfry."}
    nr=nr.zfill(10)
    headers = {
        "Accept": "application/json",
        "User-Agent": "Mozilla/5.0 (compatible; PolitikaRachunkowosci/1.0)"
    }
    url = f"https://api-krs.ms.gov.pl/api/krs/OdpisAktualny/{nr}"
    for reg in ["P","S"]:
        try:
            r=requests.get(url, params={"rejestr": reg, "format": "json"}, headers=headers, timeout=20)
            if r.status_code==200: return _parse(r.json(),nr)
            elif r.status_code==404: continue
            else: return {"error":f"HTTP {r.status_code}"}
        except requests.exceptions.Timeout: return {"error":"Timeout API KRS."}
        except requests.exceptions.ConnectionError: return {"error":"Brak polaczenia z API KRS."}
        except Exception as e: return {"error":str(e)}
    return {"error":f"Nie znaleziono KRS {nr}."}

def _parse(data,nr):
    res={"krs":nr}

    # Real KRS API structure: {"odpis": {"dane": {"dzial1": {...}, "dzial2": {...}}}}
    root = data
    if isinstance(root, dict) and "odpis" in root:
        root = root["odpis"]
    if isinstance(root, dict) and "dane" in root:
        root = root["dane"]

    d1 = root.get("dzial1", {}) if isinstance(root, dict) else {}
    dp = d1.get("danePodmiotu", {})
    idn = dp.get("identyfikatory", {})
    siedziba = d1.get("siedzibaIAdres", {})
    adr = siedziba.get("adres", {})

    res["nazwa"] = dp.get("nazwa", "")
    res["nip"] = idn.get("nip", "")
    res["regon"] = idn.get("regon", "")

    # Address - KRS uses uppercase field names
    parts = []
    ulica = adr.get("ulica", "")
    # KRS sometimes prefixes with "UL. " already
    if ulica:
        nr_domu = adr.get("nrDomu", "")
        nr_lok = adr.get("nrLokalu", "")
        lok_part = f"/{nr_lok}" if nr_lok else ""
        if ulica.upper().startswith("UL."):
            parts.append(f"{ulica} {nr_domu}{lok_part}".strip())
        else:
            parts.append(f"ul. {ulica} {nr_domu}{lok_part}".strip())
    km = " ".join(filter(None, [adr.get("kodPocztowy", ""), adr.get("miejscowosc", "")]))
    if km:
        parts.append(km)
    res["adres"] = ", ".join(parts)

    # Entity form - KRS returns full text like "SPOLKA Z OGRANICZONA ODPOWIEDZIALNOSCIA"
    fp = dp.get("formaPrawna", "")
    if isinstance(fp, dict):
        fp = fp.get("nazwa", "")
    fl = fp.lower()
    res["forma"] = ("sp_zoo" if "ograniczon" in fl else
                    "ska" if "komandytowo-akcyjn" in fl else
                    "sk" if "komandytow" in fl else
                    "sa" if "akcyjn" in fl else
                    "sj" if "jawn" in fl else
                    "fundacja" if "fundacj" in fl else
                    "stowarzyszenie" if "stowarzysz" in fl else "")

    # Representation - dzial2 > reprezentacja > sklad[]
    # Names have nested structure: {"nazwisko": {"nazwiskoICzlon": "X"}, "imiona": {"imie": "Y"}}
    d2 = root.get("dzial2", {}) if isinstance(root, dict) else {}
    rep_data = d2.get("reprezentacja", {})
    sklad = rep_data.get("sklad", [])
    if sklad:
        o = sklad[0]
        # Extract name from nested dicts
        nazwisko_obj = o.get("nazwisko", {})
        imiona_obj = o.get("imiona", {})
        if isinstance(nazwisko_obj, dict):
            nz = nazwisko_obj.get("nazwiskoICzlon", "")
        else:
            nz = str(nazwisko_obj)
        if isinstance(imiona_obj, dict):
            im = imiona_obj.get("imie", "")
        else:
            im = str(imiona_obj)
        fn = o.get("funkcjaWOrganie", o.get("funkcja", ""))
        rep = f"{im} {nz}".strip()
        if fn:
            rep += f" - {fn}"
        res["rep"] = rep
    else:
        res["rep"] = ""

    return res

# ── DOCX ──
def gen_docx():
    doc=Document()
    sec=doc.sections[0]; sec.page_width=Cm(21); sec.page_height=Cm(29.7)
    sec.top_margin=Cm(2.5); sec.bottom_margin=Cm(2.5); sec.left_margin=Cm(2.5); sec.right_margin=Cm(2)
    ns=doc.styles["Normal"]; ns.font.name="Arial"; ns.font.size=Pt(11)
    ns.paragraph_format.space_after=Pt(6); ns.paragraph_format.line_spacing=1.15
    for lv,(sz,cl) in {0:(16,"1A3C5E"),1:(13,"2B5E8C"),2:(11,"3B6B4F")}.items():
        h=doc.styles[f"Heading {lv+1}"]; h.font.name="Arial"; h.font.size=Pt(sz); h.font.bold=True; h.font.color.rgb=RGBColor.from_string(cl)
        h.paragraph_format.space_before=Pt(18 if lv==0 else 12); h.paragraph_format.space_after=Pt(8)
    hp=sec.header.paragraphs[0] if sec.header.paragraphs else sec.header.add_paragraph()
    hp.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    hr=hp.add_run(f"Polityka Rachunkowosci - {G('d_name')}"); hr.font.size=Pt(8); hr.font.color.rgb=RGBColor(153,153,153); hr.font.italic=True
    fp=sec.footer.paragraphs[0] if sec.footer.paragraphs else sec.footer.add_paragraph()
    fp.alignment=WD_ALIGN_PARAGRAPH.CENTER
    rf=fp.add_run("Strona "); rf.font.size=Pt(8); rf.font.color.rgb=RGBColor(153,153,153)
    rp=fp.add_run(); rp.font.size=Pt(8)
    f1=OxmlElement("w:fldChar"); f1.set(qn("w:fldCharType"),"begin")
    it=OxmlElement("w:instrText"); it.set(qn("xml:space"),"preserve"); it.text=" PAGE "
    f2=OxmlElement("w:fldChar"); f2.set(qn("w:fldCharType"),"end")
    rp._r.append(f1); rp._r.append(it); rp._r.append(f2)

    def P(t,b=False):
        pp=doc.add_paragraph(); r=pp.add_run(t); r.bold=b; return pp
    def PC(t,sz=11,b=False,i=False,c=None):
        pp=doc.add_paragraph(); pp.alignment=WD_ALIGN_PARAGRAPH.CENTER; r=pp.add_run(t); r.font.size=Pt(sz); r.bold=b; r.font.italic=i
        if c: r.font.color.rgb=RGBColor.from_string(c)

    efi=G("d_form"); efk=ENTITY_FORM_KEYS[efi] if isinstance(efi,int) and efi<len(ENTITY_FORM_KEYS) else ""
    efl=ENTITY_FORM_FULL.get(efk,"")
    ad=G("d_adate"); ed=G("d_edate")
    ads=ad.strftime("%d.%m.%Y") if isinstance(ad,date) else str(ad)
    eds=ed.strftime("%d.%m.%Y") if isinstance(ed,date) else str(ed)
    thr=f"{G('d_thr'):,}".replace(","," ")

    for _ in range(4): doc.add_paragraph()
    PC("POLITYKA RACHUNKOWOSCI",24,True)
    PC(G("d_name") or "[nazwa jednostki]",16)
    doc.add_paragraph()
    PC("Na podstawie Ustawy z dnia 29 wrzesnia 1994 r. o rachunkowosci\n(Dz.U. z 2023 r. poz. 120 ze zm.)",11,False,True,"666666")
    PC(f"Obowiazuje od: {eds}",11)
    doc.add_page_break()

    doc.add_heading("I. Postanowienia ogolne",level=1)
    P('1. Polityka Rachunkowosci opracowana na podstawie Ustawy z dnia 29.09.1994 r. o rachunkowosci (Dz.U. z 2023 r. poz. 120 ze zm.) oraz Krajowych Standardow Rachunkowosci.')
    kp=f", KRS: {G('d_krs')}" if G("d_krs") else ""
    P(f"2. Jednostka: {G('d_name') or '[nazwa]'}, forma: {efl or '[forma]'}, NIP: {G('d_nip') or '[NIP]'}, REGON: {G('d_regon') or '[REGON]'}{kp}, siedziba: {G('d_addr') or '[adres]'}.")
    fys="1 stycznia" if G("d_fys")=="01-01" else G("d_fys")
    fye="31 grudnia" if G("d_fye")=="12-31" else G("d_fye")
    P(f"3. Rok obrotowy: od {fys} do {fye}.")
    P("4. Ksiegi w jezyku polskim, waluta PLN.")
    if G("d_small"): P("5. Jednostka mala (art. 3 ust. 1c UoR) - uproszczenia.")
    elif G("d_micro"): P("5. Jednostka mikro (art. 3 ust. 1a UoR) - uproszczenia.")
    else: P("5. Pelne zasady rachunkowosci.")

    doc.add_heading("II. Zakladowy Plan Kont i ksiegi rachunkowe",level=1)
    zpk="wzorcowy plan kont" if "Wzorcowy" in G("d_zpk") else "indywidualny plan kont"
    P(f"1. ZPK oparty o {zpk} - Zalacznik nr 1.")
    P("2. Ksiegi: dziennik, konta ksiegi glownej, konta ksiag pomocniczych, zestawienie obrotow i sald.")
    sf=G("d_sn") or "[program]"
    if G("d_sv"): sf+=f", wersja: {G('d_sv')}"
    if G("d_sp"): sf+=f", producent: {G('d_sp')}"
    P(f"3. System informatyczny: {sf}.")
    P("4. Opis systemu - Zalacznik nr 2.")

    doc.add_heading("III. Metody wyceny aktywow i pasywow",level=1)
    doc.add_heading("A. Srodki trwale i WNiP",level=2)
    dm={"Metoda liniowa":"liniowa","Metoda degresywna":"degresywna","Jednorazowa":"jednorazowo"}
    P(f"1. ST powyzej {thr} PLN - amortyzacja {dm.get(G('d_dep'),G('d_dep'))}.")
    P(f"2. Ponizej {thr} PLN - jednorazowy odpis w koszty.")
    P("3. WNiP - metoda liniowa.")
    doc.add_heading("B. Zapasy",level=2)
    ivm={"Cena nabycia":"cen nabycia","Koszt wytworzenia":"kosztu wytworzenia","Cena rynkowa":"wartosci rynkowej"}
    P(f"4. Zapasy wg {ivm.get(G('d_iv'),G('d_iv'))}.")
    idm={"FIFO":"FIFO","LIFO":"LIFO","Srednia wazona":"sredniej wazonej","Szczegolowa identyfikacja":"szczegolowej identyfikacji"}
    P(f"5. Rozchod: {idm.get(G('d_id'),G('d_id'))}.")
    P("6. Odpisy aktualizujace przy utracie wartosci.")
    doc.add_heading("C. Naleznosci",level=2)
    P("7. W kwocie wymaganej zaplaty, po pomniejszeniu o odpisy.")
    doc.add_heading("D. Inwestycje",level=2)
    P("8. Wg ceny nabycia lub wartosci rynkowej (nizsza).")
    doc.add_heading("E. Zobowiazania",level=2)
    P("9. W kwocie wymagajacej zaplaty. Rezerwy na prawdopodobne zobowiazania.")

    doc.add_heading("IV. Ewidencja kosztow i RZiS",level=1)
    cmm={"Tylko Zespol 4 (uklad rodzajowy)":"wylacznie w Zespole 4","Tylko Zespol 5 (uklad kalkulacyjny)":"wylacznie w Zespole 5","Zespol 4 + 5 (oba uklady)":"rownolegle w Zespole 4 i 5"}
    P(f"1. Koszty {cmm.get(G('d_cm'),G('d_cm'))}.")
    plbl="porownawczym" if "porownawczy" in G("d_pl") else "kalkulacyjnym"
    atn="4" if G("d_micro") else "5" if G("d_small") else "1"
    P(f"2. RZiS wariant {plbl} (Zal. nr {atn} UoR).")
    if "Zespol 5" in G("d_cm") or "4 + 5" in G("d_cm"):
        pcl="pelnego kosztu" if "Pelny" in G("d_pc") else "zmiennego kosztu"
        P(f"3. Koszt wytworzenia metoda {pcl}.")
        ohm={"Klucz przychodowy":"kluczem przychodowym","Klucz kosztowy":"kluczem kosztowym","Bezposrednie przypisanie":"bezposrednim przypisaniem"}
        P(f"4. Koszty posrednie: {ohm.get(G('d_oh'),G('d_oh'))}.")
    else: P("3. Koszty w ukladzie rodzajowym.")

    doc.add_heading("V. Operacje walutowe",level=1)
    fxm={"Kurs sredni NBP":"sredni NBP","Kurs kupna banku":"kupna banku","Kurs sprzedazy banku":"sprzedazy banku"}
    P(f"1. Kurs: {fxm.get(G('d_fxs'),G('d_fxs'))}.")
    P("2. Dzien bilansowy - kurs sredni NBP (art. 30 ust. 1).")
    P("3. Roznice kursowe na przychody/koszty finansowe.")
    if G("d_hfx"):
        cdm={"FIFO":"FIFO","LIFO":"LIFO","Srednia wazona":"sredniej wazonej"}
        P(f"4. Rozchod walut: {cdm.get(G('d_fxd'),G('d_fxd'))}.")
        cur=G("d_cur")
        if isinstance(cur,list) and cur: P(f"5. Rachunki walutowe: {', '.join(cur)}.")
    else: P("4. Brak odrebnych rachunkow walutowych.")

    doc.add_heading("VI. Ochrona danych",level=1)
    dpm={"Elektroniczna i fizyczna":"elektroniczna i fizyczna","Wylacznie elektroniczna":"wylacznie elektroniczna","Wylacznie fizyczna":"wylacznie fizyczna"}
    P(f"1. Ochrona: {dpm.get(G('d_dp'),G('d_dp'))}.")
    P(f"2. Archiwizacja: {G('d_ay')} lat (art. 74 UoR).")
    bkm={"Codziennie":"codzienna","Co tydzien":"tygodniowa","Co miesiac":"miesieczna"}
    P(f"3. Kopie zapasowe: {bkm.get(G('d_bk'),G('d_bk'))}.")
    P("4. Indywidualne hasla, zmiana co 90 dni." if G("d_ac") else "4. Odpowiednia kontrola dostepu.")
    P(f"5. Odpowiedzialny: {G('d_rp') or '[imie i nazwisko]'}.")
    P("6. Procedury awaryjne - Zalacznik nr 3.")

    doc.add_heading("VII. Zasady dodatkowe",level=1)
    P("1. Przychody: zasada memorialowa." if "memorialowa" in G("d_rev") else "1. Przychody: zasada kasowa.")
    P("2. Leasing: klasyfikacja bilansowa (art. 3 ust. 4-6)." if "bilansow" in G("d_ls") else "2. Leasing: klasyfikacja podatkowa.")
    if G("d_prov"): P("3. Rezerwy na znane ryzyko (art. 35d). RMK wg art. 39.")
    if G("d_dt"):
        if G("d_small") or G("d_micro"): P("4. Zaniechanie odroczonego podatku (art. 37 ust. 10).")
        else: P("4. Ustalanie aktywow/rezerw odroczonego podatku (art. 37).")
    if G("d_small") or G("d_micro"): P("5. Zwolnienie z rachunku przeplywow.")
    else:
        cfl="posrednia" if "posrednia" in G("d_cf") else "bezposrednia"
        P(f"5. Przeplywy pieniezne: metoda {cfl}.")

    doc.add_heading("VIII. Postanowienia koncowe",level=1)
    P(f"1. Wchodzi w zycie: {eds}.")
    P(f"2. Zmiany zatwierdzane przez {G('d_ab') or 'kierownika jednostki'}.")
    P("3. Odpowiedzialnosc: kierownik jednostki (art. 4 ust. 5).")
    P("4. Zalaczniki:"); doc.add_paragraph("Zal. 1: Zakladowy Plan Kont",style="List Bullet"); doc.add_paragraph("Zal. 2: Opis systemu informatycznego",style="List Bullet"); doc.add_paragraph("Zal. 3: System ochrony danych",style="List Bullet")
    doc.add_paragraph()
    P(f"Zatwierdzil(a): {G('d_ab') or '___________________________'}",True)
    P(f"Data: {ads}")
    buf=io.BytesIO(); doc.save(buf); buf.seek(0); return buf

# ── STEPS ──
def step_0():
    st.subheader("Krok 1: Dane jednostki")
    if st.button("TEST POLACZENIA", key="test_btn"):
        try:
            r = requests.get("https://httpbin.org/get", timeout=10)
            st.write(f"httpbin status: {r.status_code}")
            r2 = requests.get("https://api-krs.ms.gov.pl/api/krs/OdpisAktualny/0000640431", 
                             params={"rejestr":"P","format":"json"},
                             headers={"Accept":"application/json","User-Agent":"Mozilla/5.0"},
                             timeout=20)
            st.write(f"KRS API status: {r2.status_code}")
            if r2.status_code == 200:
                data = r2.json()
                st.write(f"Klucze: {list(data.keys())}")
        except Exception as e:
            st.error(f"Blad: {type(e).__name__}: {e}")
    st.markdown("**Pobierz dane z KRS**")
    with st.form("krs_form"):
        krs_val = st.text_input("Numer KRS", placeholder="np. 0000640431")
        submitted = st.form_submit_button("Pobierz dane z KRS", use_container_width=True, type="primary")

    if submitted and krs_val and krs_val.strip():
        with st.spinner("Pobieranie z API KRS..."):
            try:
                res = fetch_krs(krs_val.strip())
                if res.get("error"):
                    st.error(f"Blad: {res['error']}")
                elif res.get("nazwa"):
                    st.session_state["krs_data"] = res
                    st.success("Dane pobrane z KRS!")
                    st.rerun()
                else:
                    st.warning("Nie znaleziono danych podmiotu.")
            except Exception as e:
                st.error(f"Wyjatek: {str(e)}")
    elif submitted:
        st.warning("Wpisz numer KRS.")

    krs = st.session_state.get("krs_data", {})

    st.divider()
    st.session_state.d_name = st.text_input("Nazwa jednostki", value=krs.get("nazwa", G("d_name")), key="wn")
    fv = st.selectbox("Forma prawna", ENTITY_FORM_LABELS,
                       index=ENTITY_FORM_KEYS.index(krs["forma"]) if krs.get("forma") and krs["forma"] in ENTITY_FORM_KEYS else G("d_form"),
                       key="wf")
    st.session_state.d_form = ENTITY_FORM_LABELS.index(fv) if fv in ENTITY_FORM_LABELS else 0
    c1, c2 = st.columns(2)
    with c1:
        st.session_state.d_nip = st.text_input("NIP", value=krs.get("nip", G("d_nip")), key="wnip")
    with c2:
        st.session_state.d_krs = st.text_input("KRS", value=krs.get("krs", G("d_krs")), key="wkrs")
    c3, c4 = st.columns(2)
    with c3:
        st.session_state.d_regon = st.text_input("REGON", value=krs.get("regon", G("d_regon")), key="wreg")
    with c4:
        st.session_state.d_addr = st.text_input("Adres", value=krs.get("adres", G("d_addr")), key="wadr")
    st.markdown("**Rok obrotowy**")
    c5, c6 = st.columns(2)
    with c5:
        st.session_state.d_fys = st.text_input("Poczatek (MM-DD)", value=G("d_fys"), key="wfys")
    with c6:
        st.session_state.d_fye = st.text_input("Koniec (MM-DD)", value=G("d_fye"), key="wfye")
    st.session_state.d_small = st.checkbox("Jednostka mala (art. 3 ust. 1c)", value=G("d_small"), key="wsm")
    st.session_state.d_micro = st.checkbox("Jednostka mikro (art. 3 ust. 1a)", value=G("d_micro"), key="wmi")
    if krs.get("rep") and not G("d_ab"):
        st.session_state.d_ab = krs["rep"]

def step_1():
    st.subheader("Krok 2: Ksiegi rachunkowe")
    st.session_state.d_zpk=st.radio("Zakladowy Plan Kont",["Wzorcowy plan kont","Indywidualny plan kont"],key="wzpk")
    st.markdown("**System informatyczny**")
    st.session_state.d_sn=st.text_input("Oprogramowanie",value=G("d_sn"),key="wsn",placeholder="np. Symfonia, Enova365")
    c1,c2=st.columns(2)
    with c1: st.session_state.d_sv=st.text_input("Wersja",value=G("d_sv"),key="wsv")
    with c2: st.session_state.d_sp=st.text_input("Producent",value=G("d_sp"),key="wsp")

def step_2():
    st.subheader("Krok 3: Metody wyceny")
    st.session_state.d_dep=st.radio("Amortyzacja ST",["Metoda liniowa","Metoda degresywna","Jednorazowa"],key="wdep")
    st.session_state.d_thr=st.slider("Prog ST (PLN)",3500,30000,G("d_thr"),500,key="wthr")
    st.session_state.d_iv=st.radio("Wycena zapasow",["Cena nabycia","Koszt wytworzenia","Cena rynkowa"],key="wiv")
    st.info("**Art. 34 ust. 4 UoR** - wybierz metode rozchodu i stosuj konsekwentnie.")
    st.session_state.d_id=st.radio("Rozchod zapasow",["FIFO","LIFO","Srednia wazona","Szczegolowa identyfikacja"],key="wid")

def step_3():
    st.subheader("Krok 4: Koszty i RZiS")
    st.session_state.d_cm=st.radio("Model kosztow",["Tylko Zespol 4 (uklad rodzajowy)","Tylko Zespol 5 (uklad kalkulacyjny)","Zespol 4 + 5 (oba uklady)"],key="wcm")
    cm=st.session_state.d_cm
    if "Zespol 4" in cm and "5" not in cm:
        st.session_state.d_pl="Wariant porownawczy"; st.info("RZiS: **porownawczy** (auto)")
    elif "Zespol 5" in cm and "4" not in cm:
        st.session_state.d_pl="Wariant kalkulacyjny"; st.info("RZiS: **kalkulacyjny** (auto)")
    else: st.session_state.d_pl=st.radio("Wariant RZiS",["Wariant porownawczy","Wariant kalkulacyjny"],key="wpl")
    if "Zespol 5" in cm or "4 + 5" in cm:
        st.session_state.d_pc=st.radio("Kalkulacja kosztu",["Pelny koszt wytworzenia","Zmienny koszt wytworzenia"],key="wpc")
        st.session_state.d_oh=st.radio("Klucz kosztow posrednich",["Klucz przychodowy","Klucz kosztowy","Bezposrednie przypisanie"],key="woh")

def step_4():
    st.subheader("Krok 5: Waluty obce")
    st.session_state.d_fxs=st.radio("Kurs walutowy",["Kurs sredni NBP","Kurs kupna banku","Kurs sprzedazy banku"],key="wfxs")
    st.session_state.d_hfx=st.checkbox("Rachunki walutowe",value=G("d_hfx"),key="whfx")
    if st.session_state.d_hfx:
        st.session_state.d_fxd=st.radio("Rozchod waluty",["FIFO","LIFO","Srednia wazona"],key="wfxd")
        st.session_state.d_cur=st.multiselect("Waluty",ALL_CUR,default=G("d_cur"),key="wcur")

def step_5():
    st.subheader("Krok 6: Ochrona danych")
    st.session_state.d_dp=st.radio("Metoda ochrony",["Elektroniczna i fizyczna","Wylacznie elektroniczna","Wylacznie fizyczna"],key="wdp")
    st.session_state.d_ay=st.slider("Archiwizacja (lata)",5,15,G("d_ay"),key="way")
    st.session_state.d_bk=st.radio("Kopie zapasowe",["Codziennie","Co tydzien","Co miesiac"],key="wbk")
    st.session_state.d_ac=st.checkbox("Kontrola dostepu z haslami",value=G("d_ac"),key="wac")
    st.session_state.d_rp=st.text_input("Osoba odpowiedzialna",value=G("d_rp"),key="wrp")

def step_6():
    st.subheader("Krok 7: Polityki dodatkowe")
    st.session_state.d_rev=st.radio("Przychody",["Zasada memorialowa","Zasada kasowa"],key="wrev")
    st.session_state.d_ls=st.radio("Leasing",["Wg przepisow bilansowych","Wg przepisow podatkowych"],key="wls")
    st.session_state.d_prov=st.checkbox("Rezerwy (art. 35d)",value=G("d_prov"),key="wprov")
    st.session_state.d_dt=st.checkbox("Podatek odroczony",value=G("d_dt"),key="wdt")
    if not(G("d_small") or G("d_micro")):
        st.session_state.d_cf=st.radio("Przeplywy pieniezne",["Metoda posrednia","Metoda bezposrednia"],key="wcf")
    st.markdown("**Zatwierdzenie**")
    c1,c2=st.columns(2)
    with c1: st.session_state.d_adate=st.date_input("Data zatwierdzenia",value=G("d_adate"),key="wad")
    with c2: st.session_state.d_edate=st.date_input("Data wejscia w zycie",value=G("d_edate"),key="wed")
    st.session_state.d_ab=st.text_input("Zatwierdzil(a)",value=G("d_ab"),key="wab",placeholder="Imie i nazwisko")

def step_7():
    st.subheader("Krok 8: Eksport DOCX")
    buf=gen_docx()
    fn=f"Polityka_Rachunkowosci_{(G('d_name') or 'jednostka').replace(' ','_')}.docx"
    st.download_button("Pobierz plik DOCX",buf,fn,"application/vnd.openxmlformats-officedocument.wordprocessingml.document",use_container_width=True,type="primary")
    st.success("Dokument gotowy!")
    st.divider()
    efi=G("d_form"); efl=ENTITY_FORM_LABELS[efi] if isinstance(efi,int) and efi<len(ENTITY_FORM_LABELS) else ""
    with st.expander("Podglad danych",expanded=True):
        st.write(f"**{G('d_name') or '-'}** ({efl})")
        st.write(f"NIP: {G('d_nip') or '-'} | KRS: {G('d_krs') or '-'} | REGON: {G('d_regon') or '-'}")
        st.write(f"Adres: {G('d_addr') or '-'}")
        st.write(f"Koszty: {G('d_cm')} | RZiS: {G('d_pl')}")
        st.write(f"Zapasy: {G('d_iv')} / {G('d_id')}")
        st.write(f"Amortyzacja: {G('d_dep')} | Prog: {G('d_thr'):,} PLN")
        st.write(f"Zatwierdzil(a): {G('d_ab') or '-'}")

# ── MAIN ──
STEPS=[step_0,step_1,step_2,step_3,step_4,step_5,step_6,step_7]
st.title("Generator Polityki Rachunkowosci")
st.caption("Zgodna z Ustawa o Rachunkowosci (art. 10 UoR)")
prog=st.session_state.step/max(len(STEPS)-1,1)
st.progress(prog,text=f"**{STEP_NAMES[st.session_state.step]}** ({st.session_state.step+1}/{len(STEPS)})")
STEPS[st.session_state.step]()
st.divider()
c1,c2,c3=st.columns([1,2,1])
with c1:
    if st.session_state.step>0 and st.button("Wstecz",use_container_width=True,key="bk"):
        st.session_state.step-=1; st.rerun()
with c2: st.markdown(f"<p style='text-align:center;color:#999;margin-top:8px'>{st.session_state.step+1} / {len(STEPS)}</p>",unsafe_allow_html=True)
with c3:
    if st.session_state.step<len(STEPS)-1 and st.button("Dalej",use_container_width=True,type="primary",key="fw"):
        st.session_state.step+=1; st.rerun()
