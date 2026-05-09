import datetime
import io
import re
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
import xlsxwriter
from supabase import create_client

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import RGBColor


PROGRAM_PLAN_INDEX = 0
PROGRAM_BILDIRILER_SHEET = "Bildiriler"
PROGRAM_OZEL_SHEET = "Ozel_Etkinlikler"
PROGRAM_MOD_SHEET = "Moderatorler"


# =========================
# TEMIZLEME VE ESLESTIRME
# =========================

def safe_str(value) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()


def turkce_buyuk(metin):
    if not isinstance(metin, str) or not metin:
        return ""
    return (
        metin.replace("i", "İ")
        .replace("ı", "I")
        .replace("ş", "Ş")
        .replace("ğ", "Ğ")
        .replace("ü", "Ü")
        .replace("ö", "Ö")
        .replace("ç", "Ç")
        .upper()
    )


def temiz_metin(metin):
    if not isinstance(metin, str) or not metin:
        return ""
    m = (
        metin.upper()
        .replace("İ", "I")
        .replace("Ğ", "G")
        .replace("Ü", "U")
        .replace("Ş", "S")
        .replace("Ö", "O")
        .replace("Ç", "C")
    )
    m = re.sub(r"[^A-Z0-9 ]", "", m)
    return re.sub(r"\s+", " ", m).strip()


def super_temiz(metin):
    if not isinstance(metin, str) or not metin:
        return ""
    m = turkce_buyuk(metin)
    return (
        m.replace("Ö", "O")
        .replace("Ü", "U")
        .replace("Ş", "S")
        .replace("Ç", "C")
        .replace("Ğ", "G")
        .replace("İ", "I")
    )


def temiz_isim(isim):
    if not isinstance(isim, str) or str(isim).strip() == "":
        return ""
    i = (
        str(isim)
        .upper()
        .replace("İ", "I")
        .replace("Ğ", "G")
        .replace("Ü", "U")
        .replace("Ş", "S")
        .replace("Ö", "O")
        .replace("Ç", "C")
    )
    unvan_paternleri = [
        r"^PROF\.?\s*DR\.?",
        r"^DOC\.?\s*DR\.?",
        r"^DOÇ\.?\s*DR\.?",
        r"^DR\.?\s*OGR\.?\s*UYESI",
        r"^DR\.?\s*ÖĞR\.?\s*ÜYESI",
        r"^DR\.?",
        r"^UZM\.?",
        r"^ARAS\.?\s*GOR\.?",
        r"^ARŞ\.?\s*GÖR\.?",
        r"^ARS\.?\s*GOR\.?",
        r"^ASSOC\.?\s*PROF\.?",
        r"^ASSIST\.?\s*PROF\.?",
        r"^PROF\.?",
    ]
    for patern in unvan_paternleri:
        i = re.sub(patern, "", i).strip()
    return re.sub(r"\s+", " ", i).strip()


def unvan_bul(row, columns, raw_val):
    for col in columns:
        if "unvan" in str(col).lower():
            u = str(row.get(col, "")).strip()
            if u and u.lower() != "nan":
                return u
    raw_upper = str(raw_val).upper()
    if "ASSOC. PROF" in raw_upper or "ASSOC PROF" in raw_upper:
        return "Assoc. Prof."
    if "ASSIST. PROF" in raw_upper or "ASSIST PROF" in raw_upper:
        return "Assist. Prof."
    if "PROF" in raw_upper:
        return "Prof. Dr."
    if "DOÇ" in raw_upper or "DOC" in raw_upper:
        return "Doç. Dr."
    if "ÖĞR" in raw_upper or "OGR" in raw_upper:
        return "Dr. Öğr. Üyesi"
    if "DR." in raw_upper or "DR " in raw_upper:
        return "Dr."
    if "UZM" in raw_upper:
        return "Uzm."
    if "ARŞ" in raw_upper or "ARS" in raw_upper:
        return "Arş. Gör."
    return ""


def parse_zaman(metin):
    match = re.search(r"(\d{2}\.\d{2}\.\d{4}).*?(\d{2}[:.]\d{2})", str(metin))
    if match:
        date_str = match.group(1)
        time_str = match.group(2).replace(".", ":")
        try:
            return datetime.datetime.strptime(f"{date_str} {time_str}", "%d.%m.%Y %H:%M")
        except Exception:
            pass
    return datetime.datetime(2099, 1, 1)


def gun_key(gun_ve_saat: str) -> str:
    text = str(gun_ve_saat).strip()
    if " | " in text:
        return text.split(" | ")[0].strip()
    m = re.search(r"\d{2}\.\d{2}\.\d{4}\s+\S+", text)
    if m:
        return m.group(0)
    return text


def clean_salon(sal):
    if isinstance(sal, pd.Timestamp) or hasattr(sal, "month"):
        if sal.month == 8 and sal.day == 30:
            return "30 Ağustos Salonu"
    n = str(sal).upper().strip()
    if "30 A" in n or "08-30" in n:
        return "30 Ağustos Salonu"
    if "MALAZGIRT" in n or "MALAZGİRT" in n:
        return "Malazgirt Salonu"
    if "ELIF" in n or "ELİF" in n:
        return "Doç. Dr. Elif Kaya Salonu"
    if "AHMET" in n or "EKIZER" in n or "EKİZER" in n:
        return "Dr. Ahmet Ekizer Salonu"
    if "TEAMS 1" in n or "ZOOM 1" in n or "ODA 1" in n:
        return "Teams Oda 1"
    if "TEAMS 2" in n or "ZOOM 2" in n or "ODA 2" in n:
        return "Teams Oda 2"
    if "TEAMS 3" in n or "ZOOM 3" in n or "ODA 3" in n:
        return "Teams Oda 3"
    return str(sal).strip()


def find_col(df: pd.DataFrame, candidates: List[str], contains: Optional[List[str]] = None) -> Optional[str]:
    if df is None or df.empty:
        return None
    normalized = {temiz_metin(str(c)): c for c in df.columns}
    for cand in candidates:
        key = temiz_metin(cand)
        if key in normalized:
            return normalized[key]
    if contains:
        for col in df.columns:
            col_low = str(col).lower()
            if all(x.lower() in col_low for x in contains):
                return col
    return None


def get_kisi_renk_emoji(k_data):
    odeme = str(k_data.get("payment", "")).lower()
    gorev = str(k_data.get("role", "")).lower()
    bildiri_sayisi = len(k_data.get("submissions", []))
    if "muaf" in odeme or "görevli" in gorev or "davetli" in gorev:
        return "🟣"
    if "indirim" in odeme or "öğrenci" in odeme or "ogrenci" in odeme:
        return "🟠"
    if "bekleniyor" in odeme:
        return "🟡"
    if "yok" in odeme:
        return "⚪"
    if bildiri_sayisi == 0:
        return "🔵"
    return "🟢"


def get_bildiri_renk_emoji(b_data):
    return "🟢" if b_data.get("payment") == "Evet" else "🟡"


# =========================
# SUPABASE
# =========================

def get_secret(*names):
    for name in names:
        try:
            value = st.secrets.get(name)
            if value:
                return value
        except Exception:
            pass
    return None


@st.cache_resource(show_spinner=False)
def get_supabase():
    url = get_secret("SUPABASE_URL", "supabase_url")
    key = get_secret("SUPABASE_SERVICE_ROLE_KEY", "SUPABASE_KEY", "supabase_key")
    if not url or not key:
        return None
    return create_client(url, key)


def sb():
    client = get_supabase()
    if client is None:
        st.error("Supabase ayarları eksik. Streamlit Secrets içine SUPABASE_URL ve SUPABASE_SERVICE_ROLE_KEY ekle.")
        st.stop()
    return client


def fetch_all(table: str, order: Optional[str] = None) -> List[dict]:
    client = sb()
    rows = []
    start = 0
    step = 1000
    while True:
        query = client.table(table).select("*")
        if order:
            query = query.order(order)
        result = query.range(start, start + step - 1).execute()
        chunk = result.data or []
        rows.extend(chunk)
        if len(chunk) < step:
            break
        start += step
    return rows


def df_from_table(table: str, order: Optional[str] = None) -> pd.DataFrame:
    return pd.DataFrame(fetch_all(table, order=order))


def insert_batches(table: str, records: List[dict], batch_size: int = 500):
    if not records:
        return
    client = sb()
    for i in range(0, len(records), batch_size):
        client.table(table).insert(records[i : i + batch_size]).execute()


def update_by_id(table: str, row_id: int, values: Dict):
    sb().table(table).update(values).eq("id", row_id).execute()


def delete_all_data():
    client = sb()
    for table in ["submission_authors", "special_events", "moderators", "submissions", "participants"]:
        client.table(table).delete().neq("id", 0).execute()


# =========================
# EXCEL ICERI AKTARMA
# =========================

def read_xls(uploaded_file):
    if uploaded_file is None:
        return None
    return pd.ExcelFile(io.BytesIO(uploaded_file.getvalue()))


def read_sheet(xls, sheet_name, default_index=None):
    if xls is None:
        return pd.DataFrame()
    try:
        if default_index is not None:
            df = pd.read_excel(xls, sheet_name=default_index, dtype=str).fillna("")
        else:
            df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str).fillna("")
        df.columns = [str(c).strip() for c in df.columns]
        df["_row"] = list(range(2, len(df) + 2))
        return df
    except Exception:
        return pd.DataFrame()


def merge_participant(participants: Dict[str, dict], name, title="", physical=False, role="Yok", payment="Ödeme Bekleniyor"):
    norm = temiz_isim(str(name))
    if not norm:
        return
    if norm not in participants:
        participants[norm] = {
            "norm_name": norm,
            "original_name": turkce_buyuk(str(name)),
            "title": title,
            "email": "",
            "phone": "",
            "institution": "",
            "role": role,
            "payment": payment,
            "attendance_type": "Fiziki" if physical else "Belirtilmedi",
            "intent": "Hayır",
            "days": {"6": "", "7": "", "8": ""},
            "events": "",
        }
    if title:
        participants[norm]["title"] = title
    if role and role != "Yok":
        participants[norm]["role"] = role
    if payment and payment != "Ödeme Bekleniyor":
        participants[norm]["payment"] = payment
    if physical:
        participants[norm]["attendance_type"] = "Fiziki"


def build_import_payload(katilimci_xls, program_xls):
    program_plan = read_sheet(program_xls, None, default_index=PROGRAM_PLAN_INDEX)
    program_bildiriler = read_sheet(program_xls, PROGRAM_BILDIRILER_SHEET)
    special_events_df = read_sheet(program_xls, PROGRAM_OZEL_SHEET)
    moderators_df = read_sheet(program_xls, PROGRAM_MOD_SHEET)

    bildiri_liste = read_sheet(katilimci_xls, "kabul edilen bildiriler")
    odeme_df = read_sheet(katilimci_xls, "ödeme durumu")
    bilgi_df = read_sheet(katilimci_xls, "katılımcı bilgileri")
    anket_df = read_sheet(katilimci_xls, "katılımcıların ankete cevapları")
    danisma_df = read_sheet(katilimci_xls, "Kongre Bilimsel Danışma Kurulu")
    duzenleme_df = read_sheet(katilimci_xls, "Kongre Düzenleme Kurulu")
    detay_df = read_sheet(katilimci_xls, "bildirilerin detayları")
    yaka_df = read_sheet(katilimci_xls, "yaka kartı ek liste")

    if program_bildiriler.empty:
        program_bildiriler = bildiri_liste.copy()

    participants = {}
    submissions = {}
    author_links = []

    detay_map = {}
    for _, row in detay_df.iterrows():
        b_adi = ""
        for col in detay_df.columns:
            if "bildiri" in str(col).lower() and ("ad" in str(col).lower() or "ism" in str(col).lower()):
                b_adi = safe_str(row.get(col))
                break
        if b_adi:
            b_key = temiz_metin(b_adi)
            detay_map[b_key] = {
                "topic": safe_str(row.get("Konu")) or "-",
                "boutique": "Evet" if "evet" in safe_str(row.get("Butik Bildiri")).lower() else "",
            }

    for sheet_df, role in [
        (danisma_df, "Görevli - Bilimsel Danışma Kurulu"),
        (duzenleme_df, "Görevli - Kongre Düzenleme Kurulu"),
    ]:
        for _, r in sheet_df.iterrows():
            isim = safe_str(r.iloc[0]) if len(r) else ""
            merge_participant(
                participants,
                isim,
                title=unvan_bul(r, sheet_df.columns, isim),
                role=role,
                payment="Görevli/Davetli (Muaf)",
                physical=True,
            )

    for _, r in yaka_df.iterrows():
        if len(r) >= 2:
            merge_participant(
                participants,
                safe_str(r.iloc[1]),
                title=safe_str(r.iloc[0]),
                role=safe_str(r.iloc[2]) if len(r) > 2 else "Davetli Katılımcı",
                payment="Görevli/Davetli (Muaf)",
                physical=True,
            )

    for _, r in odeme_df.iterrows():
        if len(r) >= 2:
            merge_participant(participants, safe_str(r.iloc[0]), payment=safe_str(r.iloc[1]))

    for _, r in bilgi_df.iterrows():
        ad_soy = safe_str(r.get("Adı Soyadı", r.iloc[0] if len(r) else ""))
        norm = temiz_isim(ad_soy)
        if norm in participants:
            for c in bilgi_df.columns:
                cl = str(c).lower()
                val = safe_str(r.get(c))
                if not val:
                    continue
                if "mail" in cl or "posta" in cl:
                    participants[norm]["email"] = val
                elif "telefon" in cl or "gsm" in cl:
                    participants[norm]["phone"] = val
                elif "kurum" in cl or "üniversite" in cl:
                    participants[norm]["institution"] = val

    for _, r in anket_df.iterrows():
        ad = safe_str(r.get("Adı Soyadı", r.iloc[0] if len(r) else ""))
        norm = temiz_isim(ad)
        if not norm:
            continue
        merge_participant(participants, ad, title=unvan_bul(r, anket_df.columns, ad))
        for c in anket_df.columns:
            cl = str(c).lower()
            val = safe_str(r.get(c))
            if not val:
                continue
            if "nasıl katılım" in cl:
                participants[norm]["attendance_type"] = val
            elif "6 mayıs" in cl:
                participants[norm]["days"]["6"] = val
            elif "7 mayıs" in cl:
                participants[norm]["days"]["7"] = val
            elif "8 mayıs" in cl:
                participants[norm]["days"]["8"] = val
            elif "etkinlik" in cl:
                participants[norm]["events"] += ", " + val
            elif "sunacaksanız" in cl and "katılımcı" not in val.lower():
                participants[norm]["intent"] = "Evet"

    schedule_map = {}
    if not program_plan.empty:
        title_col = find_col(program_plan, ["Bildiri_Adi", "Bildiri Adı", "Bildiri Ismi"], contains=["bildiri"])
        for _, row in program_plan.iterrows():
            title = safe_str(row.get(title_col)) if title_col else ""
            norm_title = temiz_metin(title)
            if not norm_title:
                continue
            schedule_map[norm_title] = {
                "title": title,
                "presentation_type": safe_str(row.get("Sunum_Tipi")),
                "session_time": safe_str(row.get("Gun_ve_Saat")),
                "hall": safe_str(row.get("Salon")),
                "session_id": safe_str(row.get("Oturum_ID")),
                "source_row": int(row.get("_row", 0) or 0),
            }

    b_col = find_col(program_bildiriler, ["Bildiri Ismi", "Bildiri_Adi", "Bildiri Adı"], contains=["bildiri"])
    presenter_col = find_col(program_bildiriler, ["Sunan Yazar", "sunucu", "Sunucu"], contains=["sunan"])
    topic_col = find_col(program_bildiriler, ["Konu"], contains=["konu"])

    for _, row in program_bildiriler.iterrows():
        title = safe_str(row.get(b_col)) if b_col else ""
        norm_title = temiz_metin(title)
        if not norm_title:
            continue
        presenter_norm = temiz_isim(safe_str(row.get(presenter_col))) if presenter_col else ""
        authors = []
        for i in range(1, 9):
            author_norm = temiz_isim(safe_str(row.get(f"Yazar {i}")))
            if author_norm and author_norm not in authors:
                authors.append(author_norm)
                merge_participant(participants, author_norm)
        if presenter_norm and presenter_norm not in authors:
            authors.append(presenter_norm)
            merge_participant(participants, presenter_norm)

        sched = schedule_map.get(norm_title, {})
        submissions[norm_title] = {
            "norm_title": norm_title,
            "title": title,
            "topic": detay_map.get(norm_title, {}).get("topic") or (safe_str(row.get(topic_col)) if topic_col else "-") or "-",
            "boutique": detay_map.get(norm_title, {}).get("boutique", ""),
            "accepted": "Evet",
            "presenter_norm_name": presenter_norm,
            "payment": "Hayır",
            "presentation_type": sched.get("presentation_type", ""),
            "session_time": sched.get("session_time", ""),
            "hall": sched.get("hall", ""),
            "session_id": sched.get("session_id", ""),
            "source_row": sched.get("source_row"),
            "_authors": authors,
        }

    for norm_title, sched in schedule_map.items():
        if norm_title not in submissions:
            submissions[norm_title] = {
                "norm_title": norm_title,
                "title": sched["title"],
                "topic": detay_map.get(norm_title, {}).get("topic", "-"),
                "boutique": detay_map.get(norm_title, {}).get("boutique", ""),
                "accepted": "Evet",
                "presenter_norm_name": "",
                "payment": "Hayır",
                "presentation_type": sched.get("presentation_type", ""),
                "session_time": sched.get("session_time", ""),
                "hall": sched.get("hall", ""),
                "session_id": sched.get("session_id", ""),
                "source_row": sched.get("source_row"),
                "_authors": [],
            }

    for sub in submissions.values():
        paid = False
        for author_norm in sub["_authors"]:
            payment = participants.get(author_norm, {}).get("payment", "Ödeme Bekleniyor")
            if "Bekleniyor" not in payment and "Yok" not in payment:
                paid = True
        sub["payment"] = "Evet" if paid else "Hayır"

    for norm, participant in participants.items():
        if participant["payment"] == "Ödeme Bekleniyor":
            has_submission = any(norm in sub["_authors"] for sub in submissions.values())
            if not has_submission:
                participant["payment"] = "Bildirisi Yok / Kayıt Yok"

    participant_records = list(participants.values())
    submission_records = []
    authors_by_submission = {}
    for norm_title, sub in submissions.items():
        authors_by_submission[norm_title] = sub.pop("_authors")
        submission_records.append(sub)

    moderator_records = []
    for _, r in moderators_df.iterrows():
        if not any(safe_str(x) for x in r.values):
            continue
        online_raw = safe_str(r.get("Online")).lower()
        moderator_records.append(
            {
                "name": safe_str(r.get("unvan_ad_soyad")),
                "institution": safe_str(r.get("kurum")),
                "duty": safe_str(r.get("Gorev")) or "Moderator",
                "session_time": safe_str(r.get("Gun_ve_Saat")),
                "hall": safe_str(r.get("Salon")),
                "session_id": safe_str(r.get("Oturum_ID")),
                "online": online_raw in ["evet", "e", "yes", "1", "true", "online"],
                "source_row": int(r.get("_row", 0) or 0),
            }
        )

    special_event_records = []
    for _, r in special_events_df.iterrows():
        if not any(safe_str(x) for x in r.values):
            continue
        special_event_records.append(
            {
                "event_order": safe_str(r.get("oturum_sirasi")),
                "hall": safe_str(r.get("salon")),
                "datetime_text": safe_str(r.get("tarih_saat")),
                "main_title": safe_str(r.get("ana_baslik")),
                "subtitle": safe_str(r.get("alt_baslik")),
                "left_text": safe_str(r.get("sol_metin")),
                "right_text": safe_str(r.get("sag_metin")),
            }
        )

    return participant_records, submission_records, authors_by_submission, moderator_records, special_event_records


def import_excel_to_supabase(katilimci_file, program_file):
    katilimci_xls = read_xls(katilimci_file)
    program_xls = read_xls(program_file) if program_file is not None else katilimci_xls
    (
        participant_records,
        submission_records,
        authors_by_submission,
        moderator_records,
        special_event_records,
    ) = build_import_payload(katilimci_xls, program_xls)

    delete_all_data()
    insert_batches("participants", participant_records)
    insert_batches("submissions", submission_records)
    insert_batches("moderators", moderator_records)
    insert_batches("special_events", special_event_records)

    participants_df = df_from_table("participants")
    submissions_df = df_from_table("submissions")
    p_id = {row["norm_name"]: row["id"] for _, row in participants_df.iterrows()}
    s_id = {row["norm_title"]: row["id"] for _, row in submissions_df.iterrows()}

    link_records = []
    for norm_title, authors in authors_by_submission.items():
        if norm_title not in s_id:
            continue
        for order, author_norm in enumerate(authors, start=1):
            display_name = participants_df.loc[participants_df["norm_name"] == author_norm, "original_name"]
            link_records.append(
                {
                    "submission_id": int(s_id[norm_title]),
                    "participant_id": int(p_id[author_norm]) if author_norm in p_id else None,
                    "author_order": order,
                    "norm_name": author_norm,
                    "display_name": display_name.iloc[0] if not display_name.empty else turkce_buyuk(author_norm),
                }
            )
    insert_batches("submission_authors", link_records)

    return {
        "participants": len(participant_records),
        "submissions": len(submission_records),
        "authors": len(link_records),
        "moderators": len(moderator_records),
        "special_events": len(special_event_records),
    }


# =========================
# DB SNAPSHOT VE SORGULAR
# =========================

@st.cache_data(ttl=30, show_spinner=False)
def load_snapshot():
    return {
        "participants": df_from_table("participants", order="original_name"),
        "submissions": df_from_table("submissions", order="session_time"),
        "authors": df_from_table("submission_authors", order="author_order"),
        "moderators": df_from_table("moderators", order="session_time"),
        "special_events": df_from_table("special_events", order="event_order"),
    }


def build_lookup(snapshot):
    participants_df = snapshot["participants"]
    submissions_df = snapshot["submissions"]
    authors_df = snapshot["authors"]
    katilimcilar = {}
    bildiriler = {}

    for _, p in participants_df.iterrows():
        katilimcilar[p["norm_name"]] = p.to_dict()
        katilimcilar[p["norm_name"]]["submissions"] = []

    for _, s in submissions_df.iterrows():
        sub_authors = authors_df[authors_df["submission_id"] == s["id"]].sort_values("author_order")
        authors = list(sub_authors["norm_name"]) if not sub_authors.empty else []
        d = s.to_dict()
        d["authors"] = authors
        bildiriler[s["norm_title"]] = d
        for author_norm in authors:
            if author_norm in katilimcilar:
                katilimcilar[author_norm]["submissions"].append(s["title"])
    return bildiriler, katilimcilar


def df_master_from_snapshot(snapshot):
    submissions_df = snapshot["submissions"]
    authors_df = snapshot["authors"]
    rows = []
    for _, s in submissions_df.iterrows():
        sub_authors = authors_df[authors_df["submission_id"] == s["id"]].sort_values("author_order")
        yazarlar = " - ".join([safe_str(x) for x in sub_authors["display_name"]]) if not sub_authors.empty else ""
        presenter_norm = safe_str(s.get("presenter_norm_name"))
        presenter_display = ""
        if presenter_norm and not sub_authors.empty:
            match = sub_authors[sub_authors["norm_name"] == presenter_norm]
            if not match.empty:
                presenter_display = safe_str(match.iloc[0]["display_name"])
        if not presenter_display:
            presenter_display = presenter_norm
        rows.append(
            {
                "id": int(s["id"]),
                "Bildiri_Adi": safe_str(s.get("title")),
                "Gun_ve_Saat": safe_str(s.get("session_time")),
                "Salon": safe_str(s.get("hall")),
                "Oturum_ID": safe_str(s.get("session_id")),
                "Sunum_Tipi": safe_str(s.get("presentation_type")),
                "yazarlar": yazarlar,
                "sunucu": presenter_display,
                "konu": safe_str(s.get("topic")) or "-",
            }
        )
    return pd.DataFrame(rows)


def moderators_df_for_matbaa(snapshot):
    df = snapshot["moderators"].copy()
    if df.empty:
        return pd.DataFrame()
    return pd.DataFrame(
        {
            "unvan_ad_soyad": df.get("name", ""),
            "kurum": df.get("institution", ""),
            "Gorev": df.get("duty", ""),
            "Gun_ve_Saat": df.get("session_time", ""),
            "Salon": df.get("hall", ""),
            "Oturum_ID": df.get("session_id", ""),
            "Online": df.get("online", False).apply(lambda x: "Evet" if bool(x) else "Hayır"),
        }
    )


def special_events_df_for_matbaa(snapshot):
    df = snapshot["special_events"].copy()
    if df.empty:
        return pd.DataFrame()
    return pd.DataFrame(
        {
            "oturum_sirasi": df.get("event_order", ""),
            "salon": df.get("hall", ""),
            "tarih_saat": df.get("datetime_text", ""),
            "ana_baslik": df.get("main_title", ""),
            "alt_baslik": df.get("subtitle", ""),
            "sol_metin": df.get("left_text", ""),
            "sag_metin": df.get("right_text", ""),
        }
    )


def program_info(sub):
    parts = []
    if safe_str(sub.get("presentation_type")):
        parts.append(safe_str(sub.get("presentation_type")))
    if safe_str(sub.get("session_time")):
        parts.append(safe_str(sub.get("session_time")))
    if safe_str(sub.get("hall")):
        parts.append(safe_str(sub.get("hall")))
    if safe_str(sub.get("session_id")):
        parts.append(safe_str(sub.get("session_id")))
    return " | ".join(parts) if parts else "Henüz programda yeri belli değil."


# =========================
# MATBAA MOTORU
# =========================

def set_cell_bg(cell, hex_color):
    hex_color = hex_color.replace("#", "")
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def add_merged_row(table, text, bg_color, text_color=(0, 0, 0), bold=False, align=None):
    row = table.add_row()
    cell = row.cells[0]
    cell.merge(row.cells[1])
    cell.text = text
    set_cell_bg(cell, bg_color)
    for paragraph in cell.paragraphs:
        if align == "center":
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in paragraph.runs:
            run.font.color.rgb = RGBColor(*text_color)
            run.font.bold = bold
    return row


def build_program_dict(df_bildiriler):
    prog = {}
    if df_bildiriler.empty:
        return prog
    for _, r in df_bildiriler.iterrows():
        b_adi = safe_str(r.get("Bildiri_Adi"))
        kz = safe_str(r.get("Gun_ve_Saat")).replace("Persembe", "Perşembe")
        if not b_adi or not kz:
            continue
        ks = clean_salon(r.get("Salon"))
        key = (kz, ks)
        prog.setdefault(key, []).append(
            {
                "b": b_adi,
                "y": safe_str(r.get("yazarlar")),
                "s": safe_str(r.get("sunucu")),
                "k": safe_str(r.get("konu")) or "-",
                "sid": safe_str(r.get("Oturum_ID")),
            }
        )
    return prog


def moderator_atamalari_olustur(prog, df_mod, tip):
    mod_atamalari = {
        k: {"mod": "Daha Sonra İlan Edilecektir", "deg": "Daha Sonra İlan Edilecektir"} for k in prog.keys()
    }
    if df_mod is None or df_mod.empty:
        return mod_atamalari

    musait_oturumlar_mod = list(prog.keys())
    musait_oturumlar_deg = list(prog.keys())
    serbest_mods, serbest_degs = [], []

    for _, r in df_mod.iterrows():
        is_online_mod = safe_str(r.get("Online")).lower() in ["evet", "e", "yes", "1", "var", "online", "true"]
        if tip == "ONLINE" and not is_online_mod:
            continue
        if tip == "YUZYUZE" and is_online_mod:
            continue
        m_gun_saat = safe_str(r.get("Gun_ve_Saat")).replace("Persembe", "Perşembe")
        m_salon_ham = safe_str(r.get("Salon"))
        m_salon = clean_salon(m_salon_ham) if m_salon_ham != "-" else "-"
        m_oturum = safe_str(r.get("Oturum_ID")) or "-"
        gorev_sutunu = safe_str(r.get("Gorev")).lower()
        kurum = safe_str(r.get("kurum"))
        mod_metni = f"{safe_str(r.get('unvan_ad_soyad'))} ({kurum})" if kurum else safe_str(r.get("unvan_ad_soyad"))
        is_deg = "deg" in gorev_sutunu
        role_key = "deg" if is_deg else "mod"
        hedef_musait = musait_oturumlar_deg if is_deg else musait_oturumlar_mod
        hedef_serbest = serbest_degs if is_deg else serbest_mods

        atandi_mi = False
        if m_gun_saat and m_gun_saat != "-" and m_salon and m_salon != "-":
            hedef_tuple = (m_gun_saat, m_salon)
            if hedef_tuple in hedef_musait:
                mod_atamalari[hedef_tuple][role_key] = mod_metni
                hedef_musait.remove(hedef_tuple)
                atandi_mi = True

        if not atandi_mi and (m_gun_saat != "-" or m_salon != "-" or m_oturum != "-"):
            for sess in list(hedef_musait):
                oturum_id_metni = prog[sess][0]["sid"] if prog[sess] else "-"
                if (
                    (m_gun_saat == "-" or m_gun_saat in sess[0])
                    and (m_salon == "-" or m_salon == sess[1])
                    and (m_oturum == "-" or m_oturum == oturum_id_metni)
                ):
                    mod_atamalari[sess][role_key] = mod_metni
                    hedef_musait.remove(sess)
                    atandi_mi = True
                    break
        if not atandi_mi and mod_metni:
            hedef_serbest.append(mod_metni)

    for mod_metni in serbest_mods:
        if musait_oturumlar_mod:
            mod_atamalari[musait_oturumlar_mod.pop(0)]["mod"] = mod_metni
    for deg_metni in serbest_degs:
        if musait_oturumlar_deg:
            mod_atamalari[musait_oturumlar_deg.pop(0)]["deg"] = deg_metni
    return mod_atamalari


def ozel_etkinlikleri_hazirla(df_ozel):
    soft_palette = ["#DDEBF7", "#FCE4D6", "#E2EFDA", "#FFF2CC", "#E6E6FA", "#F2F2F2"]
    gosterilecek = []
    if df_ozel is None or df_ozel.empty:
        return gosterilecek
    if not {"oturum_sirasi", "salon"}.issubset(set(df_ozel.columns)):
        return gosterilecek
    for color_index, ((sira, sal), grup) in enumerate(df_ozel.groupby(["oturum_sirasi", "salon"], sort=False)):
        ks = clean_salon(sal)
        renk_temasi = soft_palette[color_index % len(soft_palette)]
        ts = safe_str(grup.iloc[0].get("tarih_saat")).replace("Persembe", "Perşembe")
        gosterilecek.append((sira, ks, grup, renk_temasi, ts))
    return gosterilecek


def gun_istatistikleri_olustur(prog):
    gun_istatistikleri = {}
    for (oturum_zaman, _sal), bilds in prog.items():
        t = oturum_zaman.split(" | ")[0] if " | " in oturum_zaman else oturum_zaman
        gun_istatistikleri.setdefault(t, {"oturum_sayisi": 0, "bildiri_sayisi": 0})
        gun_istatistikleri[t]["oturum_sayisi"] += 1
        gun_istatistikleri[t]["bildiri_sayisi"] += len(bilds)
    return gun_istatistikleri


def word_bas(df_bildiriler, df_ozel, df_mod, tip):
    prog = build_program_dict(df_bildiriler)
    gosterilecek_ozel_etkinlikler = ozel_etkinlikleri_hazirla(df_ozel)
    mod_atamalari = moderator_atamalari_olustur(prog, df_mod, tip)
    gun_istatistikleri = gun_istatistikleri_olustur(prog)

    doc = Document()
    section = doc.sections[-1]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width

    ana_baslik = "11. IHMC FİZİKİ / FACE TO FACE PROGRAM" if tip == "YUZYUZE" else "11. IHMC ONLINE (DIGITAL) PROGRAM"
    h = doc.add_heading(ana_baslik, level=1)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()

    if gosterilecek_ozel_etkinlikler:
        t_ozel_baslik = doc.add_table(rows=0, cols=2)
        t_ozel_baslik.style = "Table Grid"
        add_merged_row(t_ozel_baslik, "ÖZEL ETKİNLİKLER", "000000", text_color=(255, 255, 255), bold=True, align="center")
        doc.add_paragraph()
        for _sira, ks, grup, renk_temasi, ts in gosterilecek_ozel_etkinlikler:
            t_etk = doc.add_table(rows=0, cols=2)
            t_etk.style = "Table Grid"
            add_merged_row(t_etk, f"{ts} | HALL: {ks}", renk_temasi, bold=True, align="center")
            for _, r in grup.iterrows():
                m = "\n".join(
                    [
                        safe_str(r.get(c))
                        for c in ["ana_baslik", "alt_baslik", "sol_metin", "sag_metin"]
                        if safe_str(r.get(c)) not in ["-", "nan", ""]
                    ]
                )
                add_merged_row(t_etk, m, renk_temasi, align="center")
            doc.add_paragraph()

    t_bilimsel_baslik = doc.add_table(rows=0, cols=2)
    t_bilimsel_baslik.style = "Table Grid"
    add_merged_row(t_bilimsel_baslik, "--- BİLİMSEL BİLDİRİ PROGRAMI ---", "000000", text_color=(255, 255, 255), bold=True, align="center")
    doc.add_paragraph()

    mevcut_islenen_gun = ""
    for (oturum_zaman, sal), bilds in sorted(prog.items(), key=lambda x: parse_zaman(x[0][0])):
        parcalar = oturum_zaman.split(" | ")
        t = parcalar[0] if len(parcalar) > 0 else oturum_zaman
        sa = parcalar[1] if len(parcalar) > 1 else "-"
        sn = bilds[0]["sid"] if bilds else "-"

        if t != mevcut_islenen_gun:
            mevcut_islenen_gun = t
            t_gun = doc.add_table(rows=0, cols=2)
            t_gun.style = "Table Grid"
            add_merged_row(t_gun, f">>> {t.upper()} BİLİMSEL PROGRAMI <<<", "B4C6E7", bold=True, align="center")
            add_merged_row(
                t_gun,
                "Her oturumdaki moderatör oturumu yönetecek ve Oturum Değerlendirici ile birbirinden bağımsız olarak EN İYİ BİLDİRİ (BEST PAPER) ÖDÜLLERi için bildiri sunumlarını değerlendireceklerdir.",
                "FFF2CC",
                align="center",
            )
            o_sayi = gun_istatistikleri.get(t, {}).get("oturum_sayisi", 0)
            b_sayi = gun_istatistikleri.get(t, {}).get("bildiri_sayisi", 0)
            add_merged_row(
                t_gun,
                f"Bugün toplam {o_sayi} adet bildiri sunum oturumu gerçekleşecek ve {b_sayi} adet bildiri sunulacaktır.",
                "E2EFDA",
                bold=True,
                text_color=(55, 86, 35),
                align="center",
            )
            doc.add_paragraph()

        mod_isim = mod_atamalari[(oturum_zaman, sal)]["mod"]
        deg_isim = mod_atamalari[(oturum_zaman, sal)]["deg"]
        oturum_metni = f"HALL: {sal}\n{sn}\n\n{bilds[0]['k'].upper()}\nModerator: {mod_isim}\nOturum Değerlendirici: {deg_isim}"

        t_oturum = doc.add_table(rows=0, cols=2)
        t_oturum.style = "Table Grid"
        add_merged_row(t_oturum, f"{t} | {sa}", "FFD966", bold=True)
        add_merged_row(t_oturum, oturum_metni, "FFE699", bold=True)
        for i, b in enumerate(bilds):
            bg_color = "F2F2F2" if i % 2 == 1 else "FFFFFF"
            row = t_oturum.add_row()
            yazarlar_metni = b["y"]
            sunucu = b["s"]
            cell_yazar = row.cells[0]
            cell_yazar.text = ""
            set_cell_bg(cell_yazar, bg_color)
            p = cell_yazar.paragraphs[0]
            if sunucu and sunucu not in ["-", "nan", ""] and sunucu in yazarlar_metni:
                parts = yazarlar_metni.split(sunucu, 1)
                if parts[0]:
                    p.add_run(parts[0])
                run_s = p.add_run(sunucu)
                run_s.underline = True
                run_s.bold = True
                if len(parts) > 1 and parts[1]:
                    p.add_run(parts[1])
            else:
                p.add_run(yazarlar_metni)
            row.cells[1].text = b["b"]
            set_cell_bg(row.cells[1], bg_color)
        doc.add_paragraph()

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output


def excel_bas(df_bildiriler, df_ozel, df_mod, tip):
    prog = build_program_dict(df_bildiriler)
    tum_zamanlar = {k[0] for k in prog.keys()}
    tum_salonlar = {k[1] for k in prog.keys()}
    gosterilecek_ozel_etkinlikler = []
    soft_palette = ["#DDEBF7", "#FCE4D6", "#E2EFDA", "#FFF2CC", "#E6E6FA", "#F2F2F2"]
    if df_ozel is not None and not df_ozel.empty and {"oturum_sirasi", "salon"}.issubset(set(df_ozel.columns)):
        for color_index, ((sira, sal), grup) in enumerate(df_ozel.groupby(["oturum_sirasi", "salon"], sort=False)):
            ks = clean_salon(sal)
            tum_salonlar.add(ks)
            renk_temasi = soft_palette[color_index % len(soft_palette)]
            ts = safe_str(grup.iloc[0].get("tarih_saat")).replace("Persembe", "Perşembe")
            t_event = "07.05.2026 Perşembe" if "07" in ts else ("08.05.2026 Cuma" if "08" in ts else "Bilinmeyen Gun")
            m = re.search(r"(\d{2}[.:]\d{2})\s*-\s*(\d{2}[.:]\d{2})", ts)
            sa_event = f"{m.group(1).replace('.', ':')}-{m.group(2).replace('.', ':')}" if m else ts
            zaman_key = f"{t_event} | {sa_event}"
            tum_zamanlar.add(zaman_key)
            gosterilecek_ozel_etkinlikler.append((sira, ks, grup, renk_temasi, zaman_key))

    mod_atamalari = moderator_atamalari_olustur(prog, df_mod, tip)
    gun_istatistikleri = gun_istatistikleri_olustur(prog)

    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {"in_memory": True})

    map_data = {}
    for (kz, ks), bilds in prog.items():
        map_data.setdefault((kz, ks), {"session_count": 0, "event_name": None, "event_color": None})
        map_data[(kz, ks)]["session_count"] = len(bilds)
    for _sira, ks, grup, renk, zaman_key in gosterilecek_ozel_etkinlikler:
        ana_baslik = safe_str(grup.iloc[0].get("ana_baslik")) or "Etkinlik"
        map_data.setdefault((zaman_key, ks), {"session_count": 0, "event_name": None, "event_color": None})
        map_data[(zaman_key, ks)]["event_name"] = ana_baslik
        map_data[(zaman_key, ks)]["event_color"] = renk

    list_zamanlar = sorted(list(tum_zamanlar), key=parse_zaman)
    list_salonlar = sorted(list(tum_salonlar))

    ws_harita = workbook.add_worksheet("Oturum_Haritasi")
    f_map_baslik = workbook.add_format({"bold": True, "bg_color": "#4472C4", "font_color": "white", "border": 1, "align": "center", "valign": "vcenter"})
    f_map_saat = workbook.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1, "align": "left"})
    f_map_bos = workbook.add_format({"bg_color": "#F2F2F2", "font_color": "#A6A6A6", "border": 1, "align": "center", "valign": "vcenter"})
    f_map_dolu = workbook.add_format({"bold": True, "bg_color": "#E2EFDA", "font_color": "#375623", "border": 1, "align": "center", "valign": "vcenter"})
    f_map_cakisma = workbook.add_format({"bold": True, "bg_color": "#FF0000", "font_color": "white", "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True})
    dinamik_renk_cache = {}

    ws_harita.set_column(0, 0, 35)
    ws_harita.write(0, 0, "SAAT / OTURUM", f_map_baslik)
    for col_idx, sal in enumerate(list_salonlar, 1):
        ws_harita.set_column(col_idx, col_idx, 25)
        ws_harita.write(0, col_idx, sal, f_map_baslik)
    for row_idx, z_key in enumerate(list_zamanlar, 1):
        ws_harita.write(row_idx, 0, z_key, f_map_saat)
        for col_idx, sal in enumerate(list_salonlar, 1):
            hucre = map_data.get((z_key, sal), {"session_count": 0, "event_name": None})
            s_count = hucre["session_count"]
            e_name = hucre["event_name"]
            e_color = hucre.get("event_color", "#FFD966")
            if s_count > 0 and e_name:
                ws_harita.write(row_idx, col_idx, f"ÇAKIŞMA!\n[{s_count} Bildiri] VE [{e_name}]", f_map_cakisma)
                ws_harita.set_row(row_idx, 45)
            elif s_count > 0:
                ws_harita.write(row_idx, col_idx, f"{s_count} Bildiri Var", f_map_dolu)
            elif e_name:
                if e_color not in dinamik_renk_cache:
                    dinamik_renk_cache[e_color] = workbook.add_format({"bold": True, "bg_color": e_color, "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True})
                ws_harita.write(row_idx, col_idx, f"[ETKİNLİK]\n{e_name}", dinamik_renk_cache[e_color])
                ws_harita.set_row(row_idx, 40)
            else:
                ws_harita.write(row_idx, col_idx, "BOŞ", f_map_bos)

    worksheet = workbook.add_worksheet("Kesin_Program")
    worksheet.set_column("A:A", 65)
    worksheet.set_column("B:B", 95)
    fmt = {
        "ayirici": workbook.add_format({"bg_color": "#000000", "font_color": "#FFFFFF", "bold": True, "align": "center", "valign": "vcenter"}),
        "baslik_dev": workbook.add_format({"bold": True, "font_size": 14, "align": "center", "valign": "vcenter", "border": 1}),
        "mavi_gun": workbook.add_format({"bold": True, "bg_color": "#B4C6E7", "align": "left", "valign": "vcenter", "border": 1, "text_wrap": True}),
        "odul_sari": workbook.add_format({"bg_color": "#FFF2CC", "align": "left", "valign": "vcenter", "border": 1, "text_wrap": True}),
        "istatistik_yesil": workbook.add_format({"bold": True, "bg_color": "#E2EFDA", "font_color": "#375623", "align": "center", "valign": "vcenter", "border": 1, "text_wrap": True}),
        "saat_baslik": workbook.add_format({"bold": True, "bg_color": "#FFD966", "align": "left", "valign": "vcenter", "border": 1, "text_wrap": True}),
        "salon_baslik": workbook.add_format({"bold": True, "bg_color": "#FFE699", "align": "left", "valign": "vcenter", "border": 1, "text_wrap": True}),
        "bildiri_tek": workbook.add_format({"bg_color": "#FFFFFF", "align": "left", "valign": "vcenter", "border": 1, "text_wrap": True}),
        "bildiri_tek_u": workbook.add_format({"bold": True, "underline": True, "bg_color": "#FFFFFF", "border": 1, "text_wrap": True}),
        "bildiri_tek_u_only": workbook.add_format({"bold": True, "underline": True}),
        "bildiri_cift": workbook.add_format({"bg_color": "#F2F2F2", "align": "left", "valign": "vcenter", "border": 1, "text_wrap": True}),
        "bildiri_cift_u": workbook.add_format({"bold": True, "underline": True, "bg_color": "#F2F2F2", "border": 1, "text_wrap": True}),
        "bildiri_cift_u_only": workbook.add_format({"bold": True, "underline": True}),
    }

    satir_no = 0
    ana_baslik = "11. IHMC FİZİKİ / FACE TO FACE PROGRAM" if tip == "YUZYUZE" else "11. IHMC ONLINE (DIGITAL) PROGRAM"
    worksheet.merge_range(satir_no, 0, satir_no, 1, ana_baslik, fmt["baslik_dev"])
    worksheet.set_row(satir_no, 40)
    satir_no += 2

    if gosterilecek_ozel_etkinlikler:
        worksheet.merge_range(satir_no, 0, satir_no, 1, "ÖZEL ETKİNLİKLER", fmt["ayirici"])
        satir_no += 1
        for _sira, ks, grup, renk_temasi, _z_key in gosterilecek_ozel_etkinlikler:
            ts = safe_str(grup.iloc[0].get("tarih_saat")).replace("07.05.2026", "07.05.2026 Perşembe /").replace("08.05.2026", "08.05.2026 Cuma /")
            if renk_temasi not in dinamik_renk_cache:
                dinamik_renk_cache[renk_temasi] = workbook.add_format({"bold": True, "bg_color": renk_temasi, "align": "left", "valign": "vcenter", "border": 1, "text_wrap": True})
            dinamik_f = dinamik_renk_cache[renk_temasi]
            worksheet.merge_range(satir_no, 0, satir_no, 1, f"{ts} | HALL: {ks}", dinamik_f)
            satir_no += 1
            for _, r in grup.iterrows():
                metin = "\n".join([safe_str(r.get(c)) for c in ["ana_baslik", "alt_baslik", "sol_metin", "sag_metin"] if safe_str(r.get(c)) not in ["-", "nan", ""]])
                worksheet.merge_range(satir_no, 0, satir_no, 1, metin, dinamik_f)
                worksheet.set_row(satir_no, 50)
                satir_no += 1
            satir_no += 1

    worksheet.merge_range(satir_no, 0, satir_no, 1, "--- BİLİMSEL BİLDİRİ PROGRAMI ---", fmt["ayirici"])
    satir_no += 2

    mevcut_islenen_gun = ""
    for (oturum_zaman, sal), bilds in sorted(prog.items(), key=lambda x: parse_zaman(x[0][0])):
        parcalar = oturum_zaman.split(" | ")
        t = parcalar[0] if len(parcalar) > 0 else oturum_zaman
        sa = parcalar[1] if len(parcalar) > 1 else "-"
        sn = bilds[0]["sid"] if bilds else "-"
        if t != mevcut_islenen_gun:
            mevcut_islenen_gun = t
            worksheet.merge_range(satir_no, 0, satir_no, 1, f">>> {t.upper()} BİLİMSEL PROGRAMI <<<", fmt["mavi_gun"])
            worksheet.set_row(satir_no, 40)
            satir_no += 1
            odul_notu = "Her oturumdaki moderatör oturumu yönetecek ve Oturum Değerlendirici ile birbirinden bağımsız olarak EN İYİ BİLDİRİ (BEST PAPER) ÖDÜLLERi için bildiri sunumlarını değerlendireceklerdir."
            worksheet.merge_range(satir_no, 0, satir_no, 1, odul_notu, fmt["odul_sari"])
            worksheet.set_row(satir_no, 45)
            satir_no += 1
            o_sayi = gun_istatistikleri.get(t, {}).get("oturum_sayisi", 0)
            b_sayi = gun_istatistikleri.get(t, {}).get("bildiri_sayisi", 0)
            worksheet.merge_range(satir_no, 0, satir_no, 1, f"Bugün toplam {o_sayi} adet bildiri sunum oturumu gerçekleşecek ve {b_sayi} adet bildiri sunulacaktır.", fmt["istatistik_yesil"])
            worksheet.set_row(satir_no, 25)
            satir_no += 2

        mod_isim = mod_atamalari[(oturum_zaman, sal)]["mod"]
        deg_isim = mod_atamalari[(oturum_zaman, sal)]["deg"]
        oturum_metni = f"HALL: {sal}\n{sn}\n\n{bilds[0]['k'].upper()}\nModerator: {mod_isim}\nOturum Değerlendirici: {deg_isim}"
        worksheet.merge_range(satir_no, 0, satir_no, 1, f"{t} | {sa}", fmt["saat_baslik"])
        satir_no += 1
        worksheet.merge_range(satir_no, 0, satir_no, 1, oturum_metni, fmt["salon_baslik"])
        worksheet.set_row(satir_no, 100)
        satir_no += 1
        for i, b in enumerate(bilds):
            yazarlar_metni, sunucu = b["y"], b["s"]
            is_cift = i % 2 == 1
            f_norm = fmt["bildiri_cift"] if is_cift else fmt["bildiri_tek"]
            f_ul_cell = fmt["bildiri_cift_u"] if is_cift else fmt["bildiri_tek_u"]
            f_ul_txt = fmt["bildiri_cift_u_only"] if is_cift else fmt["bildiri_tek_u_only"]
            if sunucu == yazarlar_metni and sunucu not in ["-", "nan", ""]:
                worksheet.write(satir_no, 0, yazarlar_metni, f_ul_cell)
            elif sunucu in yazarlar_metni and sunucu not in ["-", "nan", ""]:
                parts = yazarlar_metni.split(sunucu, 1)
                rich_text = []
                if parts[0]:
                    rich_text.extend([f_norm, parts[0]])
                rich_text.extend([f_ul_txt, sunucu])
                if len(parts) > 1 and parts[1]:
                    rich_text.extend([f_norm, parts[1]])
                worksheet.write_rich_string(satir_no, 0, *rich_text, f_norm) if len(rich_text) >= 2 else worksheet.write(satir_no, 0, yazarlar_metni, f_norm)
            else:
                worksheet.write(satir_no, 0, yazarlar_metni, f_norm)
            worksheet.write(satir_no, 1, b["b"], f_norm)
            satir_no += 1
        satir_no += 1

    workbook.close()
    output.seek(0)
    return output


def oturum_ozeti_olustur(df_master):
    df_ozet = df_master[["Gun_ve_Saat", "Salon", "Oturum_ID", "Sunum_Tipi"]].copy()
    df_ozet = df_ozet[~df_ozet["Gun_ve_Saat"].isin(["-", "ATANMADI", "İPTAL EDİLDİ", "nan", "NaN", ""])]
    df_ozet["Salon"] = df_ozet["Salon"].apply(clean_salon)
    df_grouped = df_ozet.groupby(["Gun_ve_Saat", "Salon", "Oturum_ID", "Sunum_Tipi"]).size().reset_index(name="Atanan_Bildiri_Sayisi")
    df_grouped["sort_time"] = df_grouped["Gun_ve_Saat"].apply(parse_zaman)
    df_grouped = df_grouped.sort_values(by=["sort_time", "Salon"]).drop(columns=["sort_time"])
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_grouped.to_excel(writer, index=False, sheet_name="Oturum_Rontgeni")
        workbook = writer.book
        worksheet = writer.sheets["Oturum_Rontgeni"]
        fmt = workbook.add_format({"bold": True, "bg_color": "#4472C4", "font_color": "white", "border": 1})
        for col_num, value in enumerate(df_grouped.columns.values):
            worksheet.write(0, col_num, value, fmt)
        worksheet.set_column("A:A", 35)
        worksheet.set_column("B:B", 25)
        worksheet.set_column("C:E", 20)
    output.seek(0)
    return output


# =========================
# RAPORLAR
# =========================

def dataframe_to_excel_bytes(df: pd.DataFrame, sheet_name="Rapor"):
    out = io.BytesIO()
    df.to_excel(out, sheet_name=sheet_name, index=False)
    return out.getvalue()


def make_genel_report(participants_df, authors_df, submissions_df):
    title_by_id = dict(zip(submissions_df["id"], submissions_df["title"])) if not submissions_df.empty else {}
    sub_by_participant = {}
    for _, a in authors_df.iterrows():
        pid = a.get("participant_id")
        if pid:
            sub_by_participant.setdefault(pid, []).append(title_by_id.get(a["submission_id"], ""))
    rows = []
    for _, p in participants_df.iterrows():
        b_list = sub_by_participant.get(p["id"], [])
        rows.append(
            {
                "Unvan": turkce_buyuk(safe_str(p.get("title"))),
                "Adı Soyadı": safe_str(p.get("original_name")),
                "Ödeme Durumu": turkce_buyuk(safe_str(p.get("payment"))),
                "Bildiri 1": turkce_buyuk(b_list[0]) if len(b_list) > 0 else "",
                "Bildiri 2": turkce_buyuk(b_list[1]) if len(b_list) > 1 else "",
                "Katılım Türü": turkce_buyuk(safe_str(p.get("attendance_type"))),
                "Telefon": safe_str(p.get("phone")),
                "E-Posta": safe_str(p.get("email")),
            }
        )
    return dataframe_to_excel_bytes(pd.DataFrame(rows).sort_values("Adı Soyadı") if rows else pd.DataFrame())


def make_unpaid_report(submissions_df, authors_df):
    rows = []
    for _, s in submissions_df.iterrows():
        if safe_str(s.get("payment")) == "Hayır":
            sub_authors = authors_df[authors_df["submission_id"] == s["id"]].sort_values("author_order")
            rows.append(
                {
                    "KABUL EDİLMİŞ AMA ÖDENMEMİŞ BİLDİRİ": turkce_buyuk(safe_str(s.get("title"))),
                    "YAZARLARI": ", ".join([turkce_buyuk(x) for x in sub_authors["display_name"]]),
                }
            )
    return dataframe_to_excel_bytes(pd.DataFrame(rows))


def make_participant_type_report(participants_df, text):
    rows = []
    for _, p in participants_df.iterrows():
        if super_temiz(text) in super_temiz(safe_str(p.get("attendance_type"))):
            rows.append(
                {
                    "AD SOYAD": safe_str(p.get("original_name")),
                    "UNVAN": turkce_buyuk(safe_str(p.get("title"))),
                    "KATILIM TÜRÜ": turkce_buyuk(safe_str(p.get("attendance_type"))),
                    "ÖDEME DURUMU": safe_str(p.get("payment")),
                    "TELEFON": safe_str(p.get("phone")),
                    "E-POSTA": safe_str(p.get("email")),
                }
            )
    return dataframe_to_excel_bytes(pd.DataFrame(rows).sort_values("AD SOYAD") if rows else pd.DataFrame(columns=["AD SOYAD", "UNVAN", "KATILIM TÜRÜ", "ÖDEME DURUMU", "TELEFON", "E-POSTA"]))


def make_meal_report(participants_df, keys):
    rows = []
    for _, p in participants_df.iterrows():
        etk = super_temiz(safe_str(p.get("events")))
        if all(super_temiz(x) in etk for x in keys):
            rows.append({"AD SOYAD": safe_str(p.get("original_name")), "TUR": turkce_buyuk(safe_str(p.get("attendance_type"))), "ODEME": turkce_buyuk(safe_str(p.get("payment")))})
    return dataframe_to_excel_bytes(pd.DataFrame(rows) if rows else pd.DataFrame(columns=["AD SOYAD", "TUR", "ODEME"]))


def export_master_excel(snapshot):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        for key, sheet_name in [
            ("participants", "participants"),
            ("submissions", "submissions"),
            ("authors", "submission_authors"),
            ("moderators", "moderators"),
            ("special_events", "special_events"),
        ]:
            snapshot[key].to_excel(writer, sheet_name=sheet_name, index=False)
    return out.getvalue()


# =========================
# STREAMLIT
# =========================

st.set_page_config(page_title="IHMC 2026 Supabase Paneli", layout="wide", page_icon="📊")

st.markdown(
    """
    <style>
    .main { background-color: #f8fafc; }
    .stButton>button, .stDownloadButton>button { width: 100%; border-radius: 6px; font-weight: 700; }
    div[data-testid="stFormSubmitButton"] button {
        background-color: #2563eb !important;
        color: white !important;
        border: none !important;
        height: 3em;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("IHMC 2026 Kongre Yönetim Paneli")
st.caption("Excel'den Supabase'e aktar, programı düzenle, salonları kontrol et, matbaa çıktılarını üret.")

if get_supabase() is None:
    st.error("Supabase bağlantısı bulunamadı.")
    st.write("Streamlit Cloud > App > Settings > Secrets bölümüne şunu ekle:")
    st.code(
        'SUPABASE_URL = "https://xxxxx.supabase.co"\nSUPABASE_SERVICE_ROLE_KEY = "supabase-service-role-key"',
        language="toml",
    )
    st.stop()

with st.sidebar:
    st.header("Veri")
    if st.button("Verileri Yenile", key="refresh_data"):
        st.cache_data.clear()
        st.rerun()

try:
    snapshot = load_snapshot()
except Exception as e:
    st.error(f"Supabase okuma hatası: {e}")
    st.info("Önce Supabase SQL Editor'de `supabase_schema.sql` dosyasındaki SQL'i çalıştırman gerekiyor.")
    st.stop()

bildiriler, katilimcilar = build_lookup(snapshot)
df_master = df_master_from_snapshot(snapshot)
df_mod_matbaa = moderators_df_for_matbaa(snapshot)
df_ozel_matbaa = special_events_df_for_matbaa(snapshot)

tab_import, tab_search, tab_program, tab_salon, tab_mod, tab_matbaa, tab_reports = st.tabs(
    ["İçe Aktar", "Hızlı Sorgu", "Bildiri Programı", "Salon Kontrolü", "Moderatörler", "Matbaa", "Raporlar"]
)

with tab_import:
    st.subheader("Excel'den Supabase'e İlk Aktarım")
    st.warning("Bu aktarım mevcut Supabase verisini silip Excel'den yeniden kurar. Normal kullanımda sadece ilk kurulumda veya büyük veri yenilemede kullan.")
    c1, c2 = st.columns(2)
    katilimci_file = c1.file_uploader("Katılımcı Excel'i", type=["xlsx", "xls"], key="import_katilimci_file")
    program_file = c2.file_uploader("Program/Matbaa Excel'i", type=["xlsx", "xls"], key="import_program_file")
    st.caption("Tek ana Excel kullanıyorsan aynı dosyayı sadece ilk alana yükle; ikinci alan boş kalabilir.")
    confirm = st.checkbox("Mevcut Supabase verisini silip bu Excel'den yeniden yükleyeceğimi onaylıyorum.", key="import_confirm")
    if st.button("Supabase'e Aktar", key="import_button"):
        if not katilimci_file:
            st.error("En az bir Excel dosyası yüklemelisin.")
        elif not confirm:
            st.error("Silme/yükleme onayını işaretlemen gerekiyor.")
        else:
            try:
                with st.spinner("Excel okunuyor ve Supabase'e aktarılıyor..."):
                    summary = import_excel_to_supabase(katilimci_file, program_file)
                st.cache_data.clear()
                st.success(
                    f"Aktarım tamamlandı: {summary['participants']} kişi, {summary['submissions']} bildiri, "
                    f"{summary['authors']} yazar bağlantısı, {summary['moderators']} moderatör, "
                    f"{summary['special_events']} özel etkinlik."
                )
                st.rerun()
            except Exception as e:
                st.error(f"Aktarım hatası: {e}")

    st.write("#### Mevcut Supabase Durumu")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Katılımcı", len(snapshot["participants"]))
    c2.metric("Bildiri", len(snapshot["submissions"]))
    c3.metric("Moderatör", len(snapshot["moderators"]))
    c4.metric("Özel Etkinlik", len(snapshot["special_events"]))

with tab_search:
    st.subheader("Katılımcı veya Bildiri Ara")
    if "arama_metni" not in st.session_state:
        st.session_state.arama_metni = ""
    with st.form("search_form"):
        c1, c2 = st.columns([4, 1])
        sorgu_input = c1.text_input("Arama", value=st.session_state.arama_metni, placeholder="Örn: Cihan veya Yapay Zeka", label_visibility="collapsed", key="search_query")
        arama = c2.form_submit_button("Ara")
    if arama:
        st.session_state.arama_metni = sorgu_input
    sorgu = temiz_metin(st.session_state.arama_metni)
    if sorgu:
        found_people = [k for k in katilimcilar if sorgu in k]
        found_subs = [k for k in bildiriler if sorgu in k]
        if found_people:
            st.write(f"### Kişi Sonuçları ({len(found_people)})")
            for key in found_people:
                p = katilimcilar[key]
                with st.expander(f"{get_kisi_renk_emoji(p)} {safe_str(p.get('title'))} {safe_str(p.get('original_name'))}", expanded=True):
                    c1, c2 = st.columns(2)
                    c1.write(f"**Ödeme:** {safe_str(p.get('payment'))}")
                    c1.write(f"**Görev:** {safe_str(p.get('role'))}")
                    c1.write(f"**Katılım Türü:** {safe_str(p.get('attendance_type'))}")
                    c1.write(f"**Telefon:** {safe_str(p.get('phone')) or 'Kayıtlı değil'}")
                    c1.write(f"**E-posta:** {safe_str(p.get('email')) or 'Kayıtlı değil'}")
                    c2.write(f"**Kurum:** {safe_str(p.get('institution')) or '-'}")
                    c2.write(f"**Etkinlikler:** {safe_str(p.get('events')) or 'Yok'}")
                    days = p.get("days") if isinstance(p.get("days"), dict) else {}
                    c2.write(f"**Günler:** 6: {days.get('6', '')} | 7: {days.get('7', '')} | 8: {days.get('8', '')}")
                    st.write("**Bildirileri:**")
                    if p["submissions"]:
                        for title in p["submissions"]:
                            sub = next((b for b in bildiriler.values() if b["title"] == title), None)
                            st.write(f"- {title}")
                            if sub:
                                st.caption(program_info(sub))
                    else:
                        st.write("Bildirisi yok.")
        if found_subs:
            st.write(f"### Bildiri Sonuçları ({len(found_subs)})")
            for key in found_subs:
                b = bildiriler[key]
                with st.expander(f"{get_bildiri_renk_emoji(b)} {safe_str(b.get('title'))}", expanded=True):
                    st.write(f"**Konu:** {safe_str(b.get('topic'))} | **Butik:** {safe_str(b.get('boutique'))}")
                    st.write(f"**Ödeme:** {'Evet, Yapıldı' if b.get('payment') == 'Evet' else 'Hayır, Bekliyor'}")
                    st.write(f"**Program:** {program_info(b)}")
                    st.write("**Yazarlar:**")
                    for author_norm in b["authors"]:
                        p = katilimcilar.get(author_norm, {})
                        st.write(f"{get_kisi_renk_emoji(p)} {safe_str(p.get('title'))} {safe_str(p.get('original_name', author_norm))}")
        if not found_people and not found_subs:
            st.warning("Sonuç bulunamadı.")

with tab_program:
    st.subheader("Bildiri Gün/Saat/Salon/Oturum/Sunum Tipi Yönetimi")
    submissions_df = snapshot["submissions"].copy()
    authors_df = snapshot["authors"].copy()
    if submissions_df.empty:
        st.warning("Henüz bildiri yok. Önce Excel aktarımı yap.")
    else:
        f1, f2, f3 = st.columns(3)
        q = f1.text_input("Bildiri adı filtrele", "", key="program_filter_title")
        tip_filter = f2.multiselect("Sunum tipi", sorted([x for x in submissions_df["presentation_type"].dropna().unique() if safe_str(x)]), key="program_filter_type")
        gunler = sorted({gun_key(x) for x in submissions_df["session_time"].dropna().astype(str) if safe_str(x)}, key=parse_zaman)
        gun_filter = f3.multiselect("Gün", gunler, key="program_filter_day")

        df_show = submissions_df.copy()
        if q:
            df_show = df_show[df_show["title"].astype(str).apply(lambda x: temiz_metin(q) in temiz_metin(x))]
        if tip_filter:
            df_show = df_show[df_show["presentation_type"].isin(tip_filter)]
        if gun_filter:
            df_show = df_show[df_show["session_time"].astype(str).apply(lambda x: gun_key(x) in gun_filter)]

        show_cols = ["id", "title", "presentation_type", "session_time", "hall", "session_id", "topic", "payment"]
        st.dataframe(df_show[show_cols], use_container_width=True, hide_index=True)
        if not df_show.empty:
            labels = {
                f"{r['id']} | {safe_str(r['title'])[:90]} | {safe_str(r['presentation_type'])} | {safe_str(r['session_time'])} | {safe_str(r['hall'])}": r
                for _, r in df_show.iterrows()
            }
            selected_label = st.selectbox("Düzenlenecek bildiri", list(labels.keys()), key="program_selected_submission")
            selected = labels[selected_label]
            sessions = submissions_df[["presentation_type", "session_time", "hall", "session_id"]].drop_duplicates()
            session_labels = [
                f"{safe_str(r['presentation_type'])} | {safe_str(r['session_time'])} | {safe_str(r['hall'])} | {safe_str(r['session_id'])}"
                for _, r in sessions.iterrows()
            ]
            with st.form("program_update_form"):
                use_existing = st.checkbox("Mevcut oturuma taşı", value=True, key="program_use_existing")
                if use_existing and session_labels:
                    current_label = f"{safe_str(selected['presentation_type'])} | {safe_str(selected['session_time'])} | {safe_str(selected['hall'])} | {safe_str(selected['session_id'])}"
                    idx = session_labels.index(current_label) if current_label in session_labels else 0
                    chosen = st.selectbox("Hedef oturum", session_labels, index=idx, key="program_target_session")
                    p = chosen.split(" | ", 3)
                    new_type, new_time, new_hall, new_session = p[0], p[1], p[2], p[3]
                else:
                    c1, c2 = st.columns(2)
                    new_type = c1.selectbox("Sunum tipi", ["Yuzyuze", "Online"], index=1 if safe_str(selected["presentation_type"]) == "Online" else 0, key="program_manual_type")
                    new_time = c2.text_input("Gün ve saat", value=safe_str(selected["session_time"]), key="program_manual_time")
                    c3, c4 = st.columns(2)
                    halls = sorted({safe_str(x) for x in submissions_df["hall"].dropna().unique() if safe_str(x)})
                    new_hall = c3.selectbox("Salon", halls, index=halls.index(safe_str(selected["hall"])) if safe_str(selected["hall"]) in halls else 0, key="program_manual_hall") if halls else c3.text_input("Salon", value=safe_str(selected["hall"]), key="program_manual_hall_text")
                    new_session = c4.text_input("Oturum ID", value=safe_str(selected["session_id"]), key="program_manual_session")

                sub_authors = authors_df[authors_df["submission_id"] == selected["id"]].sort_values("author_order")
                author_labels = [safe_str(x) for x in sub_authors["display_name"]]
                author_norm_by_label = dict(zip(author_labels, sub_authors["norm_name"]))
                presenter_choice = st.selectbox("Sunacak yazar", ["Değiştirme"] + author_labels + ["Elle yaz"], key="program_presenter")
                presenter_free = ""
                if presenter_choice == "Elle yaz":
                    presenter_free = st.text_input("Yeni sunucu adı", value=safe_str(selected["presenter_norm_name"]), key="program_presenter_free")
                save = st.form_submit_button("Bildiri Güncelle")
            if save:
                presenter_norm = safe_str(selected["presenter_norm_name"])
                if presenter_choice != "Değiştirme":
                    presenter_norm = temiz_isim(presenter_free) if presenter_choice == "Elle yaz" else author_norm_by_label[presenter_choice]
                try:
                    update_by_id(
                        "submissions",
                        int(selected["id"]),
                        {
                            "presentation_type": new_type,
                            "session_time": new_time,
                            "hall": new_hall,
                            "session_id": new_session,
                            "presenter_norm_name": presenter_norm,
                        },
                    )
                    st.cache_data.clear()
                    st.success("Bildiri güncellendi.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Güncelleme hatası: {e}")

with tab_salon:
    st.subheader("Salon ve Oturum Kapasite Kontrolü")
    if df_master.empty:
        st.warning("Program verisi yok.")
    else:
        valid = df_master[~df_master["Gun_ve_Saat"].isin(["", "-", "nan", "NaN", "ATANMADI", "İPTAL EDİLDİ"])].copy()
        valid["Gün"] = valid["Gun_ve_Saat"].apply(gun_key)
        gunler = sorted(valid["Gün"].dropna().unique(), key=parse_zaman)
        secili_gun = st.selectbox("Kontrol edilecek gün", ["Tümü"] + list(gunler), key="salon_day")
        if secili_gun != "Tümü":
            valid = valid[valid["Gün"] == secili_gun]
        c1, c2, c3 = st.columns(3)
        c1.metric("Toplam Bildiri", len(valid))
        c2.metric("Aktif Salon", valid["Salon"].nunique())
        c3.metric("Oturum", valid[["Gun_ve_Saat", "Salon", "Oturum_ID"]].drop_duplicates().shape[0])
        counts = valid.groupby("Salon").size().reset_index(name="Bildiri Sayısı").sort_values("Bildiri Sayısı", ascending=False)
        st.dataframe(counts, use_container_width=True, hide_index=True)
        st.bar_chart(counts.set_index("Salon"))
        pivot = valid.pivot_table(index="Gun_ve_Saat", columns="Salon", values="Bildiri_Adi", aggfunc="count", fill_value=0)
        pivot = pivot.sort_index(key=lambda idx: idx.map(parse_zaman))
        st.write("#### Saat x Salon Haritası")
        st.dataframe(pivot, use_container_width=True)

with tab_mod:
    st.subheader("Moderatör ve Oturum Değerlendirici Yönetimi")
    moderators_df = snapshot["moderators"].copy()
    if moderators_df.empty:
        st.warning("Moderatör verisi yok.")
    else:
        st.dataframe(moderators_df[["id", "name", "institution", "duty", "session_time", "hall", "session_id", "online"]], use_container_width=True, hide_index=True)
        mod_tab1, mod_tab2 = st.tabs(["Tek Kişi Değiştir", "İki Kişiyi Yer Değiştir"])
        with mod_tab1:
            labels = {
                f"{r['id']} | {safe_str(r['name'])} | {safe_str(r['duty'])} | {safe_str(r['session_time'])} | {safe_str(r['hall'])}": r
                for _, r in moderators_df.iterrows()
            }
            selected_label = st.selectbox("Düzenlenecek moderatör/değerlendirici", list(labels.keys()), key="mod_selected")
            selected = labels[selected_label]
            sessions = snapshot["submissions"][["presentation_type", "session_time", "hall", "session_id"]].drop_duplicates()
            session_labels = [f"{safe_str(r['presentation_type'])} | {safe_str(r['session_time'])} | {safe_str(r['hall'])} | {safe_str(r['session_id'])}" for _, r in sessions.iterrows()]
            with st.form("mod_update_form"):
                c1, c2 = st.columns(2)
                new_name = c1.text_input("Ad Soyad / Ünvan", value=safe_str(selected["name"]), key="mod_name")
                new_inst = c2.text_input("Kurum", value=safe_str(selected["institution"]), key="mod_institution")
                c3, c4 = st.columns(2)
                new_duty = c3.text_input("Görev", value=safe_str(selected["duty"]), key="mod_duty")
                new_online_text = c4.selectbox("Online", ["Hayır", "Evet"], index=1 if bool(selected["online"]) else 0, key="mod_online")
                use_session = st.checkbox("Mevcut oturuma ata", value=True, key="mod_use_session")
                if use_session and session_labels:
                    chosen = st.selectbox("Hedef oturum", session_labels, key="mod_target_session")
                    p = chosen.split(" | ", 3)
                    new_online_text = "Evet" if p[0] == "Online" else "Hayır"
                    new_time, new_hall, new_session = p[1], p[2], p[3]
                else:
                    c5, c6 = st.columns(2)
                    new_time = c5.text_input("Gün ve saat", value=safe_str(selected["session_time"]), key="mod_time")
                    new_hall = c6.text_input("Salon", value=safe_str(selected["hall"]), key="mod_hall")
                    new_session = st.text_input("Oturum ID", value=safe_str(selected["session_id"]), key="mod_session")
                save_mod = st.form_submit_button("Moderatör Kaydet")
            if save_mod:
                try:
                    update_by_id(
                        "moderators",
                        int(selected["id"]),
                        {
                            "name": new_name,
                            "institution": new_inst,
                            "duty": new_duty,
                            "session_time": new_time,
                            "hall": new_hall,
                            "session_id": new_session,
                            "online": new_online_text == "Evet",
                        },
                    )
                    st.cache_data.clear()
                    st.success("Moderatör güncellendi.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Güncelleme hatası: {e}")

        with mod_tab2:
            labels = {
                f"{r['id']} | {safe_str(r['name'])} | {safe_str(r['duty'])} | {safe_str(r['session_time'])} | {safe_str(r['hall'])}": r
                for _, r in moderators_df.iterrows()
            }
            c1, c2 = st.columns(2)
            left_label = c1.selectbox("1. kişi", list(labels.keys()), key="swap_left")
            right_label = c2.selectbox("2. kişi", list(labels.keys()), key="swap_right")
            swap_mode = st.radio("Değişim tipi", ["Oturumlarını değiştir", "Kişi isimlerini değiştir"], horizontal=True, key="swap_mode")
            if st.button("Yer Değiştir", key="swap_button"):
                if left_label == right_label:
                    st.warning("İki farklı kişi seçmelisin.")
                else:
                    left = labels[left_label]
                    right = labels[right_label]
                    fields = ["session_time", "hall", "session_id", "online"] if swap_mode == "Oturumlarını değiştir" else ["name", "institution"]
                    try:
                        update_by_id("moderators", int(left["id"]), {f: right[f] for f in fields})
                        update_by_id("moderators", int(right["id"]), {f: left[f] for f in fields})
                        st.cache_data.clear()
                        st.success("Değişim yapıldı.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Değişim hatası: {e}")

with tab_matbaa:
    st.subheader("Matbaa Çıktıları")
    if df_master.empty:
        st.warning("Henüz program verisi yok.")
    else:
        c1, c2, c3 = st.columns(3)
        c1.metric("Toplam Bildiri", len(df_master))
        c2.metric("Yüz yüze", (df_master["Sunum_Tipi"] == "Yuzyuze").sum())
        c3.metric("Online", (df_master["Sunum_Tipi"] == "Online").sum())
        if st.button("Matbaayı Çalıştır ve Dosyaları Hazırla", key="matbaa_build"):
            with st.spinner("Excel ve Word çıktıları hazırlanıyor..."):
                df_yz = df_master[df_master["Sunum_Tipi"] == "Yuzyuze"]
                df_on = df_master[df_master["Sunum_Tipi"] == "Online"]
                st.session_state["matbaa_outputs"] = {
                    "excel_yz": excel_bas(df_yz, df_ozel_matbaa, df_mod_matbaa, "YUZYUZE").getvalue(),
                    "word_yz": word_bas(df_yz, df_ozel_matbaa, df_mod_matbaa, "YUZYUZE").getvalue(),
                    "excel_on": excel_bas(df_on, df_ozel_matbaa, df_mod_matbaa, "ONLINE").getvalue(),
                    "word_on": word_bas(df_on, df_ozel_matbaa, df_mod_matbaa, "ONLINE").getvalue(),
                    "ozet": oturum_ozeti_olustur(df_master).getvalue(),
                }
                st.success("Matbaa çıktıları hazır.")
        outputs = st.session_state.get("matbaa_outputs")
        if outputs:
            c1, c2 = st.columns(2)
            with c1:
                st.write("#### Yüz Yüze Program")
                st.download_button("Excel İndir", outputs["excel_yz"], file_name="Nihai_Program_Yuzyuze.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_yz_xlsx")
                st.download_button("Word İndir", outputs["word_yz"], file_name="Nihai_Program_Yuzyuze.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="dl_yz_docx")
            with c2:
                st.write("#### Online Program")
                st.download_button("Excel İndir", outputs["excel_on"], file_name="Nihai_Program_Online.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_on_xlsx")
                st.download_button("Word İndir", outputs["word_on"], file_name="Nihai_Program_Online.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="dl_on_docx")
            st.download_button("Oturum Röntgen Özeti İndir", outputs["ozet"], file_name="Oturum_Rontgen_Ozeti.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_ozet")

with tab_reports:
    st.subheader("Raporlar ve Yedek")
    participants_df = snapshot["participants"]
    submissions_df = snapshot["submissions"]
    authors_df = snapshot["authors"]
    c1, c2 = st.columns(2)
    c1.download_button("Genel Katılımcı Analizi", data=make_genel_report(participants_df, authors_df, submissions_df), file_name="IHMC_Genel_Rapor.xlsx", use_container_width=True, key="dl_genel")
    c1.download_button("Sadece Online Katılımcılar", data=make_participant_type_report(participants_df, "ONLINE"), file_name="Online_Katilimcilar.xlsx", use_container_width=True, key="dl_online")
    c2.download_button("Ödenmemiş Bildiriler", data=make_unpaid_report(submissions_df, authors_df), file_name="Odenmemis_Bildiriler.xlsx", use_container_width=True, key="dl_unpaid")
    c2.download_button("Sadece Fiziki Katılımcılar", data=make_participant_type_report(participants_df, "FİZİKİ"), file_name="Fiziki_Katilimcilar.xlsx", use_container_width=True, key="dl_fiziki")
    st.write("#### Yemek ve Etkinlik Listeleri")
    m1, m2, m3 = st.columns(3)
    m1.download_button("07 Mayıs Öğle", data=make_meal_report(participants_df, ["7", "OGLE"]), file_name="07_Mayis_Ogle.xlsx", use_container_width=True, key="dl_meal_07")
    m2.download_button("07 Mayıs Gala", data=make_meal_report(participants_df, ["GALA"]), file_name="07_Mayis_Gala.xlsx", use_container_width=True, key="dl_gala_07")
    m3.download_button("08 Mayıs Öğle", data=make_meal_report(participants_df, ["8", "OGLE"]), file_name="08_Mayis_Ogle.xlsx", use_container_width=True, key="dl_meal_08")
    st.write("#### Supabase Yedeği")
    st.download_button("Tüm Veriyi Excel Yedeği Olarak İndir", data=export_master_excel(snapshot), file_name="IHMC_Supabase_Yedek.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="dl_backup")
