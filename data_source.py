"""
Sumber data dashboard DMF 3.

Biaya iklan, nomor, dan status campaign: langsung dari Meta Ads.

Dari Google Sheet DMF 3 (SHEET_ID):
- CS    : daftar CS (Nama Akun), modal (Penunaian), profit/loss total,
          dan sisa dana per CS.
- Admin : data transfer donasi (TANGGAL, NAMA CS, KET NOMOR, NOMINAL),
          sumber resi & donasi. Bulan = bulan tanggal transfer.

Data donatur (nama, nomor WA) TIDAK pernah dikeluarkan dari modul ini,
hanya angka rekap per CS.

Daftar CS & nama lain (alias) disimpan di data/cs.csv (menu Kelola CS).
"""
import datetime as dt
import io
import os
from pathlib import Path

import pandas as pd
import requests

DATA_DIR = Path(__file__).parent / "data"
CS_FILE = DATA_DIR / "cs.csv"

BULAN_ID = {
    "januari": 1, "februari": 2, "maret": 3, "april": 4, "mei": 5, "juni": 6,
    "juli": 7, "agustus": 8, "september": 9, "oktober": 10, "november": 11,
    "desember": 12,
}


# =========================================================
# UTIL
# =========================================================

def norm(teks) -> str:
    """Nama dibandingkan tanpa beda huruf besar/kecil dan spasi."""
    if teks is None or (isinstance(teks, float) and pd.isna(teks)):
        return ""
    return " ".join(str(teks).split()).lower()


def kode_bulan(nilai):
    """Tanggal atau teks 'September 2026' -> '2026-09'."""
    if isinstance(nilai, (dt.datetime, dt.date, pd.Timestamp)):
        return f"{nilai.year}-{nilai.month:02d}"
    if isinstance(nilai, str):
        bagian = nilai.strip().lower().split()
        if len(bagian) == 2 and bagian[0] in BULAN_ID and bagian[1].isdigit():
            return f"{bagian[1]}-{BULAN_ID[bagian[0]]:02d}"
    return None


def angka(nilai) -> float:
    try:
        return float(nilai)
    except (TypeError, ValueError):
        return 0.0


# =========================================================
# AMBIL FILE SHEET
# =========================================================

def unduh_sheet(sheet_id: str, akun_layanan: dict | None = None) -> bytes:
    """Unduh spreadsheet sebagai .xlsx.

    Kalau ada akun layanan (service account) Google, dipakai supaya sheet
    tidak perlu dibuka untuk umum. Tanpa itu, sheet harus bisa dibuka
    "siapa saja yang punya link".
    """
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    if akun_layanan:
        from google.auth.transport.requests import AuthorizedSession
        from google.oauth2.service_account import Credentials

        creds = Credentials.from_service_account_info(
            akun_layanan,
            scopes=["https://www.googleapis.com/auth/drive.readonly"],
        )
        resp = AuthorizedSession(creds).get(url, timeout=120)
    else:
        resp = requests.get(url, timeout=120)
    resp.raise_for_status()
    if "spreadsheetml" not in resp.headers.get("Content-Type", ""):
        raise RuntimeError(
            "Sheet tidak bisa dibaca. Pastikan sheet dibagikan ke akun layanan "
            "dashboard (atau bisa dibuka lewat link)."
        )
    return resp.content


def baca_sheet(isi_xlsx: bytes) -> dict:
    return pd.read_excel(
        io.BytesIO(isi_xlsx),
        sheet_name=["Admin", "CS"],
        header=None,
        engine="openpyxl",
    )


# =========================================================
# OLAH DATA
# =========================================================

def jenis_resi(teks) -> str:
    """'PERDANA' / 'PERRDANA' -> 'perdana', 'GULUNGAN' -> 'gulungan'."""
    k = norm(teks).replace(" ", "")
    if k.startswith("perd") or k.startswith("perrd"):
        return "perdana"
    if k.startswith("gul"):
        return "gulungan"
    return ""


def olah_admin(admin: pd.DataFrame) -> pd.DataFrame:
    """Tab Admin. Baris 2 = judul kolom: NO, TANGGAL, NAMA CS, NAMA DONATUR,
    NOMOR WA, PROGRAM, KET NOMOR, BANK, NOMINAL. Rekap di kanan diabaikan."""
    judul = [norm(x) for x in admin.iloc[1]]
    df = admin.iloc[2:].copy()
    df.columns = range(df.shape[1])

    def kol(nama):
        return df[judul.index(nama)]

    tanggal = pd.to_datetime(kol("tanggal"), errors="coerce")
    jenis = kol("ket nomor").map(jenis_resi)
    nominal = kol("nominal").map(angka)
    hasil = pd.DataFrame({
        "tanggal": tanggal.dt.normalize(),
        "kunci": kol("nama cs").map(norm),
        "perdana": (jenis == "perdana").astype(int),
        "gulungan": (jenis == "gulungan").astype(int),
        "donasi_perdana": nominal.where(jenis == "perdana", 0.0),
        "donasi_gulungan": nominal.where(jenis == "gulungan", 0.0),
    })
    hasil = hasil[(hasil["kunci"] != "") & hasil["tanggal"].notna()]
    return hasil.groupby(["tanggal", "kunci"], as_index=False).sum()


def olah_dana(cs_tab: pd.DataFrame) -> pd.DataFrame:
    """Tab CS: D=Nama Akun, I=Penunaian (modal disetor),
    O=profit/loss total (judul kolomnya "Spend Iklan"), P=Sisa Dana."""
    rows = []
    for _, r in cs_tab.iloc[2:].iterrows():
        nama = r.iloc[3]
        if norm(nama) in ("", "nama akun", "jumlah"):
            continue
        rows.append({
            "kunci": norm(nama),
            "modal_iklan": angka(r.iloc[8]),
            "akumulasi": angka(r.iloc[14]),
            "sisa_iklan": angka(r.iloc[15]),
        })
    return pd.DataFrame(rows, columns=["kunci", "modal_iklan", "akumulasi", "sisa_iklan"])


def cs_dari_sheet(cs_tab: pd.DataFrame) -> pd.DataFrame:
    """Daftar CS dari tab CS: nama = Nama Akun, nama lain = nama lengkap."""
    rows = []
    for _, r in cs_tab.iloc[2:].iterrows():
        akun, lengkap = str(r.iloc[3]).strip(), str(r.iloc[1]).strip()
        if norm(akun) in ("", "nan", "nama akun", "jumlah"):
            continue
        rows.append({"nama": akun.title(), "aktif": True,
                     "nama_lain": "" if norm(lengkap) in ("", "nan", norm(akun)) else lengkap.title()})
    return pd.DataFrame(rows, columns=["nama", "aktif", "nama_lain"]).drop_duplicates("nama")


# =========================================================
# DAFTAR CS & ALIAS
# =========================================================

def load_cs() -> pd.DataFrame:
    kolom = {"nama": str, "aktif": bool, "nama_lain": str}
    if not CS_FILE.exists():
        return pd.DataFrame({k: pd.Series(dtype=t) for k, t in kolom.items()})
    df = pd.read_csv(CS_FILE, encoding="utf-8-sig", dtype={"nama": str, "nama_lain": str})
    if "nama_lain" not in df:
        df["nama_lain"] = ""
    df["nama_lain"] = df["nama_lain"].fillna("")
    df["aktif"] = df["aktif"].astype(str).str.lower().isin(["true", "1", "ya"])
    return df[list(kolom)]


def save_cs(df: pd.DataFrame) -> None:
    df = df.copy()
    df["nama"] = df["nama"].astype(str).str.strip()
    df = df[df["nama"].ne("") & df["nama"].ne("nan") & df["nama"].ne("None")]
    df = df.drop_duplicates(subset="nama", keep="first")
    df["aktif"] = df["aktif"].fillna(True).astype(bool)
    df["nama_lain"] = df["nama_lain"].fillna("").astype(str).str.strip()
    df[["nama", "aktif", "nama_lain"]].to_csv(CS_FILE, index=False)


def peta_alias(cs: pd.DataFrame) -> dict:
    """{nama/alias ternormalisasi: nama CS}. Alias dipisah koma."""
    peta = {}
    for _, r in cs.iterrows():
        peta[norm(r["nama"])] = r["nama"]
        for alias in str(r.get("nama_lain") or "").split(","):
            if norm(alias):
                peta[norm(alias)] = r["nama"]
    return peta


# =========================================================
# GABUNG
# =========================================================

def siapkan(isi_xlsx: bytes, cs: pd.DataFrame, isi_recap: bytes | None = None):
    """Hasil: (resi & donasi per hari per CS, dana iklan per CS, nama tak dikenal).

    Resi & donasi dari tab Admin, dana dari tab CS. Biaya iklan & nomor tidak
    dari sheet, tapi dari Meta (iklan_meta_rentang). isi_recap tidak dipakai.
    """
    sheet = baca_sheet(isi_xlsx)
    resi = olah_admin(sheet["Admin"])
    dana = olah_dana(sheet["CS"])

    peta = peta_alias(cs)
    tak_dikenal = set()

    def pasang_nama(df):
        df = df.copy()
        df["nama"] = df["kunci"].map(peta)
        tak_dikenal.update(df.loc[df["nama"].isna(), "kunci"])
        return df.dropna(subset=["nama"]).drop(columns="kunci")

    resi, dana = pasang_nama(resi), pasang_nama(dana)

    performa = resi.groupby(["tanggal", "nama"], as_index=False).sum(numeric_only=True)
    dana = dana.groupby("nama", as_index=False).sum(numeric_only=True)
    return performa, dana, sorted(tak_dikenal)


def sheet_id_default() -> str | None:
    return os.environ.get("SHEET_ID")


# =========================================================
# STATUS CAMPAIGN DARI META ADS
# =========================================================

META_API = "https://graph.facebook.com/v21.0"

# Sumber biaya & nomor iklan dari Meta (menggantikan catatan manual Madha).
#   kata=None : semua ad set, dicocokkan ke CS lewat nama ad set / campaign
#   kata="x"  : hanya ad set yang namanya memuat "x", dicatat atas nama `cs`
def aturan_meta() -> list[dict]:
    """Akun iklan DMF 3, dari Secrets / .env: META_AKUN (pisahkan koma kalau
    lebih dari satu, tanpa "act_"). Semua ad set dicocokkan ke CS lewat nama
    ad set / campaign."""
    akun = os.environ.get("META_AKUN", "")
    return [
        {"akun": a.strip().removeprefix("act_"), "kata": None, "cs": None}
        for a in akun.split(",") if a.strip()
    ]


def cs_untuk_iklan(akun: str, adset: str, campaign: str, peta: dict):
    """Nama CS pemilik iklan menurut aturan_meta() (None = diabaikan)."""
    aturan = next((a for a in aturan_meta() if a["akun"] == akun), None)
    if aturan and aturan["kata"]:
        return aturan["cs"] if aturan["kata"] in norm(adset) else None
    return cari_cs(adset, peta) or cari_cs(campaign, peta)


def iklan_aktif(token: str, akun_iklan: list[str]) -> list[dict]:
    """Semua iklan yang benar-benar bisa tayang: status ACTIVE dan jadwal ad set
    belum selesai (ad set dengan tanggal selesai yang sudah lewat tetap
    berstatus ACTIVE di Meta, tapi tidak tayang lagi)."""
    sekarang = pd.Timestamp.now(tz="Asia/Jakarta")
    hasil = []
    for akun in akun_iklan:
        url = f"{META_API}/act_{akun.strip()}/ads"
        params = {
            "access_token": token,
            "limit": 500,
            "fields": "adset{name,end_time},campaign{name,stop_time}",
            "filtering": '[{"field":"effective_status","operator":"IN","value":["ACTIVE"]}]',
        }
        while url:
            data = requests.get(url, params=params, timeout=60).json()
            if "error" in data:
                raise RuntimeError(data["error"].get("message", "Error Meta API"))
            for ad in data.get("data", []):
                selesai = [
                    pd.Timestamp(t) for t in (
                        ad.get("adset", {}).get("end_time"), ad.get("campaign", {}).get("stop_time"),
                    ) if t
                ]
                if any(t <= sekarang for t in selesai):
                    continue
                hasil.append({
                    "akun": akun.strip(),
                    "adset": ad.get("adset", {}).get("name", ""),
                    "campaign": ad.get("campaign", {}).get("name", ""),
                })
            url, params = data.get("paging", {}).get("next"), None
    return hasil


def cari_cs(teks: str, peta: dict):
    """Cari nama CS (atau nama lain) sebagai kata utuh di teks, yang terpanjang dulu."""
    kata = f" {norm(''.join(c if c.isalnum() else ' ' for c in str(teks)))} "
    for kunci in sorted(peta, key=len, reverse=True):
        if f" {kunci} " in kata:
            return peta[kunci]
    return None


def status_campaign(token: str, akun_iklan: list[str], cs: pd.DataFrame) -> dict:
    """{nama CS: jumlah iklan aktif} dicocokkan dari nama ad set (lalu campaign)."""
    peta = peta_alias(cs)
    jumlah = {}
    for ad in iklan_aktif(token, akun_iklan):
        nama = cs_untuk_iklan(ad["akun"], ad["adset"], ad["campaign"], peta)
        if nama:
            jumlah[nama] = jumlah.get(nama, 0) + 1
    return jumlah


# =========================================================
# TOPUP IKLAN (INPUT ADMIN)
# =========================================================

TOPUP_TAB = "Topup Iklan"
TOPUP_KOLOM = ["Tanggal", "Nama CS", "Nominal", "Diinput"]
TOPUP_FILE = DATA_DIR / "topup.csv"


def _sesi_sheets(akun_layanan: dict):
    from google.auth.transport.requests import AuthorizedSession
    from google.oauth2.service_account import Credentials

    creds = Credentials.from_service_account_info(
        akun_layanan, scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    return AuthorizedSession(creds)


def _pastikan_tab_topup(sesi, sheet_id: str) -> None:
    info = sesi.get(
        f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}",
        params={"fields": "sheets.properties.title"}, timeout=60,
    )
    info.raise_for_status()
    judul = [s["properties"]["title"] for s in info.json().get("sheets", [])]
    if TOPUP_TAB in judul:
        return
    sesi.post(
        f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}:batchUpdate",
        json={"requests": [{"addSheet": {"properties": {"title": TOPUP_TAB}}}]},
        timeout=60,
    ).raise_for_status()
    sesi.put(
        f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}/values/'{TOPUP_TAB}'!A1:D1",
        params={"valueInputOption": "RAW"},
        json={"values": [TOPUP_KOLOM]}, timeout=60,
    ).raise_for_status()


def simpan_topup(baris: dict, sheet_id: str | None, akun_layanan: dict | None) -> str:
    """Simpan satu topup. Ke Google Sheet kalau ada akun layanan, kalau tidak ke file lokal.

    baris: {"tanggal": date, "nama": str, "nominal": int}
    Hasil: tempat data disimpan ("sheet" / "lokal").
    """
    nilai = [
        baris["tanggal"].strftime("%Y-%m-%d"),
        baris["nama"],
        int(baris["nominal"]),
        dt.datetime.now(dt.timezone(dt.timedelta(hours=7))).strftime("%Y-%m-%d %H:%M"),
    ]
    if akun_layanan and sheet_id:
        sesi = _sesi_sheets(akun_layanan)
        _pastikan_tab_topup(sesi, sheet_id)
        sesi.post(
            f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}/values/'{TOPUP_TAB}'!A:D:append",
            params={"valueInputOption": "USER_ENTERED", "insertDataOption": "INSERT_ROWS"},
            json={"values": [nilai]}, timeout=60,
        ).raise_for_status()
        return "sheet"

    baru = pd.DataFrame([nilai], columns=TOPUP_KOLOM)
    if TOPUP_FILE.exists():
        baru = pd.concat([pd.read_csv(TOPUP_FILE, encoding="utf-8-sig"), baru], ignore_index=True)
    baru.to_csv(TOPUP_FILE, index=False)
    return "lokal"


def baca_topup(isi_xlsx: bytes | None, dari_sheet: bool) -> pd.DataFrame:
    """Riwayat topup: dari tab 'Topup Iklan' di sheet, atau dari file lokal."""
    kosong = pd.DataFrame(columns=["tanggal", "nama", "nominal", "diinput", "bulan"])
    if dari_sheet and isi_xlsx:
        buku = pd.ExcelFile(io.BytesIO(isi_xlsx), engine="openpyxl")
        if TOPUP_TAB not in buku.sheet_names:
            return kosong
        df = buku.parse(TOPUP_TAB)
    elif TOPUP_FILE.exists():
        df = pd.read_csv(TOPUP_FILE, encoding="utf-8-sig")
    else:
        return kosong
    df = df.reindex(columns=TOPUP_KOLOM)
    df.columns = ["tanggal", "nama", "nominal", "diinput"]
    df["tanggal"] = pd.to_datetime(df["tanggal"], errors="coerce")
    df["nominal"] = df["nominal"].map(angka)
    df = df.dropna(subset=["tanggal"])
    df["bulan"] = df["tanggal"].map(kode_bulan)
    return df


def _hasil_meta(baris: dict) -> float:
    """'Hasil' iklan interaksi = percakapan pesan dimulai (sama dengan Ads Manager)."""
    aksi = {a["action_type"]: angka(a["value"]) for a in baris.get("actions", [])}
    return aksi.get("onsite_conversion.messaging_conversation_started_7d", aksi.get("lead", 0.0))


def iklan_meta_rentang(token: str, mulai: str, sampai: str, cs: pd.DataFrame) -> pd.DataFrame:
    """Biaya & hasil per CS untuk SATU rentang tanggal (sama dengan Ads Manager).

    Hasil per rentang tidak sama dengan jumlah hasil harian, karena Meta
    menghitung satu orang sekali dalam rentang yang diminta.
    """
    from concurrent.futures import ThreadPoolExecutor

    peta = peta_alias(cs)
    import json

    def ambil(akun):
        url = f"{META_API}/act_{akun}/insights"
        params = {
            "access_token": token, "level": "adset",
            "time_range": json.dumps({"since": mulai, "until": sampai}),
            "fields": "adset_name,campaign_name,spend,actions", "limit": 500,
        }
        hasil = []
        while url:
            data = requests.get(url, params=params, timeout=120).json()
            if "error" in data:
                raise RuntimeError(data["error"].get("message", "Error Meta API"))
            hasil += data.get("data", [])
            url, params = data.get("paging", {}).get("next"), None
        return akun, hasil

    with ThreadPoolExecutor(max_workers=4) as pool:
        semua = list(pool.map(ambil, [a["akun"] for a in aturan_meta()]))

    baris = []
    for akun, data in semua:
        for r in data:
            nama = cs_untuk_iklan(akun, r.get("adset_name"), r.get("campaign_name"), peta)
            if nama:
                baris.append({"nama": nama, "biaya_iklan": angka(r.get("spend")), "nomor": _hasil_meta(r)})
    df = pd.DataFrame(baris, columns=["nama", "biaya_iklan", "nomor"])
    return df.groupby("nama", as_index=False).sum()


def saldo_akun(token: str, akun: str) -> dict:
    """Sisa batas pengeluaran akun iklan (spend_cap - amount_spent), dalam Rupiah."""
    r = requests.get(f"{META_API}/act_{akun.strip()}", params={
        "access_token": token, "fields": "name,spend_cap,amount_spent,balance"}, timeout=60).json()
    if "error" in r:
        raise RuntimeError(r["error"].get("message", "Error Meta API"))
    batas = angka(r.get("spend_cap")) or None
    terpakai = angka(r.get("amount_spent"))
    nama = str(r.get("name", "")).replace("MJO - Harizi ", "")
    return {"nama": nama, "batas": batas, "terpakai": terpakai,
            "sisa": (batas - terpakai) if batas else None, "tagihan": angka(r.get("balance"))}
