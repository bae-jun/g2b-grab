# -*- coding: utf-8 -*-
"""
나라장터 입찰정보 수집 웹앱 (Streamlit) — 검색조건 확장판
검색조건: 검색유형(입찰공고/개찰결과), 기간, 업무구분, 참가제한지역, 추정가격
+ 입찰공고는 공고서 첨부파일에서 실무담당 전화번호 자동 추출
"""

import re
import io
import time
import zlib
import zipfile
import struct
import datetime
import urllib.parse

import requests
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill

st.set_page_config(page_title="나라장터 입찰정보 수집기", page_icon="📋",
                   layout="centered")

# ------------------------------------------------------------
# 코드표
# ------------------------------------------------------------
REGIONS = {
    "전국(제한없음)": None, "서울": "11", "부산": "26", "대구": "27",
    "인천": "28", "광주": "29", "대전": "30", "울산": "31", "세종": "36",
    "경기": "41", "강원": "51", "충북": "43", "충남": "44", "전북": "45",
    "전남": "46", "경북": "47", "경남": "48", "제주": "50",
}

# 업무구분 -> (입찰공고 오퍼레이션, 개찰결과 오퍼레이션)
TASKS = {
    "공사":  ("getBidPblancListInfoCnstwkPPSSrch", "getOpengResultListInfoCnstwk"),
    "물품":  ("getBidPblancListInfoThngPPSSrch",   "getOpengResultListInfoThng"),
    "용역":  ("getBidPblancListInfoServcPPSSrch",  "getOpengResultListInfoServc"),
    "외자":  ("getBidPblancListInfoFrgcptPPSSrch", "getOpengResultListInfoFrgcpt"),
}

BID_BASES = [  # 입찰공고정보서비스 (신/구 주소)
    "http://apis.data.go.kr/1230000/ad/BidPublicInfoService",
    "http://apis.data.go.kr/1230000/BidPublicInfoService05",
    "http://apis.data.go.kr/1230000/BidPublicInfoService04",
    "http://apis.data.go.kr/1230000/BidPublicInfoService",
]
SCSBID_BASES = [  # 낙찰정보서비스 (신/구 주소)
    "http://apis.data.go.kr/1230000/as/ScsbidInfoService",
    "http://apis.data.go.kr/1230000/ScsbidInfoService01",
    "http://apis.data.go.kr/1230000/ScsbidInfoService",
]

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
PHONE_RE = re.compile(r"0\d{1,2}[\s\-\.\)]{0,2}\d{3,4}[\s\-\.]{0,2}\d{4}")
CONTACT_KEYWORDS = ["문의", "담당", "감독", "실무", "연락처", "주무관",
                    "문의처", "담당자", "감독관", "공사감독", "전화",
                    "☎", "TEL", "Tel", "T.", "℡"]


# ------------------------------------------------------------
# 공통 함수
# ------------------------------------------------------------
def get_service_key():
    try:
        k = st.secrets["SERVICE_KEY"].strip()
    except Exception:
        return None
    if "%" in k:
        k = urllib.parse.unquote(k)
    return k


def call_api(bases, operation, extra_params, log):
    """여러 base URL을 순서대로 시도하며 전체 페이지 수집"""
    key = get_service_key()
    last_error = ""
    for base in bases:
        url = f"{base}/{operation}"
        items, page = [], 1
        try:
            while True:
                params = {"serviceKey": key, "pageNo": page, "numOfRows": 100,
                          "type": "json", **extra_params}
                r = requests.get(url, params=params, headers=HEADERS, timeout=30)
                if r.status_code != 200:
                    last_error = f"HTTP {r.status_code} ({url})"
                    break
                if r.text.lstrip().startswith("<"):
                    last_error = r.text[:300]
                    break
                body = r.json().get("response", {}).get("body", {})
                rows = body.get("items", [])
                if isinstance(rows, dict):
                    rows = rows.get("item", [])
                if not rows:
                    break
                items.extend(rows)
                total = int(body.get("totalCount", 0))
                log.write(f"목록 수신 중... {len(items)}/{total}건")
                if len(items) >= total:
                    break
                page += 1
                time.sleep(0.3)
            if items:
                return items, None
        except Exception as e:
            last_error = str(e)
            continue
    return [], last_error


def price_of(item):
    for f in ("presmptPrce", "asignBdgtAmt", "bdgtAmt", "sucsfbidAmt"):
        v = str(item.get(f, "") or "").replace(",", "")
        if v.replace(".", "").isdigit():
            return int(float(v))
    return 0


def attachments_of(item, limit):
    out = []
    for i in range(1, 11):
        u = item.get(f"ntceSpecDocUrl{i}")
        n = item.get(f"ntceSpecFileNm{i}") or f"file{i}"
        if u:
            out.append((n, u))
    return out[:limit]


# ---------- 문서 텍스트 추출 ----------
def text_from_pdf(data):
    try:
        import pdfplumber
        out = []
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            for pg in pdf.pages[:30]:
                out.append(pg.extract_text() or "")
        return "\n".join(out)
    except Exception:
        return ""


def text_from_hwpx(data):
    try:
        out = []
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            for name in z.namelist():
                if name.startswith("Contents/section"):
                    xml = z.read(name).decode("utf-8", "ignore")
                    out.append(re.sub(r"<[^>]+>", " ", xml))
        return "\n".join(out)
    except Exception:
        return ""


def text_from_hwp(data):
    try:
        import olefile
        if not olefile.isOleFile(io.BytesIO(data)):
            return ""
        ole = olefile.OleFileIO(io.BytesIO(data))
        header = ole.openstream("FileHeader").read()
        compressed = bool(header[36] & 1)
        chunks = []
        for entry in ole.listdir():
            if entry[0] == "BodyText":
                raw = ole.openstream(entry).read()
                if compressed:
                    try:
                        raw = zlib.decompress(raw, -15)
                    except Exception:
                        continue
                i, n = 0, len(raw)
                while i + 4 <= n:
                    hdr = struct.unpack_from("<I", raw, i)[0]
                    tag = hdr & 0x3FF
                    size = (hdr >> 20) & 0xFFF
                    i += 4
                    if size == 0xFFF:
                        if i + 4 > n:
                            break
                        size = struct.unpack_from("<I", raw, i)[0]
                        i += 4
                    if tag == 67 and i + size <= n:
                        try:
                            chunks.append(raw[i:i + size].decode("utf-16-le", "ignore"))
                        except Exception:
                            pass
                    i += size
        ole.close()
        return re.sub(r"[\x00-\x08\x0b-\x1f]", " ", "\n".join(chunks))
    except Exception:
        return ""


def extract_text(filename, data):
    low = filename.lower()
    if low.endswith(".pdf") or data[:4] == b"%PDF":
        return text_from_pdf(data)
    if low.endswith(".hwpx") or data[:2] == b"PK":
        t = text_from_hwpx(data)
        if t:
            return t
    return text_from_hwp(data)


DEPT_RE = re.compile(r"[가-힣A-Za-z0-9]{2,14}(?:과|팀|사업소|센터|본부|단|국|실|소)\b")
DEPT_STOP = ("결과", "효과", "통과", "초과", "부과", "경과", "성과")


def dept_of(ctx):
    """문맥에서 부서명(○○과, ○○팀 등) 추출 — 전화번호에 가까운 것 우선"""
    cands = [m.group() for m in DEPT_RE.finditer(ctx or "")
             if not m.group().endswith(DEPT_STOP)]
    return cands[-1] if cands else ""


def find_contacts(text):
    results = []
    if not text:
        return results
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    for idx, line in enumerate(lines):
        for m in PHONE_RE.finditer(line):
            phone = re.sub(r"[\s\.\)]", "-", m.group()).strip("-")
            phone = re.sub(r"-{2,}", "-", phone)
            ctx = " / ".join(lines[max(0, idx - 2): idx + 1])[-200:]
            score = sum(1 for kw in CONTACT_KEYWORDS if kw in ctx)
            if "팩스" in line or "FAX" in line.upper():
                score -= 3
            results.append((score, phone, ctx))
    results.sort(key=lambda x: -x[0])
    seen, out = set(), []
    for s, p, c in results:
        if p not in seen:
            seen.add(p)
            out.append((s, p, c))
    return out[:3]


def make_excel(headers, rows, sheet):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet
    ws.append(headers)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="D9E1F2")
    for r in rows:
        ws.append(r)
    for i, h in enumerate(headers, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = \
            max(12, min(55, len(str(h)) * 2 + 8))
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ------------------------------------------------------------
# 화면 — 검색조건
# ------------------------------------------------------------
st.title("📋 나라장터 입찰정보 수집기")
st.caption("입찰공고·개찰결과 조회 + 공고서 실무담당 연락처 자동 추출")

if get_service_key() is None:
    st.error("인증키가 설정되지 않았습니다. 앱 설정(Settings → Secrets)에 "
             "`SERVICE_KEY = \"발급받은키\"` 를 추가하세요.")
    st.stop()

search_type = st.radio("검색유형", ["입찰공고", "개찰결과"], horizontal=True)

today = datetime.date.today()
c1, c2 = st.columns(2)
with c1:
    date_from = st.date_input("조회 시작일", today - datetime.timedelta(days=14))
with c2:
    date_to = st.date_input("조회 종료일", today)

c3, c4 = st.columns(2)
with c3:
    task = st.selectbox("업무구분", list(TASKS.keys()))
with c4:
    region_name = st.selectbox("참가제한지역", list(REGIONS.keys()),
                               index=list(REGIONS.keys()).index("경남"))

c5, c6 = st.columns(2)
with c5:
    price_min_uk = st.number_input("추정가격 하한 (억원)", 0.0, 10000.0, 1.0, step=0.5)
with c6:
    price_max_uk = st.number_input("추정가격 상한 (억원, 0=제한없음)",
                                   0.0, 10000.0, 0.0, step=0.5)

if search_type == "입찰공고":
    do_contacts = st.checkbox("공고서에서 실무담당 전화번호 추출", value=True)
    max_attach = st.slider("공고당 첨부파일 분석 개수", 1, 5, 3,
                           help="많을수록 정확하지만 느려집니다") if do_contacts else 0
else:
    do_contacts, max_attach = False, 0
    st.caption("※ 개찰결과는 낙찰업체·낙찰금액 중심으로 조회됩니다. "
               "'나라장터 낙찰정보서비스' 활용신청이 된 인증키여야 합니다.")

# ------------------------------------------------------------
# 실행
# ------------------------------------------------------------
if st.button("🔍 조회 시작", type="primary", use_container_width=True):
    if date_from > date_to:
        st.error("조회 시작일이 종료일보다 늦습니다.")
        st.stop()

    begin = date_from.strftime("%Y%m%d") + "0000"
    end = date_to.strftime("%Y%m%d") + "2359"
    p_min = int(price_min_uk * 100_000_000)
    p_max = int(price_max_uk * 100_000_000)  # 0이면 제한없음
    region_code = REGIONS[region_name]

    extra = {"inqryDiv": 1, "inqryBgnDt": begin, "inqryEndDt": end}
    log = st.empty()

    if search_type == "입찰공고":
        if region_code:
            extra["prtcptLmtRgnCd"] = region_code
        op = TASKS[task][0]
        with st.spinner("입찰공고 목록 조회 중..."):
            items, err = call_api(BID_BASES, op, extra, log)
    else:
        op = TASKS[task][1]
        with st.spinner("개찰결과 목록 조회 중..."):
            items, err = call_api(SCSBID_BASES, op, extra, log)
        # 개찰결과는 지역 파라미터 미지원 → 기관명 텍스트로 2차 필터
        if items and region_code:
            kw = region_name
            items = [it for it in items if kw in
                     (str(it.get("dminsttNm", "")) + str(it.get("ntceInsttNm", "")))]

    log.empty()
    if err and not items:
        st.error(f"API 조회 실패: {err}")
        if search_type == "개찰결과":
            st.info("공공데이터포털에서 '조달청_나라장터 낙찰정보서비스' "
                    "활용신청이 승인되었는지 확인하세요.")
        else:
            st.info("인증키 활성화 전이거나(발급 후 몇 시간 소요), "
                    "'나라장터 입찰공고정보서비스' 활용신청을 확인하세요.")
        st.stop()

    # 금액 필터
    filtered = [it for it in items
                if price_of(it) >= p_min and (p_max == 0 or price_of(it) <= p_max)]
    st.success(f"조건 충족 {len(filtered)}건 (전체 수신 {len(items)}건)")
    if not filtered:
        st.stop()

    fname_date = today.strftime("%Y%m%d")

    # ---------------- 입찰공고 결과 ----------------
    if search_type == "입찰공고":
        rows = []
        progress = st.progress(0.0)
        status = st.empty()
        for i, item in enumerate(filtered, 1):
            name = item.get("bidNtceNm", "")
            best_phone, best_ctx, src_file = "", "", ""
            if do_contacts:
                status.write(f"[{i}/{len(filtered)}] {name[:35]}... 공고서 분석 중")
                for att_name, att_url in attachments_of(item, max_attach):
                    try:
                        r = requests.get(att_url, headers=HEADERS, timeout=60)
                        if r.status_code != 200 or len(r.content) < 500:
                            continue
                        contacts = find_contacts(extract_text(att_name, r.content))
                        if contacts and contacts[0][0] > 0:
                            best_phone, best_ctx, src_file = \
                                contacts[0][1], contacts[0][2], att_name
                            break
                        elif contacts and not best_phone:
                            best_phone, best_ctx, src_file = \
                                contacts[0][1], contacts[0][2], att_name
                    except Exception:
                        continue
                    time.sleep(0.2)
            rows.append([item.get("bidNtceNo", ""), item.get("bidNtceOrd", ""),
                         name, item.get("dminsttNm", ""), price_of(item),
                         dept_of(best_ctx), best_phone, best_ctx, src_file,
                         item.get("ntceInsttOfclNm", ""),
                         item.get("ntceInsttOfclTelNo", ""),
                         item.get("bidNtceDtlUrl") or item.get("bidNtceUrl") or "",
                         "추출성공" if best_phone else
                         ("수동확인필요" if do_contacts else "-")])
            progress.progress(i / len(filtered))
        status.empty()

        if do_contacts:
            ok = sum(1 for r in rows if r[-1] == "추출성공")
            st.success(f"완료! 자동 추출 {ok}건 / 수동확인 필요 {len(rows) - ok}건")

        st.dataframe(
            [{"공고명": r[2][:30], "수요기관": r[3],
              "실무담당 부서": r[5], "실무담당 전화": r[6],
              "집행관(계약) 전화": r[10],
              "추정가격(억)": round(r[4] / 1e8, 2),
              "공고 바로가기": r[11], "상태": r[-1]} for r in rows],
            use_container_width=True, hide_index=True,
            column_config={"공고 바로가기": st.column_config.LinkColumn(
                "공고 바로가기", display_text="열기")})

        headers = ["공고번호", "차수", "공고명", "수요기관", "추정가격(원)",
                   "실무담당 부서", "실무담당 전화", "추출 문맥", "출처 파일",
                   "집행관(계약담당)", "집행관 전화", "공고 상세URL", "상태"]
        st.download_button("📥 엑셀 다운로드",
                           data=make_excel(headers, rows, "입찰공고"),
                           file_name=f"입찰공고_{task}_{region_name}_{fname_date}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument"
                                ".spreadsheetml.sheet",
                           use_container_width=True)

    # ---------------- 개찰결과 결과 ----------------
    else:
        rows = [[it.get("bidNtceNo", ""), it.get("bidNtceNm", "") or
                 it.get("prdctClsfcNoNm", ""),
                 it.get("dminsttNm", "") or it.get("ntceInsttNm", ""),
                 it.get("opengDt", "") or it.get("rlOpengDt", ""),
                 it.get("bidwinnrNm", "") or it.get("opengCorpInfo", ""),
                 price_of(it),
                 it.get("sucsfbidRate", "")] for it in filtered]

        st.dataframe(
            [{"공고명": r[1][:35], "수요기관": r[2], "개찰일": str(r[3])[:10],
              "낙찰업체": r[4], "낙찰금액(억)": round(r[5] / 1e8, 2)}
             for r in rows],
            use_container_width=True, hide_index=True)

        headers = ["공고번호", "공고명", "수요기관", "개찰일시",
                   "낙찰업체", "낙찰금액(원)", "낙찰률(%)"]
        st.download_button("📥 엑셀 다운로드",
                           data=make_excel(headers, rows, "개찰결과"),
                           file_name=f"개찰결과_{task}_{region_name}_{fname_date}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument"
                                ".spreadsheetml.sheet",
                           use_container_width=True)
