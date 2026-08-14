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
# 참가제한지역 (나라장터 검색화면과 동일한 목록, 2026 행정구역 반영)
# 값: 참가제한지역 코드 목록. 행정구역 개편(통합·승격) 전 공고까지 잡히도록
#     구 코드를 함께 조회한다. None이면 지역 필터 없음.
REGIONS = {
    "전체": None,
    "전국(제한없음)": [],            # 코드 필터 없이 조회 후 지역명으로 판별
    "서울특별시": ["11"],
    "전남광주통합특별시": ["46", "29"],   # 통합 전 전남(46)·광주(29) 코드 포함
    "부산광역시": ["26"],
    "대구광역시": ["27"],
    "인천광역시": ["28"],
    "대전광역시": ["30"],
    "울산광역시": ["31"],
    "세종특별자치시": ["36"],
    "경기도": ["41"],
    "충청북도": ["43"],
    "충청남도": ["44"],
    "경상북도": ["47"],
    "경상남도": ["48"],
    "제주특별자치도": ["50"],
    "강원특별자치도": ["51", "42"],      # 특별자치도 승격 전 코드(42) 포함
    "전북특별자치도": ["52", "45"],      # 특별자치도 승격 전 코드(45) 포함
}

# 공고의 '참가가능지역명'과 매칭하기 위한 지역명 키워드
RGN_NAME_KEYWORDS = {
    "전국(제한없음)": ["전국", "제한없음"],
    "서울특별시": ["서울"],
    "전남광주통합특별시": ["전남", "전라남도", "광주"],
    "부산광역시": ["부산"],
    "대구광역시": ["대구"],
    "인천광역시": ["인천"],
    "대전광역시": ["대전"],
    "울산광역시": ["울산"],
    "세종특별자치시": ["세종"],
    "경기도": ["경기"],
    "충청북도": ["충북", "충청북도"],
    "충청남도": ["충남", "충청남도"],
    "경상북도": ["경북", "경상북도"],
    "경상남도": ["경남", "경상남도"],
    "제주특별자치도": ["제주"],
    "강원특별자치도": ["강원"],
    "전북특별자치도": ["전북", "전라북도"],
}

# 참가가능지역 정보를 조회할 수 없는 공고용 예비 필터:
# 수요기관명에 지역 소속 시·군 이름이 있으면 해당 지역으로 간주
INSTT_KEYWORDS = {
    "서울특별시": ["서울"],
    "부산광역시": ["부산"],
    "대구광역시": ["대구", "군위"],
    "인천광역시": ["인천", "강화", "옹진"],
    "대전광역시": ["대전"],
    "울산광역시": ["울산"],
    "세종특별자치시": ["세종"],
    "경기도": ["경기", "수원", "성남", "의정부", "안양", "부천", "광명", "평택",
             "동두천", "안산", "고양", "과천", "구리", "남양주", "오산", "시흥",
             "군포", "의왕", "하남", "용인", "파주", "이천", "안성", "김포",
             "화성", "광주시", "양주", "포천", "여주", "연천", "가평", "양평"],
    "강원특별자치도": ["강원", "춘천", "원주", "강릉", "동해", "태백", "속초",
             "삼척", "홍천", "횡성", "영월", "평창", "정선", "철원", "화천",
             "양구", "인제", "고성", "양양"],
    "충청북도": ["충북", "충청북도", "청주", "충주", "제천", "보은", "옥천",
             "영동", "증평", "진천", "괴산", "음성", "단양"],
    "충청남도": ["충남", "충청남도", "천안", "공주", "보령", "아산", "서산",
             "논산", "계룡", "당진", "금산", "부여", "서천", "청양", "홍성",
             "예산", "태안"],
    "전북특별자치도": ["전북", "전라북도", "전주", "군산", "익산", "정읍", "남원",
             "김제", "완주", "진안", "무주", "장수", "임실", "순창", "고창", "부안"],
    "전남광주통합특별시": ["전남", "전라남도", "광주", "목포", "여수", "순천",
             "나주", "광양", "담양", "곡성", "구례", "고흥", "보성", "화순",
             "장흥", "강진", "해남", "영암", "무안", "함평", "영광", "장성",
             "완도", "진도", "신안"],
    "경상북도": ["경북", "경상북도", "포항", "경주", "김천", "안동", "구미",
             "영주", "영천", "상주", "문경", "경산", "의성", "청송", "영양",
             "영덕", "청도", "고령", "성주", "칠곡", "예천", "봉화", "울진", "울릉"],
    "경상남도": ["경남", "경상남도", "창원", "진주", "통영", "사천", "김해",
             "밀양", "거제", "양산", "의령", "함안", "창녕", "고성", "남해",
             "하동", "산청", "함양", "거창", "합천"],
    "제주특별자치도": ["제주", "서귀포"],
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


def fetch_notice(bid_ntce_no, task, cache={}):
    """공고번호로 원 입찰공고 1건 조회 (개찰결과 → 담당자 연락처 추출용)"""
    if not bid_ntce_no:
        return None
    if bid_ntce_no in cache:
        return cache[bid_ntce_no]
    op = TASKS[task][0]
    key = get_service_key()
    item = None
    for base in BID_BASES:
        try:
            r = requests.get(f"{base}/{op}",
                             params={"serviceKey": key, "pageNo": 1,
                                     "numOfRows": 10, "type": "json",
                                     "inqryDiv": 2, "bidNtceNo": bid_ntce_no},
                             headers=HEADERS, timeout=30)
            if r.status_code != 200 or r.text.lstrip().startswith("<"):
                continue
            body = r.json().get("response", {}).get("body", {})
            rows = body.get("items", [])
            if isinstance(rows, dict):
                rows = rows.get("item", [])
            if rows:
                item = rows[-1]  # 재공고 등 차수가 여러 개면 최신 차수
                break
        except Exception:
            continue
    cache[bid_ntce_no] = item
    return item


def fetch_psbl_rgn(bid_ntce_no, cache={}):
    """공고번호로 참가가능(제한)지역명 목록 조회.
    반환: 지역명 리스트 / [] (정상응답이며 제한 미지정) / None (조회 불가)"""
    if not bid_ntce_no:
        return None
    if bid_ntce_no in cache:
        return cache[bid_ntce_no]
    key = get_service_key()
    result = None
    for base in BID_BASES:
        try:
            r = requests.get(f"{base}/getBidPblancListInfoPrtcptPsblRgn",
                             params={"serviceKey": key, "pageNo": 1,
                                     "numOfRows": 30, "type": "json",
                                     "inqryDiv": 2, "bidNtceNo": bid_ntce_no},
                             headers=HEADERS, timeout=30)
            if r.status_code != 200 or r.text.lstrip().startswith("<"):
                continue
            resp = r.json().get("response", {})
            code = str(resp.get("header", {}).get("resultCode", ""))
            if code not in ("00", "0"):     # 오류 응답은 '조회 불가'로 처리
                continue
            body = resp.get("body", {})
            rows = body.get("items", [])
            if isinstance(rows, dict):
                rows = rows.get("item", [])
            if isinstance(rows, dict):
                rows = [rows]
            names = []
            for it in rows:
                if not isinstance(it, dict):
                    continue
                for f in ("prtcptPsblRgnNm", "prtcptLmtRgnNm", "rgnNm",
                          "prtcptPsblRgnCdNm"):
                    v = str(it.get(f, "") or "").strip()
                    if v:
                        names.append(v)
            # 정상응답이지만 행이 0건이면 totalCount로 진위 확인
            total = str(body.get("totalCount", "0"))
            if not names and total not in ("0", ""):
                result = None       # 행은 있다는데 지역명을 못 읽음 → 판단 보류
            else:
                result = names      # []: 제한 미지정(전국) 확정
            break
        except Exception:
            continue
    cache[bid_ntce_no] = result
    return result


def region_match(region_name, rgn_names, item):
    """나라장터 참가제한지역 필터와 동일한 판정.
    rgn_names: 공고의 참가가능지역명 목록(None=조회실패, []=제한없음)"""
    if region_name == "전체":
        return True
    if rgn_names is None:
        # 지역정보 조회 실패 → 수요기관명으로 예비 판정
        if region_name == "전국(제한없음)":
            return False
        kws = INSTT_KEYWORDS.get(region_name, [region_name])
        blob = str(item.get("dminsttNm", "")) + str(item.get("ntceInsttNm", ""))
        return any(k in blob for k in kws)
    if not rgn_names:          # 제한지역 미지정 = 전국(제한없음)
        return region_name == "전국(제한없음)"
    joined = " ".join(rgn_names)
    if region_name == "전국(제한없음)":
        return any(k in joined for k in RGN_NAME_KEYWORDS["전국(제한없음)"])
    return any(k in joined for k in RGN_NAME_KEYWORDS.get(region_name, [region_name]))


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
date_basis = "개찰일" if search_type == "개찰결과" else "공고게시일"
c1, c2 = st.columns(2)
with c1:
    date_from = st.date_input(f"조회 시작일 ({date_basis} 기준)",
                              today - datetime.timedelta(days=14))
with c2:
    date_to = st.date_input(f"조회 종료일 ({date_basis} 기준)", today)

c3, c4 = st.columns(2)
with c3:
    task = st.selectbox("업무구분", list(TASKS.keys()))
with c4:
    region_name = st.selectbox("참가제한지역", list(REGIONS.keys()),
                               index=list(REGIONS.keys()).index("경상남도"))
    include_nationwide = False
    if region_name not in ("전체", "전국(제한없음)"):
        include_nationwide = st.checkbox(
            "전국(제한없음) 공고도 수요기관이 해당 지역이면 포함 (개찰결과)",
            value=True,
            help="나라장터 참가제한지역 필터는 전국(제한없음) 공고를 제외하지만, "
                 "체크하면 수요기관이 선택 지역 소속인 전국 공고도 함께 잡습니다.")

c5, c6 = st.columns(2)
with c5:
    price_min_uk = st.number_input("추정가격 하한 (억원)", 0.0, 10000.0, 1.0, step=0.5)
with c6:
    price_max_uk = st.number_input("추정가격 상한 (억원, 0=제한없음)",
                                   0.0, 10000.0, 0.0, step=0.5)

do_contacts = st.checkbox("공고서에서 실무담당 전화번호 추출", value=True)
max_attach = st.slider("공고당 첨부파일 분석 개수", 1, 5, 3,
                       help="많을수록 정확하지만 느려집니다") if do_contacts else 0
if search_type == "개찰결과":
    st.caption("※ 개찰결과는 공고번호로 원 입찰공고를 찾아 담당자 연락처를 함께 추출합니다. "
               "'나라장터 낙찰정보서비스' 활용신청이 된 인증키여야 하며, "
               "건마다 공고 조회가 추가되어 시간이 더 걸립니다.")

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
    region_codes = REGIONS[region_name]

    extra = {"inqryDiv": 1, "inqryBgnDt": begin, "inqryEndDt": end}
    log = st.empty()

    if search_type == "입찰공고":
        op = TASKS[task][0]
        with st.spinner("입찰공고 목록 조회 중..."):
            if region_codes:   # 지역코드별로 조회 후 병합 (통합·개편 구코드 포함)
                items, err, seen = [], None, set()
                for code in region_codes:
                    part, e = call_api(BID_BASES, op,
                                       {**extra, "prtcptLmtRgnCd": code}, log)
                    err = err or e
                    for it in part:
                        k = (it.get("bidNtceNo"), it.get("bidNtceOrd"))
                        if k not in seen:
                            seen.add(k)
                            items.append(it)
            else:              # 전체 / 전국(제한없음): 코드 필터 없이 조회
                items, err = call_api(BID_BASES, op, extra, log)
        if items and region_name == "전국(제한없음)":
            # 참가제한이 걸리지 않은(전국) 공고만 남김
            items = [it for it in items
                     if region_match(region_name,
                                     fetch_psbl_rgn(it.get("bidNtceNo")), it)]
    else:
        op = TASKS[task][1]
        # 낙찰정보 API의 기간 조회는 '등록일시' 기준이라 나라장터 화면의
        # '개찰일자' 기준과 어긋난다 → 넉넉한 기간으로 받은 뒤 개찰일로 필터.
        wide = {"inqryDiv": 1,
                "inqryBgnDt": (date_from - datetime.timedelta(days=3)
                               ).strftime("%Y%m%d") + "0000",
                "inqryEndDt": (date_to + datetime.timedelta(days=2)
                               ).strftime("%Y%m%d") + "2359"}
        with st.spinner("개찰결과 목록 조회 중..."):
            items, err = call_api(SCSBID_BASES, op, wide, log)

        def _openg_date(it):
            s = str(it.get("opengDt", "") or it.get("rlOpengDt", ""))[:10]
            s = s.replace(".", "-").replace("/", "-")
            try:
                return datetime.date.fromisoformat(s)
            except ValueError:
                return None
        if items:
            n_all = len(items)
            items = [it for it in items
                     if (_openg_date(it) is None)          # 개찰일 판독불가 건은 유지
                     or (date_from <= _openg_date(it) <= date_to)]
            if n_all != len(items):
                st.caption(f"개찰일 {date_from}~{date_to} 범위 밖 "
                           f"{n_all - len(items)}건 제외")
        # 지역 필터는 아래 루프에서 공고번호로 '참가가능지역'을 조회해
        # 나라장터 참가제한지역 필터와 동일한 기준으로 적용한다.

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
    # ※ 개찰결과 목록 API에는 추정가격 필드가 없으므로(모두 0으로 계산됨)
    #   목록 단계에서는 필터하지 않고, 아래 루프에서 원 공고의 추정가격으로 필터한다.
    if search_type == "입찰공고":
        filtered = [it for it in items
                    if price_of(it) >= p_min and (p_max == 0 or price_of(it) <= p_max)]
        st.success(f"조건 충족 {len(filtered)}건 (전체 수신 {len(items)}건)")
    else:
        filtered = items
        st.success(f"개찰일 기준 수신 {len(items)}건 — 건별로 참가제한지역·"
                   f"추정가격을 확인해 필터를 적용합니다")
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
        # ⚡ 속도 개선: 건별 지역확인 대신, 선택 지역으로 제한된 입찰공고
        # 목록(최근 120일)을 한 번에 받아 공고번호로 대조한다.
        # 대조된 건은 원 공고 정보(추정가격·담당자)도 함께 확보되어
        # 공고 재조회가 필요 없다.
        notice_map = {}
        if region_name not in ("전체", "전국(제한없음)") and region_codes:
            nb = (date_from - datetime.timedelta(days=120)
                  ).strftime("%Y%m%d") + "0000"
            ne = date_to.strftime("%Y%m%d") + "2359"
            with st.spinner(f"{region_name} 제한 공고 목록 대조용 조회 중..."):
                for code in region_codes:
                    part, _ = call_api(
                        BID_BASES, TASKS[task][0],
                        {"inqryDiv": 1, "inqryBgnDt": nb, "inqryEndDt": ne,
                         "prtcptLmtRgnCd": code}, log)
                    for n in part:
                        notice_map[str(n.get("bidNtceNo", ""))] = n
            log.empty()
            st.caption(f"{region_name} 제한 공고 {len(notice_map)}건과 대조합니다")

        rows = []
        progress = st.progress(0.0)
        status = st.empty()
        stats = {"지역 불일치": 0, "금액 범위 밖": 0, "지역정보 미확인(기관명 판정)": 0}
        for i, it in enumerate(filtered, 1):
            no = str(it.get("bidNtceNo", ""))
            name = it.get("bidNtceNm", "") or it.get("prdctClsfcNoNm", "")
            best_phone, best_ctx, src_file = "", "", ""
            ofcl_nm, ofcl_tel, url = "", "", ""
            notice = None
            # ① 참가제한지역 확인 (나라장터 필터와 동일 기준)
            if region_name == "전국(제한없음)":
                status.write(f"[{i}/{len(filtered)}] {name[:35]}... 참가제한지역 확인 중")
                rgn = fetch_psbl_rgn(no)
                if rgn is None:
                    stats["지역정보 미확인(기관명 판정)"] += 1
                if not region_match(region_name, rgn, it):
                    stats["지역 불일치"] += 1
                    progress.progress(i / len(filtered))
                    continue
            elif region_name != "전체":
                if no in notice_map:
                    notice = notice_map[no]     # 지역제한 일치 확정 + 공고 확보
                else:
                    # 대조 실패 → 수요기관명이 지역 소속인 후보만 정밀 확인
                    kws = INSTT_KEYWORDS.get(region_name, [region_name])
                    blob = (str(it.get("dminsttNm", "")) +
                            str(it.get("ntceInsttNm", "")))
                    if not any(k in blob for k in kws):
                        stats["지역 불일치"] += 1
                        progress.progress(i / len(filtered))
                        continue
                    status.write(f"[{i}/{len(filtered)}] {name[:35]}... 참가제한지역 확인 중")
                    rgn = fetch_psbl_rgn(no)
                    if rgn:          # 지역제한 명시 → 지역명으로 판정
                        ok = any(k in " ".join(rgn)
                                 for k in RGN_NAME_KEYWORDS.get(region_name,
                                                                [region_name]))
                    elif rgn == []:  # 전국(제한없음) 공고
                        ok = include_nationwide
                    else:            # 확인 불가 → 기관명이 일치하므로 포함
                        ok = True
                        stats["지역정보 미확인(기관명 판정)"] += 1
                    if not ok:
                        stats["지역 불일치"] += 1
                        progress.progress(i / len(filtered))
                        continue
            # ② 원 공고 확보 → 추정가격 필터 + 담당자 정보
            if notice is None:
                status.write(f"[{i}/{len(filtered)}] {name[:35]}... 원 공고 조회 중")
                notice = fetch_notice(no, task)
            # 원 공고의 추정가격으로 금액 필터 (공고 미확인 건은 일단 포함)
            prc = price_of(notice) if notice else 0
            if notice and (prc < p_min or (p_max and prc > p_max)):
                stats["금액 범위 밖"] += 1
                progress.progress(i / len(filtered))
                continue
            if notice:
                ofcl_nm = notice.get("ntceInsttOfclNm", "")
                ofcl_tel = notice.get("ntceInsttOfclTelNo", "")
                url = notice.get("bidNtceDtlUrl") or notice.get("bidNtceUrl") or ""
                if do_contacts:
                    status.write(f"[{i}/{len(filtered)}] {name[:35]}... 공고서 분석 중")
                    for att_name, att_url in attachments_of(notice, max_attach):
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
            state = ("추출성공" if best_phone else
                     ("공고미확인" if not notice else
                      ("수동확인필요" if do_contacts else "-")))
            rows.append([no, name,
                         it.get("dminsttNm", "") or it.get("ntceInsttNm", ""),
                         it.get("opengDt", "") or it.get("rlOpengDt", ""),
                         it.get("bidwinnrNm", "") or it.get("opengCorpInfo", ""),
                         prc, it.get("sucsfbidRate", ""),
                         dept_of(best_ctx), best_phone, best_ctx, src_file,
                         ofcl_nm, ofcl_tel, url, state])
            progress.progress(i / len(filtered))
            time.sleep(0.2)
        status.empty()

        st.success(f"지역·금액 조건 충족 {len(rows)}건 (수신 {len(filtered)}건 중)")
        st.caption(" · ".join(f"{k} {v}건" for k, v in stats.items() if v))
        if not rows:
            st.stop()
        if do_contacts:
            ok = sum(1 for r in rows if r[-1] == "추출성공")
            st.success(f"자동 추출 {ok}건 / 수동확인 필요 {len(rows) - ok}건")

        st.dataframe(
            [{"공고명": r[1][:30], "수요기관": r[2], "개찰일": str(r[3])[:10],
              "낙찰업체": r[4], "추정가격(억)": round(r[5] / 1e8, 2),
              "실무담당 부서": r[7], "실무담당 전화": r[8],
              "집행관(계약) 전화": r[12],
              "공고 바로가기": r[13], "상태": r[-1]} for r in rows],
            use_container_width=True, hide_index=True,
            column_config={"공고 바로가기": st.column_config.LinkColumn(
                "공고 바로가기", display_text="열기")})

        headers = ["공고번호", "공고명", "수요기관", "개찰일시",
                   "낙찰업체", "추정가격(원)", "낙찰률(%)",
                   "실무담당 부서", "실무담당 전화", "추출 문맥", "출처 파일",
                   "집행관(계약담당)", "집행관 전화", "공고 상세URL", "상태"]
        st.download_button("📥 엑셀 다운로드",
                           data=make_excel(headers, rows, "개찰결과"),
                           file_name=f"개찰결과_{task}_{region_name}_{fname_date}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument"
                                ".spreadsheetml.sheet",
                           use_container_width=True)
