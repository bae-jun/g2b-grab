# -*- coding: utf-8 -*-
"""
KOSIS 건축 착공 통계로 부산·울산·경남 시·군·구별 민간 포함 착공 규모 집계
1단계(TBL_ID 미지정): '착공' 관련 통계표를 검색해 후보 표와 항목 구조를 출력
2단계(TBL_ID, ORG_ID 지정): 해당 표에서 최근 12개월 자료를 받아 지역별 합계 출력
필요 환경변수: KOSIS_KEY
"""
import os
import sys
import json
import datetime
from collections import defaultdict

import requests

KEY = os.environ.get("KOSIS_KEY", "").strip()
BASE = "https://kosis.kr/openapi"
REGION_WORDS = ("부산", "울산", "경상남도", "경남", "창원", "진주", "통영",
                "사천", "김해", "밀양", "거제", "양산", "의령", "함안", "창녕",
                "고성", "남해", "하동", "산청", "함양", "거창", "합천")


def get(path, params):
    p = {"apiKey": KEY, "format": "json", "jsonVD": "Y", **params}
    r = requests.get(f"{BASE}/{path}", params=p, timeout=60)
    txt = r.text.strip()
    try:
        return json.loads(txt)
    except ValueError:
        print(f"  [응답 해석 실패] {path} HTTP {r.status_code}: {txt[:300]}")
        return None


def search_tables():
    found = {}
    for word in ("시군구별 건축 착공", "건축착공", "착공현황"):
        res = get("statisticsSearch.do", {"method": "getList",
                                          "searchNm": word,
                                          "resultCount": 50})
        if isinstance(res, dict):
            print(f"  [검색 '{word}'] 응답: {str(res)[:300]}")
            continue
        for row in res or []:
            k = (row.get("ORG_ID"), row.get("TBL_ID"))
            found[k] = row
    print(f"후보 통계표 {len(found)}개")
    for (org, tbl), row in found.items():
        print(f"TABLE | {org} | {tbl} | {row.get('TBL_NM')} | "
              f"{row.get('ORG_NM')} | 주기:{row.get('PRD_SE', '')} | "
              f"수록:{row.get('STRT_PRD_DE', '')}~{row.get('END_PRD_DE', '')}")
    # 이름에 '착공'과 '시군구'가 모두 들어간 표는 분류항목도 출력
    for (org, tbl), row in found.items():
        nm = str(row.get("TBL_NM", ""))
        if "착공" in nm and ("시군구" in nm or "시군" in nm):
            meta = get("statisticsData.do", {"method": "getMeta",
                                             "type": "ITM", "orgId": org,
                                             "tblId": tbl})
            if isinstance(meta, list):
                print(f"META | {tbl} | 항목 {len(meta)}개")
                for m in meta[:60]:
                    print(f"META | {tbl} | {m.get('OBJ_ID')} | "
                          f"{m.get('OBJ_NM')} | {m.get('ITM_ID')} | "
                          f"{m.get('ITM_NM')} | {m.get('UNIT_NM', '')}")
            else:
                print(f"META | {tbl} | {str(meta)[:200]}")


def fetch_table(org, tbl):
    today = datetime.date.today()
    end = today.replace(day=1) - datetime.timedelta(days=1)
    start = (end.replace(day=1) - datetime.timedelta(days=330)).replace(day=1)
    params = {"method": "getList", "orgId": org, "tblId": tbl,
              "itmId": os.environ.get("ITM_ID", "ALL"),
              "objL1": os.environ.get("OBJ_L1", "ALL"),
              "prdSe": os.environ.get("PRD_SE", "M"),
              "startPrdDe": os.environ.get("START", start.strftime("%Y%m")),
              "endPrdDe": os.environ.get("END", end.strftime("%Y%m"))}
    for lv in ("OBJ_L2", "OBJ_L3", "OBJ_L4"):
        if os.environ.get(lv):
            params["objL" + lv[-1]] = os.environ[lv]
    rows = get("Param/statisticsParameterData.do", params)
    if not isinstance(rows, list):
        print(f"자료 조회 실패: {str(rows)[:300]}")
        sys.exit(1)
    print(f"수신 {len(rows)}행 (기간 {params['startPrdDe']}~{params['endPrdDe']})")
    if rows:
        print("행 예시:", {k: v for k, v in rows[0].items()})
    agg = defaultdict(float)
    for r in rows:
        names = " ".join(str(r.get(f"C{i}_NM", "") or "") for i in range(1, 5))
        if not any(w in names for w in REGION_WORDS):
            continue
        try:
            v = float(str(r.get("DT", "0")).replace(",", "") or 0)
        except ValueError:
            continue
        key = (names.strip(), str(r.get("ITM_NM", "")),
               str(r.get("UNIT_NM", "")))
        agg[key] += v
    for (names, itm, unit), v in sorted(agg.items()):
        print(f"KOSIS | {names} | {itm} | {v:,.0f} | {unit}")


def show_meta(org, tbl):
    """표의 분류(OBJ)·항목(ITM) 구조와 최근 1개월 표본 출력"""
    meta = get("statisticsData.do", {"method": "getMeta", "type": "ITM",
                                     "orgId": org, "tblId": tbl})
    if not isinstance(meta, list):
        print(f"META | {tbl} | {str(meta)[:300]}")
        return
    objs = defaultdict(list)
    for m in meta:
        objs[(m.get("OBJ_ID"), m.get("OBJ_NM"))].append(
            f"{m.get('ITM_ID')}={m.get('ITM_NM')}")
    for (oid, onm), items in objs.items():
        sample = [x for x in items if any(w in x for w in REGION_WORDS)]
        print(f"META | {tbl} | {oid} | {onm} | 항목 {len(items)}개 | "
              f"앞부분: {items[:8]} | 지역관련: {sample[:40]}")


def main():
    if not KEY:
        print("오류: KOSIS_KEY가 GitHub Secret에 등록되지 않았습니다.")
        sys.exit(1)
    if os.environ.get("META_TBLS"):
        for t in os.environ["META_TBLS"].split(","):
            org, tbl = t.split(":")
            show_meta(org, tbl)
        return
    tbl = os.environ.get("TBL_ID")
    if not tbl:
        search_tables()
    else:
        fetch_table(os.environ.get("ORG_ID", "116"), tbl)


if __name__ == "__main__":
    main()
