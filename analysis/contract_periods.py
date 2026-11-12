# -*- coding: utf-8 -*-
"""
나라장터 계약정보(공사)로 공사금액 구간별 공사기간·법정 지도횟수 분포 계산
PROBE=1: 주소·오퍼레이션 조합을 몇 번만 시험해 응답 구조(착공일·준공일 항목) 확인
필요 환경변수: SERVICE_KEY
"""
import os
import sys
import json
import time
import datetime
import urllib.parse

import requests

RAW = os.environ.get("SERVICE_KEY", "").strip()
KEY = urllib.parse.unquote(RAW) if "%" in RAW else RAW
HEADERS = {"User-Agent": "Mozilla/5.0"}
BASES = [
    "https://apis.data.go.kr/1230000/ao/CntrctInfoService",  # 01은 폐기 확인됨
]
OPS = ["getCntrctInfoListCnstwk", "getCntrctInfoListCnstwkPPSSrch"]
CALLS = [0]


def call(base, op, params):
    CALLS[0] += 1
    q = {"serviceKey": KEY, "type": "json", "numOfRows": 5, "pageNo": 1,
         **params}
    try:
        r = requests.get(f"{base}/{op}", params=q, headers=HEADERS, timeout=60)
    except Exception as e:
        return None, f"접속오류 {e}"
    txt = r.text.strip()
    try:
        return json.loads(txt), None
    except ValueError:
        return None, f"HTTP {r.status_code} 비JSON: {txt[:300]}"


def items_of(js):
    resp = js.get("response", js)
    body = resp.get("body", {}) or {}
    rows = body.get("items", [])
    if isinstance(rows, dict):
        rows = rows.get("item", [])
    if isinstance(rows, dict):
        rows = [rows]
    return body.get("totalCount"), rows or []


def probe():
    today = datetime.date.today()
    y = today - datetime.timedelta(days=1)
    p = {"inqryDiv": 1, "inqryBgnDt": y.strftime("%Y%m%d") + "0000",
         "inqryEndDt": y.strftime("%Y%m%d") + "2359"}
    for base in BASES:
        for op in OPS:
            js, err = call(base, op, p)
            tag = f"{base.split('/')[-1]}/{op}"
            if err:
                print(f"[{tag}] {err}")
                continue
            total, rows = items_of(js)
            if not rows:
                print(f"[{tag}] 0건, 원문: "
                      f"{json.dumps(js, ensure_ascii=False)[:400]}")
                continue
            print(f"[{tag}] 성공: 전체 {total}건")
            print("  필드:", ", ".join(
                f"{k}={str(v)[:25]}" for k, v in sorted(rows[0].items())))
            time.sleep(0.3)
    print(f"PROBE 종료: 호출 {CALLS[0]}회")


def main():
    if not KEY:
        print("오류: SERVICE_KEY 없음")
        sys.exit(1)
    if os.environ.get("PROBE") == "1":
        probe()
        return
    print("본 집계는 시험 결과 확인 후 추가 예정")


if __name__ == "__main__":
    main()
