# -*- coding: utf-8 -*-
"""
나라장터 개찰결과 일일 자동 리포트
- app.py의 파이프라인을 그대로 실행(중복 구현 없음)해 엑셀 생성 후 메일 발송
- GitHub Actions 등에서 매일 실행하는 용도. 설정은 환경변수로 주입:

  필수:  SERVICE_KEY            나라장터 공공데이터 인증키
  선택:  ANTHROPIC_API_KEY      Claude AI 정밀 추출 사용 시
  검색:  SEARCH_TYPE(개찰결과) TASK(공사) REGION(경상남도)
         PRICE_MIN_UK(1.0) PRICE_MAX_UK(0) DAYS_BACK(1: 어제~오늘)
         INCLUDE_NATIONWIDE(0/1) MAX_ATTACH(4)
  메일:  SMTP_HOST SMTP_PORT SMTP_USER SMTP_PASS MAIL_TO [MAIL_FROM]
         (미설정 시 메일 없이 out/ 폴더에 파일만 저장)
"""
import os
import re
import sys
import ssl
import glob
import types
import smtplib
import datetime
import contextlib
from email.message import EmailMessage

OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "out")
os.makedirs(OUT_DIR, exist_ok=True)


# ------------------------------------------------------------
# 1) 헤드리스 streamlit 대역: app.py를 UI 없이 실행하기 위한 최소 구현
# ------------------------------------------------------------
def build_stub(widgets, secrets):
    stub = types.ModuleType("streamlit")
    saved = {"files": [], "logs": []}

    def _val(label, default):
        for k, v in widgets.items():
            if k in str(label):
                return v
        return default

    def _log(kind, msg=""):
        line = f"[{kind}] {msg}"
        saved["logs"].append(line)
        print(line, flush=True)

    class Stop(Exception):
        pass

    class Ctx:
        def __enter__(self): return self
        def __exit__(self, *a): return False
        def write(self, t): print(f"  {t}", flush=True)
        def empty(self): pass
        def progress(self, v): pass

    class AnyAttr:
        def __getattr__(self, n): return lambda *a, **k: None

    stub.secrets = secrets
    stub.StopException = Stop
    stub.set_page_config = lambda **k: None
    stub.title = lambda t: _log("제목", t)
    stub.caption = lambda t: _log("안내", t)
    stub.error = lambda t: _log("오류", t)
    stub.info = lambda t: _log("정보", t)
    stub.warning = lambda t: _log("경고", t)
    stub.success = lambda t: _log("결과", t)
    stub.markdown = lambda *a, **k: None
    stub.stop = lambda: (_ for _ in ()).throw(Stop())
    stub.radio = lambda label, options, **k: _val(label, options[0])
    stub.selectbox = lambda label, options, index=0, **k: _val(
        label, list(options)[index])
    stub.date_input = lambda label, value=None, **k: _val(label, value)
    stub.number_input = lambda label, mn=None, mx=None, value=None, **k: _val(
        label, value)
    stub.checkbox = lambda label, value=False, **k: _val(label, value)
    stub.slider = lambda label, mn=None, mx=None, value=None, **k: _val(
        label, value)
    stub.button = lambda label, **k: "조회 시작" in str(label)
    stub.session_state = {}
    stub.cache_resource = lambda fn: fn
    stub.code = lambda *a, **k: None
    stub.text_area = lambda *a, **k: None

    @contextlib.contextmanager
    def expander(label, expanded=False):
        yield Ctx()
    stub.expander = expander
    stub.dataframe = lambda *a, **k: None
    stub.link_button = lambda *a, **k: None
    stub.columns = lambda n: [Ctx() for _ in range(
        n if isinstance(n, int) else len(n))]
    stub.empty = lambda: Ctx()
    stub.progress = lambda v: Ctx()
    stub.column_config = AnyAttr()

    @contextlib.contextmanager
    def spinner(t):
        print(f"  ... {t}", flush=True)
        yield
    stub.spinner = spinner

    def download_button(label, data=None, file_name=None, **k):
        raw = data.getvalue() if hasattr(data, "getvalue") else (data or b"")
        path = os.path.join(OUT_DIR, file_name or "result.xlsx")
        with open(path, "wb") as f:
            f.write(raw)
        saved["files"].append(path)
        _log("저장", f"{path} ({len(raw):,} bytes)")
    stub.download_button = download_button
    return stub, saved


# ------------------------------------------------------------
# 2) 검색 조건 (환경변수 → app.py 위젯 값)
# ------------------------------------------------------------
def build_widgets():
    today = datetime.date.today()          # 워크플로에서 TZ=Asia/Seoul 지정
    days_back = int(os.environ.get("DAYS_BACK", "1"))
    return {
        "검색유형": os.environ.get("SEARCH_TYPE", "개찰결과"),
        "조회 시작일": today - datetime.timedelta(days=days_back),
        "조회 종료일": today,
        "업무구분": os.environ.get("TASK", "공사"),
        "참가제한지역": os.environ.get("REGION", "경상남도"),
        "전국(제한없음) 공고도": os.environ.get(
            "INCLUDE_NATIONWIDE", "0") == "1",
        "추정가격 하한": float(os.environ.get("PRICE_MIN_UK", "1.0")),
        "추정가격 상한": float(os.environ.get("PRICE_MAX_UK", "0")),
        "공고서에서 실무담당": True,
        "Claude AI 정밀 추출": bool(os.environ.get("ANTHROPIC_API_KEY")),
        "첨부파일 분석 개수": int(os.environ.get("MAX_ATTACH", "4")),
    }


# ------------------------------------------------------------
# 3) 메일 발송
# ------------------------------------------------------------
def send_mail(files, summary):
    host = os.environ.get("SMTP_HOST")
    to = os.environ.get("MAIL_TO")
    if not host or not to:
        print("메일 설정(SMTP_HOST/MAIL_TO) 없음 → 파일 저장만 하고 종료")
        return False
    port = int(os.environ.get("SMTP_PORT", "465"))
    user = os.environ["SMTP_USER"]
    pw = os.environ["SMTP_PASS"]
    sender = os.environ.get("MAIL_FROM", user)

    today = datetime.date.today().strftime("%Y-%m-%d")
    msg = EmailMessage()
    msg["Subject"] = f"[나라장터] {today} 개찰결과 리포트"
    msg["From"] = sender
    msg["To"] = to
    body = [f"{today} 나라장터 자동 조회 결과입니다.", ""]
    body += summary
    body += ["", "첨부된 엑셀에서 실무담당 연락처를 확인하세요.",
             "- 자동 발송 (g2b-grab daily report)"]
    msg.set_content("\n".join(body))
    for path in files:
        with open(path, "rb") as f:
            msg.add_attachment(
                f.read(), maintype="application",
                subtype=("vnd.openxmlformats-officedocument"
                         ".spreadsheetml.sheet"),
                filename=os.path.basename(path))

    if port == 465:
        with smtplib.SMTP_SSL(host, port,
                              context=ssl.create_default_context()) as s:
            s.login(user, pw)
            s.send_message(msg)
    else:                                   # 587 STARTTLS
        with smtplib.SMTP(host, port) as s:
            s.starttls(context=ssl.create_default_context())
            s.login(user, pw)
            s.send_message(msg)
    print(f"메일 발송 완료 → {to} (첨부 {len(files)}개)")
    return True


# ------------------------------------------------------------
# 4) 실행
# ------------------------------------------------------------
def main():
    if not os.environ.get("SERVICE_KEY"):
        print("오류: SERVICE_KEY 환경변수가 없습니다.")
        sys.exit(1)
    secrets = {"SERVICE_KEY": os.environ["SERVICE_KEY"]}
    if os.environ.get("ANTHROPIC_API_KEY"):
        secrets["ANTHROPIC_API_KEY"] = os.environ["ANTHROPIC_API_KEY"]

    stub, saved = build_stub(build_widgets(), secrets)
    sys.modules["streamlit"] = stub

    app_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            "app.py")
    src = open(app_path, encoding="utf-8").read()
    g = {"__name__": "__main__"}
    try:
        exec(compile(src, "app.py", "exec"), g)
    except stub.StopException:
        pass

    summary = [l for l in saved["logs"]
               if l.startswith(("[결과]", "[안내]", "[경고]", "[오류]"))]
    if not saved["files"]:
        print("생성된 엑셀이 없습니다(조건 충족 0건 또는 조회 실패).")
        # 0건이어도 알림 메일은 발송해 매일 확인 가능하게 한다
        send_mail([], summary or ["조건 충족 0건"])
        return
    send_mail(saved["files"], summary)


if __name__ == "__main__":
    main()
