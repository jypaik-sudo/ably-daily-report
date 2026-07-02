"""
4910 SSA 데일리 리포트 자동화 (zipfile 직접 수정 방식 — 슬라이서 보존)
"""
import io, json, time, zipfile, re, urllib.request, os
from datetime import datetime, timedelta
import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

AB_TOKEN     = os.environ["AIRBRIDGE_TOKEN"]
DEST_FOLDER  = os.environ["DRIVE_FOLDER_ID"]          # 업로드 대상 폴더
SRC_FOLDER   = "1Vo4WHoVllXQA8CjYTPKVRb3g1tI3Td_O"   # 템플릿 소스 폴더
GCP_KEY      = json.loads(os.environ["GCP_SERVICE_ACCOUNT_JSON"])

TARGET = (datetime.utcnow() + timedelta(hours=9) - timedelta(days=1)).strftime("%Y-%m-%d")
print(f"적재 대상: {TARGET}")

# ── Google Drive ──────────────────────────────────────────────────
creds = service_account.Credentials.from_service_account_info(
    GCP_KEY, scopes=["https://www.googleapis.com/auth/drive"])
drive = build("drive", "v3", credentials=creds)

def list_files(folder_id):
    return drive.files().list(
        q=f"'{folder_id}' in parents and trashed=false",
        fields="files(id,name)", orderBy="modifiedTime desc",
        supportsAllDrives=True, includeItemsFromAllDrives=True
    ).execute().get("files", [])

# 중복 확인: 목적지 폴더에 이미 오늘 날짜 파일이 있으면 스킵
month_day = f"{int(TARGET[5:7])}.{int(TARGET[8:10])}"
dest_files = list_files(DEST_FOLDER)
if any(month_day in f["name"] for f in dest_files):
    print(f"이미 {month_day} 파일 존재. 스킵.")
    exit(0)

# 소스 폴더에서 가장 최근 파일 다운로드
src_files = list_files(SRC_FOLDER)
if not src_files:
    print("소스 폴더에 파일 없음. 종료."); exit(1)
src = src_files[0]
print(f"소스: {src['name']}")

buf = io.BytesIO()
dl = MediaIoBaseDownload(buf, drive.files().get_media(fileId=src["id"], supportsAllDrives=True))
done = False
while not done: _, done = dl.next_chunk()
xlsx_bytes = buf.getvalue()
print(f"다운로드: {len(xlsx_bytes):,} bytes")

# ── Airbridge ─────────────────────────────────────────────────────
HEADERS = {"Authorization": f"Bearer {AB_TOKEN}", "Content-Type": "application/json"}
BODY = {
    "groupBys": ["event_date","campaign","ad_group","ad_creative"],
    "metrics": ["cost_channel","impressions_channel","clicks_channel","web_opens",
                "app_web_sign_up","app_installs","app_first_installs",
                "app_custom_users_COMPLETE_ORDER_DOMESTIC","web_custom_users_COMPLETE_ORDER_DOMESTIC",
                "app_web_order_complete","app_web_revenue","web_order_complete_users","app_order_complete_users"],
    "filters": [
        {"field":"channel","filterType":"IN","values":["naver.searchad"]},
        {"field":"campaign","filterType":"IN","values":["구매 캠페인"]},
        {"field":"ad_group","filterType":"LIKE","values":["*네이버_검색광고"]}
    ],
    "sorts": [{"fieldName":"event_date","isAscending":False}],
    "option": {"timezone": None, "eventTimestampSource": "event_occurred_date"},
    "size": 500000, "isSummaryAvailable": True, "viewFormat": True, "skip": 0,
    "from": TARGET, "to": TARGET
}

with urllib.request.urlopen(urllib.request.Request(
    "https://api.airbridge.io/reports/api/v7/apps/4910/actuals/query",
    data=json.dumps(BODY).encode(), headers=HEADERS, method="POST"), timeout=15) as r:
    task_id = json.loads(r.read())["task"]["taskId"]
print(f"taskId: {task_id}")

for _ in range(40):
    time.sleep(3)
    with urllib.request.urlopen(urllib.request.Request(
        f"https://api.airbridge.io/reports/api/v7/apps/4910/actuals/query/{task_id}",
        headers=HEADERS), timeout=10) as r:
        status = json.loads(r.read())["task"]["status"]
    if status == "SUCCESS":
        print("SUCCESS"); break
    print(f"  대기 중... ({status})")

with urllib.request.urlopen(urllib.request.Request(
    f"https://api.airbridge.io/reports/api/v7/apps/4910/actuals/export/{task_id}"
    f"?skip=0&size=500000&isSummaryAvailable=true&viewFormat=true",
    headers=HEADERS), timeout=10) as r:
    s3_url = json.loads(r.read())["url"]

csv_bytes = urllib.request.urlopen(s3_url, timeout=60).read()
df = pd.read_csv(io.BytesIO(csv_bytes), encoding="utf-8-sig")
df = df[df["Event Date"] == TARGET].reset_index(drop=True)
n = len(df)
print(f"데이터: {n}행")
if n == 0:
    print("데이터 없음. 종료."); exit(0)

# ── ZIP 직접 수정 ─────────────────────────────────────────────────
with zipfile.ZipFile(io.BytesIO(xlsx_bytes)) as z:
    wb_xml  = z.read("xl/workbook.xml").decode("utf-8")
    wb_rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    sheets  = re.findall(r'name="([^"]+)"[^>]+r:id="([^"]+)"', wb_xml)
    rels    = dict(re.findall(r'Id="([^"]+)"[^>]+Target="([^"]+)"', wb_rels))

    rd_rid   = next(rid for name, rid in sheets if name == "RD")
    rd_file  = f"xl/{rels[rd_rid]}"

    rd_rels_file = rd_file.replace("/worksheets/sheet", "/worksheets/_rels/sheet").replace(".xml", ".xml.rels")
    rd_rels_xml  = z.read(rd_rels_file).decode("utf-8")
    tbl_match    = re.search(r'Type="[^"]*table[^"]*"[^>]+Target="([^"]+)"', rd_rels_xml)
    tbl_rel      = tbl_match.group(1) if tbl_match else None
    tbl_file     = tbl_rel.replace("../", "xl/") if tbl_rel else None
    print(f"RD: {rd_file}, Table: {tbl_file}")

    # sharedStrings
    ss_xml     = z.read("xl/sharedStrings.xml").decode("utf-8")
    ss_strings = []
    for si in re.findall(r"<si>(.*?)</si>", ss_xml, re.DOTALL):
        t = re.search(r"<t[^>]*>([^<]*)</t>", si)
        ss_strings.append(t.group(1) if t else "")
    ss_map      = {s: i for i, s in enumerate(ss_strings)}
    new_strings = list(ss_strings)

    def get_ss_idx(text):
        if text not in ss_map:
            ss_map[text] = len(new_strings)
            new_strings.append(text)
        return ss_map[text]

    rd_xml   = z.read(rd_file).decode("utf-8")
    last_row = int(re.findall(r'<row r="(\d+)"', rd_xml)[-1])
    print(f"마지막 행: {last_row}")

    # 수식 템플릿 (row 2에서 추출)
    sample_m = re.search(r'<row r="2"[^>]*>(.*?)</row>', rd_xml, re.DOTALL)
    if not sample_m:
        sample_m = re.search(r'<row r="\d+"[^>]*>(.*?)</row>', rd_xml, re.DOTALL)
    sample_cells = sample_m.group(1) if sample_m else ""

    def get_cell_template(cells_str, col):
        m = re.search(rf'<c r="{col}\d+"((?:[^>])*?)>(.*?)</c>', cells_str, re.DOTALL)
        return (m.group(1), m.group(2)) if m else (None, None)

    formula_cols = {col: get_cell_template(sample_cells, col) for col in ["A","B","C","D","E","F","Z"]}

    # I열 날짜 스타일
    i_m = re.search(r'<c r="I\d+"([^>]*)>', rd_xml)
    i_s = re.search(r's="(\d+)"', i_m.group(1)) if i_m else None
    date_style  = i_s.group(1) if i_s else "3"
    excel_date  = (datetime.strptime(TARGET, "%Y-%m-%d") - datetime(1899, 12, 30)).days

    CSV_TEXT = [("J","Campaign"), ("K","Ad Group"), ("L","Ad Creative")]
    CSV_NUM  = [
        ("M","Cost (Channel)"), ("N","Impressions (Channel)"), ("O","Clicks (Channel)"),
        ("P","Opens (Web)"), ("Q","Sign-up (App+Web)"), ("R","Installs (App)"),
        ("S","First Installs (App)"), ("T","COMPLETE_ORDER_DOMESTIC Users (App)"),
        ("U","COMPLETE_ORDER_DOMESTIC Users (Web)"), ("V","Order Complete (App+Web)"),
        ("W","Revenue (App+Web)"), ("X","Order Complete Users (Web)"), ("Y","Order Complete Users (App)")
    ]

    def xml_esc(s):
        return s.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

    rows_parts = []
    for i, row_data in df.iterrows():
        r = last_row + 1 + i
        cells = []

        for col in ["A","B","C","D","E","F"]:
            attrs, content = formula_cols[col]
            if attrs is not None:
                cells.append(f'<c r="{col}{r}"{attrs}>{content}</c>')

        cells.append(f'<c r="G{r}" s="8" t="str"><f>IF(COUNTIF(&apos;운영 중인 그룹&apos;!$B$2:$B$196,K{r})&gt;0,&quot;Y&quot;,&quot;N&quot;)</f></c>')
        cells.append(f'<c r="H{r}" s="8" t="str"><f>IFERROR(MID(L{r},FIND(&quot;[&quot;,L{r})+1,FIND(&quot;]&quot;,L{r})-FIND(&quot;[&quot;,L{r})-1),&quot;&quot;)</f></c>')
        cells.append(f'<c r="I{r}" s="{date_style}"><v>{excel_date}</v></c>')

        for col, csv_col in CSV_TEXT:
            val = str(row_data.get(csv_col, ""))
            val = "" if val == "nan" else val
            idx = get_ss_idx(val)
            cells.append(f'<c r="{col}{r}" t="s"><v>{idx}</v></c>')

        for col, csv_col in CSV_NUM:
            val = row_data.get(csv_col, 0)
            v   = 0.0 if str(val) == "nan" else float(val)
            cells.append(f'<c r="{col}{r}"><v>{v}</v></c>')

        attrs, content = formula_cols["Z"]
        if attrs is not None:
            cells.append(f'<c r="Z{r}"{attrs}>{content}</c>')

        rows_parts.append(f'<row r="{r}" spans="1:26" x14ac:dyDescent="0.6">{"".join(cells)}</row>')
        if (i + 1) % 500 == 0:
            print(f"  행 생성 {i+1}/{n}")

    new_last    = last_row + n
    new_rd_xml  = rd_xml.replace("</sheetData>", "".join(rows_parts) + "</sheetData>")
    # RD 시트 커서도 A1으로
    new_rd_xml  = re.sub(r'<selection[^/]*/>', '<selection activeCell="A1" sqref="A1"/>', new_rd_xml)

    tbl_xml     = z.read(tbl_file).decode("utf-8") if tbl_file else ""
    tbl_xml_new = re.sub(r'ref="([A-Z]+1:[A-Z]+)\d+"', rf'ref="\g<1>{new_last}"', tbl_xml) if tbl_xml else ""

    ss_count  = len(new_strings)
    ss_items  = "".join(f'<si><t xml:space="preserve">{xml_esc(s)}</t></si>' for s in new_strings)
    new_ss    = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                 f'count="{ss_count}" uniqueCount="{ss_count}">{ss_items}</sst>')

    # workbook.xml: fullCalcOnLoad + Summary를 activeTab으로 설정
    wb_mod = wb_xml
    if "fullCalcOnLoad" not in wb_mod:
        wb_mod = re.sub(r'<calcPr([^/]*)/>', r'<calcPr\1 fullCalcOnLoad="1"/>', wb_mod)
        if "calcPr" not in wb_mod:
            wb_mod = wb_mod.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')

    sheet_names = re.findall(r'<sheet\b[^>]*\bname="([^"]+)"', wb_mod)
    summary_idx = next((i for i, n in enumerate(sheet_names) if n == "Summary"), 0)
    print(f"Summary 탭 인덱스: {summary_idx} (전체 시트: {sheet_names})")
    wb_mod = re.sub(r'(<workbookView\b[^>]*\b)activeTab="\d+"', rf'\1activeTab="{summary_idx}"', wb_mod)
    if "activeTab" not in wb_mod:
        wb_mod = re.sub(r'<workbookView\b', f'<workbookView activeTab="{summary_idx}"', wb_mod, count=1)

    # 모든 worksheet 파일 목록 (RD 제외 → RD는 이미 new_rd_xml로 처리)
    worksheet_files = {f"xl/{rels[rid]}" for _, rid in sheets}

    def set_a1_cursor(xml_str):
        """모든 <selection> 을 A1으로 초기화"""
        xml_str = re.sub(r'<selection\b[^/]*/>', '<selection activeCell="A1" sqref="A1"/>', xml_str)
        xml_str = re.sub(r'<selection\b[^>]*>.*?</selection>', '<selection activeCell="A1" sqref="A1"/>', xml_str, flags=re.DOTALL)
        return xml_str

    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(xlsx_bytes), "r") as zin, \
         zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            fn = item.filename
            if fn == "xl/calcChain.xml":
                continue
            elif fn == rd_file:
                zout.writestr(item, new_rd_xml.encode("utf-8"))
            elif tbl_file and fn == tbl_file:
                zout.writestr(item, tbl_xml_new.encode("utf-8"))
            elif fn == "xl/sharedStrings.xml":
                zout.writestr(item, new_ss.encode("utf-8"))
            elif fn == "xl/workbook.xml":
                zout.writestr(item, wb_mod.encode("utf-8"))
            elif "pivotCacheDefinition" in fn:
                data = zin.read(fn).decode("utf-8")
                data = re.sub(r'refreshOnLoad="0"', 'refreshOnLoad="1"', data)
                if "refreshOnLoad" not in data:
                    data = data.replace("<pivotCacheDefinition ", '<pivotCacheDefinition refreshOnLoad="1" ', 1)
                zout.writestr(item, data.encode("utf-8"))
            elif fn in worksheet_files and fn != rd_file:
                # RD 외 모든 시트: 커서 A1 설정 (슬라이서 등 나머지는 그대로)
                data = set_a1_cursor(zin.read(fn).decode("utf-8"))
                zout.writestr(item, data.encode("utf-8"))
            else:
                zout.writestr(item, zin.read(fn))

final_bytes = out.getvalue()
print(f"최종 크기: {len(final_bytes):,} bytes")

new_name = f"4910_SSA_데일리리포트_{month_day}.xlsx"
print(f"업로드: {new_name}")

media = MediaIoBaseUpload(io.BytesIO(final_bytes),
    mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", resumable=True)
uploaded = drive.files().create(
    body={"name": new_name, "parents": [DEST_FOLDER]},
    media_body=media, fields="id,name",
    supportsAllDrives=True
).execute()
print(f"완료: {uploaded['name']} ({uploaded['id']})")
