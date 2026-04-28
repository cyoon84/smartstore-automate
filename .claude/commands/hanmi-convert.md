네이버 스마트스토어 출고 준비중 엑셀 파일을 한미택배 업로드용 엑셀 파일로 변환한다.

## 사용법
인자: `$ARGUMENTS`
- 인자가 없으면 사용자에게 스마트스토어 파일 경로와 비밀번호를 질문한다.
- 인자 형식: `파일경로 비밀번호` (예: `~/Downloads/스마트스토어.xlsx 1111`)

## 실행 절차

1. **필수 라이브러리 확인 및 설치**

```bash
pip3 install msoffcrypto-tool pandas openpyxl xlrd -q
```

2. **파일 읽기 및 변환** 아래 Python 코드를 실행한다:

```python
import pandas as pd
import msoffcrypto, io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import sys, os
from datetime import datetime

SMARTSTORE_FILE = "<스마트스토어_파일_경로>"
PASSWORD = "<비밀번호>"
HANMI_TEMPLATE = os.path.expanduser("~/smartstore-project/templates/hanmi-form.xls")
PRODUCT_MAPPING = os.path.expanduser("~/smartstore-project/templates/product-mapping.xlsx")
CONFIG_FILE = os.path.expanduser("~/smartstore-project/config.json")

import json
with open(CONFIG_FILE) as f:
    cfg = json.load(f)
BUSINESS_ID    = cfg["business_id"]
SENDER_NAME    = cfg["sender_name"]
SENDER_EMAIL   = cfg["sender_email"]
SENDER_PHONE   = cfg["sender_phone"]
SENDER_ADDRESS = cfg["sender_address"]

# 스마트스토어 파일 복호화
with open(os.path.expanduser(SMARTSTORE_FILE), 'rb') as f:
    office_file = msoffcrypto.OfficeFile(f)
    office_file.load_key(password=PASSWORD)
    decrypted = io.BytesIO()
    office_file.decrypt(decrypted)
    decrypted.seek(0)
    ss = pd.read_excel(decrypted, engine='openpyxl', header=1)

# 한미택배 컬럼 읽기
hanmi_df = pd.read_excel(HANMI_TEMPLATE, header=0)
hanmi_cols = [col.split('.')[0] for col in hanmi_df.columns]

# 상품 매핑 테이블 읽기 (상품번호 → 영문명, HS Code, Site URL)
mapping_df = pd.read_excel(PRODUCT_MAPPING, dtype={'상품번호': str})
mapping_df['상품번호'] = mapping_df['상품번호'].str.split('.').str[0]
mapping_df = mapping_df.drop_duplicates(subset='상품번호', keep='first')
product_map = mapping_df.set_index('상품번호').to_dict('index')

def val(row, col):
    v = row.get(col)
    return '' if pd.isna(v) else v

def phone(row, col):
    v = val(row, col)
    return str(v).replace('-', '') if v != '' else ''

def zipcode(row, col):
    v = val(row, col)
    if v == '': return ''
    s = str(v).replace('.0', '').strip()
    return s.zfill(5) if s.isdigit() else s

def mget(pmap, col):
    if pmap is None: return ''
    v = pmap.get(col, '')
    return '' if pd.isna(v) else str(v) if v != '' else ''

# 같은 주문 + 같은 상품번호 → 수량 합산
ss['_product_num'] = ss['상품번호'].astype(str).str.split('.').str[0]
ss['_order_num']   = ss['주문번호'].astype(str)
agg_dict = {col: 'first' for col in ss.columns if col not in ['수량', '_product_num', '_order_num']}
agg_dict['수량'] = 'sum'
ss = ss.groupby(['_order_num', '_product_num'], sort=False).agg(agg_dict).reset_index()

# HS CODE 컬럼 인덱스 (0-based)
hs_col_idx = hanmi_cols.index('HS CODE')

missing_products = []
rows = []
seen_orders = {}  # 주문번호 → 첫 번째 행 번호 추적
row_num = 1

# 컬럼 인덱스 사전 계산 (중복 컬럼명 문제 방지 — dict 대신 list 사용)
ZIPCODE_COL_IDX = next(i for i, c in enumerate(hanmi_cols) if c == '우편번호')

for i, row in ss.iterrows():
    product_num = row['_product_num']
    order_num   = row['_order_num']
    pmap = product_map.get(product_num)

    eng_name   = mget(pmap, '상품명(영문)')
    hs_code    = mget(pmap, 'HS CODE')
    brand      = mget(pmap, '브랜드')
    unit_price = mget(pmap, '단가')
    site_url   = mget(pmap, 'SITE URL')
    seller     = mget(pmap, '해외판매자 상호')

    if not eng_name:
        missing_products.append((product_num, val(row, '상품명')))

    is_first = order_num not in seen_orders
    if is_first:
        seen_orders[order_num] = row_num

    # list 기반으로 구성 (중복 컬럼명 충돌 방지)
    new_row = [''] * len(hanmi_cols)

    # HS CODE 이후는 항상 채움
    new_row[hs_col_idx]      = hs_code
    new_row[hs_col_idx + 1]  = ''
    new_row[hs_col_idx + 2]  = eng_name if eng_name else val(row, '상품명')
    new_row[hs_col_idx + 3]  = brand
    new_row[hs_col_idx + 4]  = unit_price if unit_price else val(row, '상품가격')
    new_row[hs_col_idx + 5]  = val(row, '수량')
    new_row[hs_col_idx + 6]  = site_url
    new_row[hs_col_idx + 7]  = ''
    new_row[hs_col_idx + 8]  = 'B'
    new_row[hs_col_idx + 9]  = ''
    new_row[hs_col_idx + 10] = seller
    new_row[hs_col_idx + 11] = ''
    new_row[hs_col_idx + 12] = SENDER_NAME
    new_row[hs_col_idx + 13] = ''
    new_row[hs_col_idx + 14] = ''
    new_row[hs_col_idx + 15] = val(row, '주문번호')

    # 첫 번째 아이템만 앞쪽 전체 채움
    if is_first:
        new_row[0]  = row_num
        new_row[1]  = BUSINESS_ID
        new_row[2]  = SENDER_NAME
        new_row[3]  = SENDER_EMAIL
        new_row[4]  = SENDER_PHONE
        new_row[5]  = SENDER_ADDRESS
        new_row[6]  = 1
        new_row[7]  = val(row, '수취인명')
        new_row[8]  = phone(row, '수취인연락처1')
        new_row[9]  = phone(row, '수취인연락처2')
        new_row[10] = zipcode(row, '우편번호')
        new_row[11] = val(row, '기본배송지')
        new_row[12] = val(row, '상세배송지')
        new_row[13] = val(row, '개인통관고유부호')
        new_row[14] = val(row, '배송메세지')
        new_row[15] = 1
        new_row[16] = 'a'
        new_row[17] = 1
        new_row[18] = 1
        new_row[19] = 1
        new_row[20] = 1
        new_row[21] = 1
        new_row[22] = 1
        row_num += 1

    rows.append(new_row)

missing_products_output = []
if missing_products:
    for num, name in missing_products:
        missing_products_output.append(f"  상품번호 {num}: {name}")
    print("MISSING:" + "\n".join(missing_products_output))

wb = Workbook()
ws = wb.active
header_fill = PatternFill('solid', start_color='366092', end_color='366092')
header_font = Font(name='Arial', bold=True, color='FFFFFF', size=9)
data_font   = Font(name='Arial', size=9)
center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
left_align   = Alignment(horizontal='left',   vertical='center')
thin = Side(style='thin', color='AAAAAA')
border = Border(left=thin, right=thin, top=thin, bottom=thin)

for col_idx, col_name in enumerate(hanmi_cols, 1):
    cell = ws.cell(row=1, column=col_idx, value=col_name)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center_align
    cell.border = border

ZIPCODE_COL = ZIPCODE_COL_IDX + 1  # 1-based

for row_idx, row_data in enumerate(rows, 2):
    for col_idx, v in enumerate(row_data, 1):
        cell = ws.cell(row=row_idx, column=col_idx, value=v)
        cell.font = data_font
        cell.border = border
        cell.alignment = center_align if col_idx in [1,6,15,16,17,18,19,20,21,22,28,31] else left_align
        if col_idx == ZIPCODE_COL:
            cell.number_format = '@'

col_widths = {1:6,2:14,3:14,4:20,5:13,6:50,7:10,8:16,9:14,10:14,
              11:10,12:40,13:20,14:16,15:20,16:12,17:50,18:10,19:10,
              20:10,21:8,22:10,23:8,24:12,25:10,26:45,27:12,28:10,29:8,
              30:15,31:15,32:45,33:15,34:15,35:15,36:15,37:15,38:15,39:40}
for col_idx, width in col_widths.items():
    ws.column_dimensions[get_column_letter(col_idx)].width = width
ws.row_dimensions[1].height = 40

today = datetime.now().strftime('%Y%m%d')
output_path = os.path.expanduser(f"~/smartstore-project/output/한미택배_업로드용_{today}.xlsx")
wb.save(output_path)
print(f"✅ 변환 완료: {output_path} ({len(rows)}건)")
```

3. **매핑 누락 상품 처리 (있을 경우)**

   Python 코드 실행 결과에서 `missing_products` 리스트가 비어있지 않으면:

   **3-1. 자동 유추**

   누락된 상품마다 한글 상품명을 분석하여 아래 항목을 **직접 유추**한다:
   - **영문 상품명**: 한글 상품명에서 브랜드명·제품명·용량/수량 정보를 조합해 자연스러운 영문명 생성
   - **HS CODE**: `~/smartstore-project/templates/hs-code-reference.xlsx` 를 참고해 가장 적합한 코드 선택
   - **브랜드**: 상품명에서 추출 (예: "팀홀튼 믹스커피" → "Tim Hortons")
   - **해외판매자 상호**: 사이트 URL의 도메인에서 추출 (대문자로, 예: `www.costco.ca` → `COSTCO`, `www.walmart.ca` → `WALMART`). URL을 모르는 경우 브랜드명으로 대체

   **3-2. 사용자 확인 요청**

   유추한 내용을 아래 형식으로 한 번에 보여주고 확인을 요청한다:

   ```
   [새 상품 확인] 상품번호 XXXXXXX
   한글명: (원본 상품명)
   영문명: (유추한 영문명)  ← 맞나요?
   HS Code: (유추한 코드)  ← 맞나요?
   브랜드: (유추한 브랜드)  ← 맞나요?

   확인되면 아래 두 가지만 알려주세요:
   - 단가 (CAD 기준 숫자)
   - 구매 사이트 URL (예: www.costco.ca)
   ```

   - 사용자가 영문명/HS Code/브랜드를 수정하면 그 값을 사용한다.
   - 사용자가 확인만 하면 유추한 값 그대로 사용한다.
   - **SITE URL** 은 입력받은 URL을 아래 형식으로 자동 조합:
     `https://smartstore.naver.com/finchmart_ca/ (입력받은_URL)`
     예: `www.costco.ca` 입력 시 → `https://smartstore.naver.com/finchmart_ca/ (www.costco.ca)`

   - 입력받은 내용을 아래 Python으로 `product-mapping.xlsx`에 추가한다:

   ```python
   from openpyxl import load_workbook

   PRODUCT_MAPPING = os.path.expanduser('~/smartstore-project/templates/product-mapping.xlsx')
   wb_map = load_workbook(PRODUCT_MAPPING)
   ws_map = wb_map.active

   def format_site_url(url):
       if not url: return ''
       return f"https://smartstore.naver.com/finchmart_ca/ ({url})"

   # 누락 상품마다 반복 (site_url은 format_site_url() 적용)
   new_entries = [
       # (상품번호, 한글상품명, hs_code, 영문상품명, 브랜드, 단가, site_url, 해외판매자상호)
       # 사용자 입력값으로 채울 것
   ]
   for entry in new_entries:
       ws_map.append(list(entry))

   wb_map.save(PRODUCT_MAPPING)
   print(f"✅ product-mapping.xlsx에 {len(new_entries)}개 상품 추가됨")
   ```

   - 저장 후 **2번 Python 코드 전체를 다시 실행**하여 최종 한미택배 파일을 생성한다.

4. **완료 후 결과 보고**
   - 저장 경로와 변환된 주문 건수를 사용자에게 알려준다.
   - 오류 발생 시 원인을 설명하고 해결 방법을 안내한다.

## 고정값 (발송인 정보)
발송인 정보는 `~/smartstore-project/config.json` 에서 읽는다. 해당 파일은 gitignore 처리되어 있으며, `config-example.json`을 복사하여 작성한다.

| 항목 | config.json 키 |
|---|---|
| 비즈니스회원 아이디 | business_id |
| 보내는 사람 | sender_name |
| 이메일 | sender_email |
| 전화 | sender_phone |
| 주소 | sender_address |

## 참고
- 한미택배 양식 파일 경로: `~/smartstore-project/templates/hanmi-form.xls`
- 상품 매핑 테이블 경로: `~/smartstore-project/templates/product-mapping.xlsx`
- 결과 파일 저장 경로: `~/smartstore-project/output/`
- 한미택배 양식 파일이 없거나 경로가 다르면 사용자에게 경로를 질문한다.
- 스마트스토어 파일에 비밀번호가 없는 경우 PASSWORD를 빈 문자열로 처리한다.
- 영문 상품명이 없는 상품이 있으면 변환 후 경고를 출력하고, product-mapping.xlsx에 추가하도록 안내한다.
