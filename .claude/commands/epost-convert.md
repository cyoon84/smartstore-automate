네이버 스마트스토어 출고 준비중 엑셀 파일을 우체국택배 업로드용 엑셀 파일로 변환한다.

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
import msoffcrypto, io, pandas as pd, os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

SMARTSTORE_FILE = "<스마트스토어_파일_경로>"
PASSWORD = "<비밀번호>"

# 스마트스토어 파일 복호화
with open(os.path.expanduser(SMARTSTORE_FILE), 'rb') as f:
    office_file = msoffcrypto.OfficeFile(f)
    office_file.load_key(password=PASSWORD)
    decrypted = io.BytesIO()
    office_file.decrypt(decrypted)
    decrypted.seek(0)
    ss = pd.read_excel(decrypted, engine='openpyxl', header=1)

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

# 주문번호 기준 중복 제거 (한 주문 = 한 행)
ss['_order_num'] = ss['주문번호'].astype(str)
ss_dedup = ss.drop_duplicates(subset='_order_num', keep='first').reset_index(drop=True)

# 우체국택배 컬럼 헤더
headers = [
    '받는 분', '우편번호', '주소(시도+시군구+도로명+건물번호)',
    '상세주소(동, 호수, 洞명칭, 아파트, 건물명 등)',
    '일반전화(02-1234-5678)', '휴대전화(010-1234-5678)',
    '중량(kg)', '부피(cm)=가로+세로+높이', '내용품코드', '내용물',
    '배달방식', '배송시요청사항', '분할접수 여부(Y/N)',
    '분할접수 첫번째 중량(kg)', '분할접수 첫번째 부피(cm)',
    '분할접수 두번째 중량(kg)', '분할접수 두번째 부피(cm)'
]

rows = []
for _, row in ss_dedup.iterrows():
    rows.append([
        val(row, '수취인명'),        # A 받는 분
        zipcode(row, '우편번호'),    # B 우편번호
        val(row, '기본배송지'),      # C 주소
        val(row, '상세배송지'),      # D 상세주소
        phone(row, '수취인연락처2'), # E 일반전화
        phone(row, '수취인연락처1'), # F 휴대전화
        3,                           # G 중량(kg) — 고정값
        80,                          # H 부피 — 고정값
        '생활용품',                  # I 내용품코드 — 고정값
        '',                          # J 내용물
        '',                          # K 배달방식
        val(row, '배송메세지'),      # L 배송시요청사항
        'N',                         # M 분할접수 여부
        '', '', '', ''               # N~Q 분할접수 관련
    ])

# 엑셀 생성
wb = Workbook()
ws = wb.active

header_fill = PatternFill('solid', start_color='C00000', end_color='C00000')
header_font = Font(name='Arial', bold=True, color='FFFFFF', size=9)
data_font   = Font(name='Arial', size=9)
center = Alignment(horizontal='center', vertical='center', wrap_text=True)
left   = Alignment(horizontal='left',   vertical='center', wrap_text=False)
thin   = Side(style='thin', color='CCCCCC')
border = Border(left=thin, right=thin, top=thin, bottom=thin)

col_widths = [12, 8, 40, 22, 14, 14, 8, 8, 16, 10, 8, 24, 14, 14, 14, 14, 14]

for ci, (h, w) in enumerate(zip(headers, col_widths), 1):
    cell = ws.cell(row=1, column=ci, value=h)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center
    cell.border = border
    ws.column_dimensions[get_column_letter(ci)].width = w
ws.row_dimensions[1].height = 30

for ri, row_data in enumerate(rows, 2):
    for ci, v in enumerate(row_data, 1):
        cell = ws.cell(row=ri, column=ci, value=v)
        cell.font = data_font
        cell.border = border
        cell.alignment = center if ci in [1, 2, 7, 8, 9, 13] else left
        if ci == 2:
            cell.number_format = '@'
    ws.row_dimensions[ri].height = 16

today = datetime.now().strftime('%Y%m%d')
output_path = os.path.expanduser(f'~/smartstore-project/output/우체국택배_업로드용_{today}.xlsx')
wb.save(output_path)
print(f'✅ 변환 완료: {output_path} ({len(rows)}건)')
```

3. **완료 후 결과 보고**
   - 저장 경로와 변환된 주문 건수를 사용자에게 알려준다.
   - 오류 발생 시 원인을 설명하고 해결 방법을 안내한다.

## 고정값

| 컬럼 | 값 |
|------|-----|
| 중량 (G) | 3kg |
| 부피 (H) | 80 |
| 내용품코드 (I) | 생활용품 |
| 분할접수 여부 (M) | N |

## 참고
- 주문번호 기준으로 중복 제거 — 한 주문에 여러 상품이 있어도 한 행만 생성
- 우편번호는 텍스트 형식으로 저장 (앞자리 0 보존)
- 비밀번호가 없는 경우 PASSWORD를 빈 문자열로 처리
- 결과 파일 저장 경로: `~/smartstore-project/output/`
