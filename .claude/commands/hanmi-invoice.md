한미택배 송장 엑셀을 네이버 스마트스토어 출고 파일에 매칭하여 송장번호를 입력하고 발송처리용 파일을 생성한다.

## 사용법
인자: `$ARGUMENTS`
- 인자가 없으면 사용자에게 스마트스토어 파일 경로/비밀번호, 한미 송장 파일 경로를 질문한다.
- 인자 형식: `스마트스토어파일경로 비밀번호 한미송장파일경로`
  - 예: `~/Downloads/스마트스토어.xlsx 1111 ~/Downloads/송장리스트-Apr21.xls`

## 실행 절차

1. **필수 라이브러리 확인 및 설치**

```bash
pip3 install msoffcrypto-tool pandas openpyxl xlrd -q
```

2. **매칭 및 파일 생성** 아래 Python 코드를 실행한다:

```python
import msoffcrypto, io, os, pandas as pd
from openpyxl import load_workbook
from datetime import datetime

SMARTSTORE_FILE = "<스마트스토어_파일_경로>"
PASSWORD = "<비밀번호>"
HANMI_INVOICE_FILE = "<한미송장_파일_경로>"

def normalize_phone(v):
    if pd.isna(v): return ''
    return str(v).replace('-', '').replace(' ', '').strip()

# 한미 송장 파일 읽기 — 이름+전화번호1로 매핑
hanmi = pd.read_excel(os.path.expanduser(HANMI_INVOICE_FILE), header=0)
name_phone_to_tracking = {}
for _, r in hanmi.iterrows():
    key = (r['받는사람'].strip(), normalize_phone(r['전화번호1']))
    name_phone_to_tracking[key] = str(r['Tracking No'])

# 스마트스토어 파일 복호화
with open(os.path.expanduser(SMARTSTORE_FILE), 'rb') as f:
    office_file = msoffcrypto.OfficeFile(f)
    office_file.load_key(password=PASSWORD)
    decrypted = io.BytesIO()
    office_file.decrypt(decrypted)
    decrypted.seek(0)
    wb = load_workbook(decrypted)

ws = wb.active

# 2행에서 컬럼 인덱스 찾기
col_index = {}
for col in range(1, ws.max_column + 1):
    val = ws.cell(2, col).value
    if val:
        col_index[val] = col

song_jang_col = col_index.get('송장번호')
name_col = col_index.get('수취인명')
phone_col = col_index.get('수취인연락처1')

# 3행부터 매칭하여 송장번호 입력 (택배사는 원본 유지)
filled = 0
unmatched = []
for row in range(3, ws.max_row + 1):
    name = ws.cell(row, name_col).value
    if not name:
        continue
    phone = normalize_phone(ws.cell(row, phone_col).value)
    key = (str(name).strip(), phone)
    tracking = name_phone_to_tracking.get(key)
    if tracking:
        ws.cell(row, song_jang_col).value = tracking
        filled += 1
    else:
        unmatched.append((str(name).strip(), phone))

# 1행(안내문) 삭제, 시트명 변경
ws.delete_rows(1)
ws.title = '발송처리'

today = datetime.now().strftime('%Y%m%d')
output_path = os.path.expanduser(f"~/smartstore-project/output/스마트스토어_발송처리_{today}.xlsx")
wb.save(output_path)

print(f"DONE:{output_path}:{filled}")
if unmatched:
    print("UNMATCHED:" + "|".join([f"{n}({p})" for n, p in unmatched]))
```

3. **결과 보고**
   - 저장 경로와 매칭된 건수를 사용자에게 알려준다.
   - `UNMATCHED:` 로 시작하는 줄이 있으면 매칭되지 않은 수취인 목록을 사용자에게 보여준다.
     - 미매칭은 오늘 스마트스토어 파일에 없는 다른 날 주문일 수 있으므로 참고용으로만 안내한다.

## 매칭 기준
- **수취인명** (스마트스토어 `수취인명` ↔ 한미 `받는사람`)
- **전화번호** (스마트스토어 `수취인연락처1` ↔ 한미 `전화번호1`) — 하이픈 제거 후 비교

## 고정 처리
- 1행(안내문) 삭제
- 시트명 `발주발송관리` → `발송처리`
- 택배사는 원본 파일 값 그대로 유지 (우체국택배 등)
- 결과 파일 저장: `~/smartstore-project/output/스마트스토어_발송처리_YYYYMMDD.xlsx`
