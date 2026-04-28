네이버 스마트스토어 출고 준비중 엑셀 파일을 우체국택배 업로드용 엑셀 파일로 변환한다.

## 사용법
인자: `$ARGUMENTS`
- 인자가 없으면 사용자에게 스마트스토어 파일 경로와 비밀번호를 질문한다.
- 인자 형식: `파일경로 비밀번호` (예: `~/Downloads/스마트스토어.xlsx 1111`)

## 실행 절차

1. **필수 라이브러리 확인 및 설치**

```bash
pip3 install msoffcrypto-tool pandas openpyxl xlrd xlwt xlutils -q
```

2. **파일 읽기 및 변환** 아래 Python 코드를 실행한다:

```python
import msoffcrypto, io, os, pandas as pd
import xlrd, xlwt
from xlutils.copy import copy as xl_copy
from datetime import datetime

SMARTSTORE_FILE = "<스마트스토어_파일_경로>"
PASSWORD = "<비밀번호>"
TEMPLATE = os.path.expanduser("~/smartstore-project/templates/post-form.xls")

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

def zipcode(row, col):
    v = val(row, col)
    if v == '': return ''
    s = str(v).replace('.0', '').strip()
    return s.zfill(5) if s.isdigit() else s

# 템플릿 열기 (formatting 유지)
rb = xlrd.open_workbook(TEMPLATE, formatting_info=True)
wb = xl_copy(rb)
ws = wb.get_sheet(0)

# 2행부터 데이터 입력 (0-based index)
for i, (_, row) in enumerate(ss.iterrows(), 1):
    ws.write(i, 0, val(row, '수취인명'))       # 받는 분
    ws.write(i, 1, zipcode(row, '우편번호'))    # 우편번호
    ws.write(i, 2, val(row, '기본배송지'))      # 주소
    ws.write(i, 3, val(row, '상세배송지'))      # 상세주소
    ws.write(i, 4, '')                          # 일반전화
    ws.write(i, 5, val(row, '수취인연락처1'))   # 휴대전화
    ws.write(i, 6, 3)                           # 중량(kg)
    ws.write(i, 7, 80)                          # 부피(cm)
    ws.write(i, 8, '생활용품')                  # 내용품코드
    ws.write(i, 9, '')                          # 내용물
    ws.write(i, 10, '')                         # 배달방식
    ws.write(i, 11, val(row, '배송메세지'))     # 배송시요청사항
    ws.write(i, 12, 'N')                        # 분할접수 여부
    ws.write(i, 13, '')
    ws.write(i, 14, '')
    ws.write(i, 15, '')
    ws.write(i, 16, '')

today = datetime.now().strftime('%Y%m%d')
output_path = os.path.expanduser(f"~/smartstore-project/output/우체국택배_업로드용_{today}.xls")
wb.save(output_path)
print(f"✅ 변환 완료: {output_path} ({len(ss)}건)")
```

3. **완료 후 결과 보고**
   - 저장 경로와 변환된 주문 건수를 사용자에게 알려준다.
   - 오류 발생 시 원인을 설명하고 해결 방법을 안내한다.

## 고정값
- 중량: 3kg
- 부피: 80cm
- 내용품코드: 생활용품
- 분할접수: N
- 템플릿 경로: `~/smartstore-project/templates/post-form.xls`
- 결과 파일 저장: `~/smartstore-project/output/우체국택배_업로드용_YYYYMMDD.xls`
