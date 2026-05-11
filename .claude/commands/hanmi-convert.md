네이버 스마트스토어 출고 준비중 엑셀 파일을 한미택배 업로드용 엑셀 파일로 변환한다.

> 실행 로직은 `~/smartstore-project/.cowork-skills/hanmi-flow/scripts/convert.py` 에 있다.
> 로직을 바꾸려면 위 스크립트를 직접 수정한다.

## 사용법
인자: `$ARGUMENTS`
- 인자가 없으면 사용자에게 스마트스토어 파일 경로와 비밀번호를 질문한다.
- 인자 형식: `파일경로 비밀번호` (예: `~/Downloads/스마트스토어.xlsx 1111`)
- 비밀번호가 없으면 빈 문자열로 처리한다.

## 실행 절차

1. **라이브러리 설치**

```bash
pip3 install msoffcrypto-tool pandas openpyxl xlrd -q
```

2. **변환 스크립트 실행**

```bash
python3 ~/smartstore-project/.cowork-skills/hanmi-flow/scripts/convert.py "<스마트스토어_경로>" "<비밀번호>"
```

표준출력 형식:
- `MISSING:상품번호 XXX: 한글명\n상품번호 YYY: 한글명` — 매핑 누락 상품이 있을 때
- `✅ 변환 완료: <output_path> (N건)` — 정상 종료

3. **매핑 누락 상품 처리** (MISSING 라인이 있을 때만)

   **3-1. 자동 유추** — 누락된 상품마다 한글 상품명을 분석해 직접 유추:
   - **영문 상품명**: 한글명에서 브랜드·제품·용량을 조합한 자연스러운 영문명
   - **HS CODE**: `~/smartstore-project/templates/hs-code-reference.xlsx` 참고
   - **브랜드**: 상품명에서 추출 (예: "팀홀튼 믹스커피" → "Tim Hortons")
   - **해외판매자 상호**: 사이트 도메인 대문자 (예: `costco.ca` → `COSTCO`). 모르면 브랜드명

   **3-2. 사용자 확인 요청** — 유추 내용을 한 번에 보여준다:

   ```
   [새 상품 확인] 상품번호 XXXXXXX
   한글명: (원본)
   영문명: (유추) ← 맞나요?
   HS Code: (유추) ← 맞나요?
   브랜드: (유추) ← 맞나요?

   확인되면 아래 두 가지만:
   - 단가 (CAD 숫자)
   - 구매 사이트 URL (예: www.costco.ca)
   ```

   **SITE URL 자동 조합**: `https://smartstore.naver.com/finchmart_ca/ (입력받은_URL)`.

   - 사용자가 영문명/HS Code/브랜드를 수정하면 그 값을 사용
   - 확인만 하면 유추값 그대로

   **3-3. product-mapping.xlsx 에 행 추가**

   ```python
   import os
   from openpyxl import load_workbook

   PRODUCT_MAPPING = os.path.expanduser('~/smartstore-project/templates/product-mapping.xlsx')
   wb_map = load_workbook(PRODUCT_MAPPING)
   ws_map = wb_map.active

   new_entries = [
       # (상품번호, 한글상품명, hs_code, 영문상품명, 브랜드, 단가, site_url, 해외판매자상호)
   ]
   for entry in new_entries:
       ws_map.append(list(entry))
   wb_map.save(PRODUCT_MAPPING)
   ```

   저장 후 **2번 변환 스크립트를 다시 실행**해 최종 파일을 생성한다.

4. **결과 보고**
   - 저장 경로와 변환된 주문 건수를 사용자에게 알려준다.
   - 오류 발생 시 원인 설명 + 해결 방법 안내.

## 참고
- 발송인 정보: `~/smartstore-project/config.json` (gitignore, `config-example.json` 복사해 작성)
- 한미택배 양식: `~/smartstore-project/templates/hanmi-form.xls`
- 상품 매핑: `~/smartstore-project/templates/product-mapping.xlsx`
- 결과 저장 경로: `~/smartstore-project/output/한미택배/`
