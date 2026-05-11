네이버 스마트스토어 출고 준비중 엑셀 파일을 우체국택배 업로드용 .xls 파일로 변환한다.

> 실행 로직은 `~/smartstore-project/.cowork-skills/epost-flow/scripts/convert.py` 에 있다.
> 로직을 바꾸려면 위 스크립트를 직접 수정한다.

## 사용법
인자: `$ARGUMENTS`
- 인자가 없으면 사용자에게 스마트스토어 파일 경로와 비밀번호를 질문한다.
- 인자 형식: `파일경로 비밀번호` (예: `~/Downloads/스마트스토어.xlsx 1111`)
- 비밀번호가 없으면 빈 문자열로 처리한다.

## 실행 절차

1. **라이브러리 설치**

```bash
pip3 install msoffcrypto-tool pandas openpyxl xlrd xlwt -q
```

2. **변환 스크립트 실행**

```bash
python3 ~/smartstore-project/.cowork-skills/epost-flow/scripts/convert.py "<스마트스토어_경로>" "<비밀번호>"
```

표준출력 형식:
- `✅ 변환 완료: <output_path> (N건)` — 정상 종료

3. **결과 보고**
   - 저장 경로와 변환된 주문 건수를 사용자에게 알려준다.
   - 오류 발생 시 원인 설명 + 해결 방법 안내.

## 고정값 (스크립트 내부)

| 컬럼 | 값 |
|------|-----|
| 중량 (G) | 3kg |
| 부피 (H) | 80 |
| 내용품코드 (I) | 생활용품 |
| 분할접수 여부 (M) | N |

## 참고
- 주문번호 기준으로 중복 제거 — 한 주문에 여러 상품이 있어도 한 행만 생성
- 우편번호는 텍스트 형식 (앞자리 0 보존)
- **저장 형식은 .xls (xlwt)** — 우체국택배 사이트가 .xlsx를 받지 않음
- 결과 저장 경로: `~/smartstore-project/output/우체국택배/`
