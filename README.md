# smartstore-project

네이버 스마트스토어 출고 처리를 자동화하는 Claude Code 커맨드 모음.

## 워크플로우

```
스마트스토어 출고준비중 엑셀
        ↓ /hanmi-convert
한미택배 업로드용 엑셀 → 한미택배 사이트 업로드
        ↓ (한미에서 송장 엑셀 다운로드)
        ↓ /hanmi-invoice
스마트스토어 발송처리용 엑셀 → 셀러센터 엑셀 일괄발송 업로드
```

## 커맨드

### `/hanmi-convert` — 한미택배 업로드 파일 생성
스마트스토어 출고준비중 엑셀을 한미택배 업로드 양식으로 변환한다.

```
/hanmi-convert 파일경로 비밀번호
예: /hanmi-convert ~/Downloads/스마트스토어_출고준비중.xlsx 1111
```

**출력:** `output/한미택배_업로드용_YYYYMMDD.xlsx`

---

### `/hanmi-invoice` — 발송처리 파일 생성
한미택배 송장 엑셀을 스마트스토어 출고 파일에 매칭하여 송장번호를 입력한다.

```
/hanmi-invoice 스마트스토어파일 비밀번호 한미송장파일
예: /hanmi-invoice ~/Downloads/스마트스토어.xlsx 1111 ~/Downloads/송장리스트-Apr21.xls
```

- 수취인명 + 전화번호로 매칭
- 택배사는 원본 유지 (우체국택배)
- 1행 삭제, 시트명 `발주발송관리` → `발송처리` 자동 처리

**출력:** `output/스마트스토어_발송처리_YYYYMMDD.xlsx`

---

## 초기 설정

### 1. config.json 생성
```bash
cp config-example.json config.json
```
`config.json`을 열어 발송인 정보 입력 (gitignore 처리됨):

| 키 | 설명 |
|---|---|
| `business_id` | 한미택배 비즈니스 회원 아이디 |
| `sender_name` | 보내는 사람 영문 이름 |
| `sender_email` | 이메일 |
| `sender_phone` | 전화번호 (하이픈 없이) |
| `sender_address` | 발송지 주소 (영문) |

### 2. 템플릿 파일 확인
| 파일 | 설명 |
|---|---|
| `templates/hanmi-form.xls` | 한미택배 업로드 양식 |
| `templates/product-mapping.xlsx` | 상품번호 → 영문명/HS CODE/브랜드/단가 매핑 |

`product-mapping.xlsx` 컬럼: `상품번호`, `상품명`, `상품명(영문)`, `HS CODE`, `브랜드`, `단가`, `SITE URL`, `해외판매자 상호`
