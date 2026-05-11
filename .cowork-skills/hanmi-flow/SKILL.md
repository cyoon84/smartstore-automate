---
name: hanmi-flow
description: 한미택배 전체 발송 워크플로 한 번에 실행. 네이버 스마트스토어 출고 엑셀을 한미택배 업로드용으로 변환하고, 사용자가 한미 사이트에 업로드 후 받아온 송장 엑셀로 발송처리 파일까지 만들고, #all-finchmart Slack 채널에 처리 건수를 알린다. 사용자가 "한미플로우", "한미 발송", "스마트스토어 발송 처리", "hanmi flow", "한미택배 처리" 같은 표현을 쓰거나, 스마트스토어 출고 엑셀을 한미택배로 보내려고 할 때 항상 이 스킬을 사용한다.
---

# 한미택배 전체 발송 루틴

이 스킬은 두 단계로 나눠진 한 번의 흐름이다. 단계 사이에 사용자가 한미택배 사이트에 직접 업로드하는 시간이 필요하므로, 첫 단계 결과를 보고한 뒤 사용자가 송장 파일을 가져올 때까지 기다린다.

## 흐름 개요

```
[1단계 변환]  스마트스토어 출고 엑셀  →  한미택배 업로드용 엑셀
       ↓
   사용자가 한미택배 사이트에 업로드 → 송장 엑셀 다운로드 (대기)
       ↓
[2단계 매칭]  스마트스토어 + 한미 송장  →  발송처리 엑셀
       ↓
[3단계 알림]  #all-finchmart Slack 채널에 처리 건수 메시지
```

## 1단계 — 변환

사용자에게 두 가지를 묻는다 (인자로 같이 들어오면 그 값을 사용):
- 스마트스토어 출고 엑셀 파일 경로
- 비밀번호 (없으면 빈 값으로)

물은 다음 `scripts/convert.py`를 실행한다:

```bash
pip3 install msoffcrypto-tool pandas openpyxl xlrd -q
python3 ~/.claude/skills/hanmi-flow/scripts/convert.py "<스마트스토어_경로>" "<비밀번호>"
```

스크립트 표준출력 형식:
- `MISSING:상품번호 XXX: 한글명\n상품번호 YYY: 한글명` — 매핑에 없는 상품이 있을 때
- `✅ 변환 완료: <output_path> (N건)` — 정상 종료

### 매핑 누락 상품 처리

`MISSING:` 라인이 있으면 누락된 상품마다 한글명을 분석해 다음을 **유추한다**.
**프로젝트 공통 규칙은 `~/smartstore-project/.cowork-skills/conventions.md` 를 먼저 본다** (영문명 대소문자, 브랜드별 SITE URL 매핑 등).

- **영문명**: 한글 상품명에서 브랜드·제품·용량을 조합한 자연스러운 영문명. **ALL CAPS** 로 작성 (conventions.md).
- **HS CODE**: `~/smartstore-project/templates/hs-code-reference.xlsx` 를 참고해 가장 적합한 코드
- **브랜드**: 상품명에서 추출 (예: "팀홀튼 믹스커피" → "Tim Hortons")
- **해외판매자 상호**: 사이트 도메인을 대문자로 (예: `costco.ca` → `COSTCO`). URL 모르면 브랜드명으로
- **SITE URL**: conventions.md 의 브랜드 매핑 규칙을 따른다 (예: Nespresso Vertuo capsule 은 항상 `www.nespresso.ca`). 규칙에 없으면 사용자에게 묻거나 기본 추정 사용.

유추한 내용을 한 번에 보여주고 확인을 요청한다:

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

**SITE URL은 자동 조합**: `https://smartstore.naver.com/finchmart_ca/ (입력받은_URL)` 형식.

확인된 내용으로 `~/smartstore-project/templates/product-mapping.xlsx`에 행을 추가한다:

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

저장 후 **변환 스크립트를 다시 실행**해서 최종 파일을 만든다.

### 1단계 종료 보고

변환이 끝나면 다음 형식으로 사용자에게 보고하고 **여기서 멈춘다**:

```
✅ 1단계 변환 완료
파일: ~/smartstore-project/output/한미택배/한미택배_업로드용_YYYYMMDD.xlsx
변환 건수: N건

이제 한미택배 사이트에 업로드해 주세요.
업로드 후 받은 송장 엑셀 파일을 첨부하거나 경로를 알려주시면 2단계를 진행합니다.
```

사용자의 다음 메시지에서 송장 파일이 올 때까지 능동적으로 다음 단계로 넘어가지 않는다.

## 2단계 — 송장 매칭

사용자가 한미 송장 파일(예: `송장리스트-Apr21.xls`)을 첨부하거나 경로를 주면, 이전에 받았던 스마트스토어 경로/비밀번호를 그대로 재사용해서 `scripts/match_invoices.py`를 실행한다:

```bash
python3 ~/.claude/skills/hanmi-flow/scripts/match_invoices.py "<스마트스토어_경로>" "<비밀번호>" "<한미송장_경로>"
```

표준출력 형식:
- `DONE:<output_path>:<filled_count>` — 정상 종료, 매칭된 건수
- `UNMATCHED:이름1(전화1)|이름2(전화2)...` — 매칭 안 된 수취인이 있을 때

미매칭은 보통 다른 날 주문이 섞여 있는 경우이므로 참고용으로만 안내한다.

## 3단계 — Slack 알림

매칭이 끝나면 `#all-finchmart` 채널에 메시지를 보낸다. `slack_send_message` 도구를 사용:

```
channel: all-finchmart  (또는 채널 ID)
text:
한미택배 발송처리 완료 ✅
• 변환: {convert_count}건
• 송장 매칭: {filled_count}건
• 미매칭: {unmatched_count}건 (다른 날 주문 가능성)
• 결과: 스마트스토어_발송처리_YYYYMMDD.xlsx
```

채널을 이름으로 바로 찾지 못하면 `slack_search_channels` 로 `all-finchmart`를 검색해 채널 ID를 얻은 뒤 보낸다.

## 마무리 보고

사용자에게 짧게 정리:

```
✅ 한미플로우 완료
• 변환: {convert_count}건
• 매칭: {filled_count}건 (미매칭 {unmatched_count}건)
• 발송처리 파일: ~/smartstore-project/output/발송처리/스마트스토어_발송처리_YYYYMMDD.xlsx
• Slack #all-finchmart 알림 전송 완료
```

## 참고

- 발송인 정보는 `~/smartstore-project/config.json` 에서 읽음 (BUSINESS_ID, SENDER_NAME 등). 이 파일은 gitignore 됨.
- 한미택배 양식: `~/smartstore-project/templates/hanmi-form.xls`
- 상품 매핑: `~/smartstore-project/templates/product-mapping.xlsx`
- 매칭 기준: 수취인명 + 수취인연락처1 (하이픈 제거 후 비교)
- 이 스킬은 변환과 매칭을 **별도 단계로** 처리해야 한다. 두 단계 사이에는 반드시 사용자의 한미택배 사이트 업로드 시간이 필요하므로, 1단계 끝에서 멈추고 사용자 입력을 기다린다.
