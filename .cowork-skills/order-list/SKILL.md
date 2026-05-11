---
name: order-list
description: 네이버 스마트스토어 출고 엑셀에서 구매자(주문자)별 주문 품목 리스트를 영문 상품명으로 정리한 워드(.docx) 파일을 만든다. 사용자가 "주문자별 주문품목 리스트", "구매자별 주문 리스트", "주문자별 리스트", "order list", "buyer order list" 같은 표현을 쓰거나, 스마트스토어 주문을 사람용으로 보기 좋게 정리하고 싶어할 때 항상 이 스킬을 사용한다. 사용자가 영문 상품명에 대해 선호 표기가 있어 학습된 이름을 우선 적용한다.
---

# 주문자별 주문품목 리스트 (Word)

스마트스토어 출고 엑셀에서 구매자별로 주문 품목을 모아 영문 상품명으로 정리한 워드 파일을 만든다.

## 핵심 원칙

- **영문명 우선순위**:
  1. `preferred-names.json` 의 학습된 이름 (key = `상품번호||옵션정보`)
  2. `~/smartstore-project/templates/product-mapping.xlsx` 의 `상품명(영문)` 컬럼
  3. 한글 상품명 (둘 다 없을 때 폴백)
- **파일명 / 제목**:
  - 파일명: `핀치마트_YYYYMMDD.docx`
  - 문서 제목: `핀치마트 YYYY/M/D order list` (월/일은 0 패딩 안 함)
- **포맷 (사용자 선호)**:
  - 구매자명만 헤더로 사용. 주문번호 줄은 넣지 않음.
  - 한 구매자가 여러 주문번호를 가지면 모두 한 표에 합쳐서 보여줌.
  - 같은 (상품번호, 옵션, 주문번호) 조합은 수량 합산.
  - 수취인이 구매자와 다를 때만 헤더 옆에 작은 글씨로 "(수취인: 이름)" 표시.
  - Product / Qty 두 컬럼 표.
- **사용자가 매번 다르게 손으로 추가하는 주석은 학습하지 않는다**:
  - 헤더 옆 추적번호 (예: `정대원 (Amazon TBC869956907009)`) — 발송별 추적 번호
  - 사용자가 직접 수동 추가하는 행 (예: 스마트스토어에 없는 외부 주문자)
  - 단, 항목 영문명 자체에 사용자가 일관되게 붙이는 표시 (예: `(Amazon tracking n/a)`) 는 **그대로 학습**한다 — 사용자가 그 상품에 대해 늘 그렇게 표기하고 싶다는 신호.

## 사용법

인자: `$ARGUMENTS`
- 인자가 없으면 사용자에게 스마트스토어 파일 경로와 비밀번호를 질문한다.
- 인자 형식: `파일경로 비밀번호` — 비밀번호 없으면 빈 문자열.

## 실행

```bash
pip3 install msoffcrypto-tool pandas openpyxl xlrd python-docx -q
python3 ~/smartstore-project/.cowork-skills/order-list/scripts/generate.py "<스마트스토어_경로>" "<비밀번호>"
```

표준출력 형식:
- `DONE:<output_path>:<buyer_count>:<line_count>` — 정상 종료
- `MISSING_NAMES:상품번호1|상품번호2...` — 매핑·preferred 모두 없는 상품 (있을 때)

생성 파일은 `~/smartstore-project/output/` 에 `핀치마트_YYYYMMDD.docx` 로 저장된다.

## 결과 보고 후 학습 사이클 (중요)

1. 파일을 사용자에게 링크로 알려준다.
2. 사용자가 직접 워드 파일을 검토·수정할 수 있게 안내한다.
3. **사용자가 "내가 고쳤어, 학습해" 또는 비슷한 표현을 하면**:
   - 수정된 docx 를 pandoc 으로 텍스트 추출
   - 행 인덱스 순서대로 `preferred-names.json` 의 (상품번호 + 옵션정보) 키에 새 영문명을 매핑 (구매자별 행 순서는 스마트스토어 원본 행 순서를 따름)
   - 사용자가 손으로 추가한 헤더 옆 메모 (예: `정대원 (Amazon TBC...)`) 와 스마트스토어 데이터에 없는 행 (예: 외부 주문자 추가 행) 은 학습하지 않는다
   - 항목 영문명 자체는 사용자가 적은 그대로 저장 (괄호 주석 포함)
   - 다음 실행부터 자동 적용
4. **파일명/제목 패턴이 바뀌었으면** (예: `핀치마트_` 가 아닌 다른 접두사를 원함) `scripts/generate.py` 와 본 SKILL.md 양쪽을 갱신한다.

## preferred-names.json 형식

```json
{
  "13438803519||": {
    "product_id": "13438803519",
    "option": "",
    "kor_name": "오레오 케이크스터즈 ...",
    "preferred_eng": "Oreo Cakesters Soft Mini Cakes Original 285g"
  },
  "10044874416||종류: Hazelnut": {
    "product_id": "10044874416",
    "option": "종류: Hazelnut",
    "kor_name": "네스카페 (Nescafe) 리치 인스턴트 커피 - Hazelnut",
    "preferred_eng": "Nescafe Hazelnut"
  }
}
```

같은 상품번호라도 옵션이 다르면 (예: 네스카페 Hazelnut vs French Vanilla) 별개 키로 저장된다.

## 참고

- 비밀번호가 있는 스마트스토어 파일은 msoffcrypto 로 복호화한다.
- 결과 폴더(`~/smartstore-project/output/`)가 없으면 생성한다.
