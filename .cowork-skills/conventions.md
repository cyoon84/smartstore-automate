# 핀치마트 프로젝트 규칙

스마트스토어 / 한미택배 / 우체국택배 / 주문 리스트 등 모든 스킬이 공통으로 따라야 할 규칙. 새로운 상품을 `product-mapping.xlsx` 에 추가하거나 영문명을 만들 때 적용한다.

## 영문 상품명 (product-mapping.xlsx 의 `상품명(영문)`)

- **ALL CAPS** 로 작성한다. (예: `NUTELLA PEANUT CHOCOLATE SPREAD 725G`)
- 사용자가 명시적으로 다른 표기를 지정하면 그 값을 우선한다.
- 기존 항목은 자동 일괄 변경하지 않는다 (요청이 있을 때만).

## SITE URL 매핑 (브랜드 → 사이트)

특정 브랜드는 항상 같은 사이트로 매핑한다:

| 브랜드 | SITE URL 안에 들어가는 도메인 |
|---|---|
| Nespresso Vertuo capsule (모든 종류) | `www.nespresso.ca` |

SITE URL 최종 형식은 `https://smartstore.naver.com/finchmart_ca/ (도메인)` 이다.

브랜드 규칙에 해당하지 않으면 사용자에게 묻거나 기본 추정 (`www.costco.ca`, `www.walmart.ca` 등) 을 사용한다.

## 해외판매자 상호

- 대문자, 도메인의 핵심 부분만 사용 (예: `www.walmart.ca` → `WALMART`).
- SITE URL 의 도메인과 일치해야 한다.

## 학습된 영문명 (order-list 스킬)

- `order-list/preferred-names.json` 의 `preferred_eng` 값은 사용자가 docx 에서 적은 그대로 저장한다 — 이 파일은 사람용 표기이므로 ALL CAPS 규칙을 적용하지 않는다.
- `product-mapping.xlsx` 와 `preferred-names.json` 은 별도 소스이고, order-list 는 preferred-names 를 먼저 본다.
