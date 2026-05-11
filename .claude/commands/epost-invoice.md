우체국택배 송장 엑셀을 네이버 스마트스토어 출고 파일에 매칭하여 송장번호(등기번호)를 입력하고 발송처리용 파일을 생성한다.

> 실행 로직은 `~/smartstore-project/.cowork-skills/epost-flow/scripts/match_invoices.py` 에 있다.
> 로직을 바꾸려면 위 스크립트를 직접 수정한다.

## 사용법
인자: `$ARGUMENTS`
- 인자가 없으면 사용자에게 스마트스토어 파일 경로/비밀번호, 우체국 송장 파일 경로를 질문한다.
- 인자 형식: `스마트스토어파일경로 비밀번호 우체국송장파일경로`
  - 예: `~/Downloads/스마트스토어.xlsx 1111 ~/Downloads/2026050656246453_20260506.xls`

## 실행 절차

1. **라이브러리 설치**

```bash
pip3 install msoffcrypto-tool pandas openpyxl xlrd -q
```

2. **매칭 스크립트 실행**

```bash
python3 ~/smartstore-project/.cowork-skills/epost-flow/scripts/match_invoices.py "<스마트스토어_경로>" "<비밀번호>" "<우체국송장_경로>"
```

표준출력 형식:
- `DONE:<output_path>:<filled_count>` — 정상 종료, 매칭된 건수
- `UNMATCHED:이름1(전화1)|이름2(전화2)...` — 매칭 안 된 수취인이 있을 때

3. **결과 보고**
   - 저장 경로와 매칭된 건수를 사용자에게 알려준다.
   - `UNMATCHED:` 가 있으면 미매칭 수취인 목록을 보여준다.
   - 미매칭은 보통 다른 날 주문이 섞인 경우이므로 참고용으로만 안내.

## 매칭 기준
- **수취인명** (스마트스토어 `수취인명` ↔ 우체국 `받는분`)
- **전화번호** (스마트스토어 `수취인연락처1` ↔ 우체국 `받는분 전화번호`) — 하이픈 제거 후 비교

## 고정 처리 (스크립트 내부에서 수행)
- 1행(안내문) 삭제
- 시트명 `발주발송관리` → `발송처리`
- 송장번호 컬럼에 우체국 `등기번호` 값 입력
- 결과 파일: `~/smartstore-project/output/발송처리/스마트스토어_발송처리_YYYYMMDD.xlsx`
