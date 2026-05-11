#!/usr/bin/env python3
"""
스마트스토어 출고 엑셀 → 우체국택배 업로드용 .xls 변환.

사용법:
    python3 convert.py <스마트스토어_경로> <비밀번호>

비밀번호가 없으면 빈 문자열로 호출. 결과는 .xls (xlwt) 형식으로 저장됨 —
우체국택배 사이트가 .xlsx를 받지 않기 때문.
"""
import io
import os
import sys
from datetime import datetime

import msoffcrypto
import pandas as pd
import xlwt


def main():
    if len(sys.argv) < 2:
        print("Usage: convert.py <smartstore_path> [password]", file=sys.stderr)
        sys.exit(1)

    smartstore_file = sys.argv[1]
    password = sys.argv[2] if len(sys.argv) > 2 else ""

    # 스마트스토어 복호화
    decrypted = io.BytesIO()
    with open(os.path.expanduser(smartstore_file), "rb") as f:
        if password:
            office_file = msoffcrypto.OfficeFile(f)
            office_file.load_key(password=password)
            office_file.decrypt(decrypted)
            decrypted.seek(0)
        else:
            decrypted = io.BytesIO(f.read())
    ss = pd.read_excel(decrypted, engine="openpyxl", header=1)

    def val(row, col):
        v = row.get(col)
        return "" if pd.isna(v) else v

    def phone(row, col):
        v = val(row, col)
        return str(v).replace("-", "") if v != "" else ""

    def zipcode(row, col):
        v = val(row, col)
        if v == "":
            return ""
        s = str(v).replace(".0", "").strip()
        return s.zfill(5) if s.isdigit() else s

    # 주문번호 기준 중복 제거 (한 주문 = 한 행)
    ss["_order_num"] = ss["주문번호"].astype(str)
    ss_dedup = ss.drop_duplicates(subset="_order_num", keep="first").reset_index(drop=True)

    headers = [
        "받는 분", "우편번호", "주소(시도+시군구+도로명+건물번호)",
        "상세주소(동, 호수, 洞명칭, 아파트, 건물명 등)",
        "일반전화(02-1234-5678)", "휴대전화(010-1234-5678)",
        "중량(kg)", "부피(cm)=가로+세로+높이", "내용품코드", "내용물",
        "배달방식", "배송시요청사항", "분할접수 여부(Y/N)",
        "분할접수 첫번째 중량(kg)", "분할접수 첫번째 부피(cm)",
        "분할접수 두번째 중량(kg)", "분할접수 두번째 부피(cm)",
    ]

    rows = []
    for _, row in ss_dedup.iterrows():
        rows.append([
            val(row, "수취인명"),
            zipcode(row, "우편번호"),
            val(row, "기본배송지"),
            val(row, "상세배송지"),
            phone(row, "수취인연락처2"),
            phone(row, "수취인연락처1"),
            3,
            80,
            "생활용품",
            "",
            "",
            val(row, "배송메세지"),
            "N",
            "", "", "", "",
        ])

    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet("Sheet1")

    header_style = xlwt.easyxf(
        "font: name Arial, bold on, height 180, color white;"
        "pattern: pattern solid, fore_color dark_red;"
        "alignment: horizontal center, vertical center, wrap on;"
        "borders: left thin, right thin, top thin, bottom thin;"
    )
    text_left = xlwt.easyxf(
        "font: name Arial, height 180;"
        "alignment: horizontal left, vertical center;"
        "borders: left thin, right thin, top thin, bottom thin;"
    )
    text_center = xlwt.easyxf(
        "font: name Arial, height 180;"
        "alignment: horizontal center, vertical center;"
        "borders: left thin, right thin, top thin, bottom thin;"
    )
    zip_style = xlwt.easyxf(
        "font: name Arial, height 180;"
        "alignment: horizontal center, vertical center;"
        "borders: left thin, right thin, top thin, bottom thin;",
        num_format_str="@",
    )

    col_widths = [12, 8, 40, 22, 14, 14, 8, 8, 16, 10, 8, 24, 14, 14, 14, 14, 14]
    for ci, w in enumerate(col_widths):
        ws.col(ci).width = 256 * w

    ws.row(0).height_mismatch = True
    ws.row(0).height = 30 * 20

    for ci, h in enumerate(headers):
        ws.write(0, ci, h, header_style)

    center_cols = {0, 6, 7, 8, 12}

    for ri, row_data in enumerate(rows, 1):
        for ci, v in enumerate(row_data):
            if ci == 1:
                style = zip_style
            elif ci in center_cols:
                style = text_center
            else:
                style = text_left
            ws.write(ri, ci, v, style)

    today = datetime.now().strftime("%Y%m%d")
    output_dir = os.path.expanduser("~/smartstore-project/output/우체국택배")
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, f"우체국택배_업로드용_{today}.xls")
    wb.save(output_path)
    print(f"✅ 변환 완료: {output_path} ({len(rows)}건)")


if __name__ == "__main__":
    main()
