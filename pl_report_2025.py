"""
홉아이엔씨 2025년도 손익보고서 생성 스크립트
데이터 소스: 2025 복식장부.xlsm > '(법인)손익계산서' 시트
"""

import openpyxl
from datetime import datetime

FILE_PATH = '2025 복식장부.xlsm'
SHEET_NAME = '(법인)손익계산서'


def w(n):
    """원 단위 포맷 (소수점 반올림)"""
    if n is None:
        return "0"
    return f"{round(n):,}"


def pct(part, total):
    """비율(%) 계산"""
    if not total:
        return "  -  "
    return f"{part / total * 100:5.1f}%"


def read_pl_data():
    """손익계산서 시트에서 계정과목별 금액을 코드 기준으로 읽어온다."""
    wb = openpyxl.load_workbook(FILE_PATH, read_only=True, keep_vba=True, data_only=True)
    ws = wb[SHEET_NAME]

    data = {}  # code → (name, amount)

    for row in ws.iter_rows(values_only=True):
        # 페이지1: 컬럼 A(0)=계정명, B(1)=코드, D(3)=금액
        if row[1] is not None and isinstance(row[1], int) and row[3] is not None:
            data[row[1]] = (str(row[0]).strip(), row[3])
        # 페이지2: 컬럼 E(4)=계정명, G(6)=코드, H(7)=금액
        if row[6] is not None and isinstance(row[6], int) and row[7] is not None:
            data[row[6]] = (str(row[4]).strip() if row[4] else '', row[7])
        # 페이지3: 컬럼 J(9)=계정명, K(10)=코드, M(12)=금액
        if row[10] is not None and isinstance(row[10], int) and len(row) > 12 and row[12] is not None:
            data[row[10]] = (str(row[9]).strip() if row[9] else '', row[12])
        # 페이지4: 컬럼 N(13)=계정명, P(15)=코드, Q(16)=금액
        if len(row) > 16 and row[15] is not None and isinstance(row[15], int) and row[16] is not None:
            data[row[15]] = (str(row[13]).strip() if row[13] else '', row[16])

    return data


def get(data, code):
    """코드에 해당하는 금액 반환 (없으면 0)"""
    item = data.get(code)
    return round(item[1]) if item and item[1] is not None else 0


def main():
    data = read_pl_data()

    # ── 매출액 ──
    rev_domestic   = get(data, 6)    # 가.국내제품매출
    rev_export     = get(data, 7)    # 나.수출제품매출
    rev_product    = get(data, 5)    # 2.제품매출
    rev_goods      = get(data, 2)    # 1.상품매출
    rev_total      = get(data, 1)    # Ⅰ.매출액

    # ── 매출원가 ──
    cogs_goods     = get(data, 36)   # (1)상품매출원가
    cogs_mfg       = get(data, 42)   # (2)제조원가
    cogs_total     = get(data, 35)   # Ⅱ.매출원가

    # ── 매출총손익 ──
    gross_profit   = get(data, 66)   # Ⅲ.매출총손익

    # ── 판매비와관리비 ──
    sga_salary     = get(data, 68)   # 1.급여
    sga_retire     = get(data, 74)   # 2.퇴직급여
    sga_insurance  = get(data, 78)   # 3.보험료
    sga_welfare    = get(data, 79)   # 4.복리후생비
    sga_travel     = get(data, 80)   # 5.여비교통비
    sga_rent       = get(data, 81)   # 6.임차료
    sga_entertain  = get(data, 85)   # 7.접대비
    sga_depr       = get(data, 86)   # 8.유형자산감가상각비
    sga_tax        = get(data, 90)   # 10.세금과공과
    sga_advert     = get(data, 91)   # 11.광고선전비
    sga_vehicle    = get(data, 93)   # 13.차량유지비
    sga_research   = get(data, 94)   # 14.연구비
    sga_fee        = get(data, 104)  # 20.지급수수료
    sga_supplies   = get(data, 108)  # 22.소모품비
    sga_telecom    = get(data, 109)  # 23.통신비
    sga_freight    = get(data, 110)  # 24.운반비
    sga_storage    = get(data, 111)  # 25.보관료
    sga_utilities  = get(data, 114)  # 28.수도광열비
    sga_customs    = get(data, 119)  # 33.수출입제비용
    sga_outsrc_dom = get(data, 122)  # 35가.국내외주용역비
    sga_outsrc_ovs = get(data, 123)  # 35나.해외외주용역비
    sga_outsrc     = get(data, 121)  # 35.외주용역비
    sga_misc       = get(data, 124)  # 36.기타판관비
    sga_total      = get(data, 67)   # Ⅳ.판매비와관리비

    # ── 영업손익 ──
    operating_inc  = get(data, 129)  # Ⅴ.영업손익

    # ── 영업외수익 ──
    non_op_rev_int = get(data, 131)  # 1.이자수익
    non_op_rev     = get(data, 130)  # Ⅵ.영업외수익

    # ── 영업외비용 ──
    non_op_exp_int = get(data, 180)  # 1.이자비용
    non_op_exp_etc = get(data, 212)  # 21.기타 영업외비용
    non_op_exp     = get(data, 179)  # Ⅶ.영업외비용

    # ── 법인세비용차감전손익 / 당기순손익 ──
    ebt            = get(data, 217)  # Ⅷ.법인세비용차감전손익
    tax            = get(data, 218)  # Ⅸ.법인세비용
    net_income     = get(data, 219)  # Ⅹ.당기순손익

    SEP1 = "=" * 65
    SEP2 = "-" * 65
    sep3 = " " + "-" * 63

    print()
    print(SEP1)
    print("         홉아이엔씨  2025년도 손익보고서")
    print("         과세기간: 2025.01.01 ~ 2025.12.31  (단위: 원)")
    print(SEP1)
    print(f"  {'계  정  과  목':<36}  {'금  액':>15}  {'비율':>6}")
    print(SEP2)

    # 매출액
    print(f"  {'Ⅰ. 매출액':<36}  {w(rev_total):>15}  {pct(rev_total, rev_total)}")
    print(f"    {'국내제품매출':<34}  {w(rev_domestic):>15}  {pct(rev_domestic, rev_total)}")
    print(f"    {'수출제품매출':<34}  {w(rev_export):>15}  {pct(rev_export, rev_total)}")
    print(SEP2)

    # 매출원가
    print(f"  {'Ⅱ. 매출원가':<36}  {w(cogs_total):>15}  {pct(cogs_total, rev_total)}")
    if cogs_goods:
        print(f"    {'상품매출원가':<34}  {w(cogs_goods):>15}  {pct(cogs_goods, rev_total)}")
    print(f"    {'제조기초재고액':<34}  {w(cogs_mfg):>15}  {pct(cogs_mfg, rev_total)}")
    print(SEP2)

    # 매출총이익
    sign = "▲" if gross_profit < 0 else " "
    print(f"  {'Ⅲ. 매출총이익':<36}  {sign}{w(abs(gross_profit)):>14}  {pct(gross_profit, rev_total)}")
    print(SEP2)

    # 판관비
    print(f"  {'Ⅳ. 판매비와관리비':<36}  {w(sga_total):>15}  {pct(sga_total, rev_total)}")
    items = [
        ("외주용역비 (국내)",  sga_outsrc_dom),
        ("외주용역비 (해외)",  sga_outsrc_ovs),
        ("기타판관비",         sga_misc),
        ("급여",               sga_salary),
        ("지급수수료",         sga_fee),
        ("차량유지비",         sga_vehicle),
        ("여비교통비",         sga_travel),
        ("연구비",             sga_research),
        ("수출입제비용",       sga_customs),
        ("임차료",             sga_rent),
        ("접대비",             sga_entertain),
        ("운반비",             sga_freight),
        ("복리후생비",         sga_welfare),
        ("수도광열비",         sga_utilities),
        ("보관료",             sga_storage),
        ("세금과공과",         sga_tax),
        ("통신비",             sga_telecom),
        ("소모품비",           sga_supplies),
        ("광고선전비",         sga_advert),
        ("보험료",             sga_insurance),
        ("퇴직급여",           sga_retire),
        ("유형자산감가상각비", sga_depr),
    ]
    for name, amt in items:
        if amt:
            print(f"    {name:<34}  {w(amt):>15}  {pct(amt, rev_total)}")
    print(SEP2)

    # 영업이익
    sign = "▲" if operating_inc < 0 else " "
    print(f"  {'Ⅴ. 영업이익':<36}  {sign}{w(abs(operating_inc)):>14}  {pct(operating_inc, rev_total)}")
    print(SEP2)

    # 영업외수익
    print(f"  {'Ⅵ. 영업외수익':<36}  {w(non_op_rev):>15}  {pct(non_op_rev, rev_total)}")
    if non_op_rev_int:
        print(f"    {'이자수익':<34}  {w(non_op_rev_int):>15}")
    print(sep3)

    # 영업외비용
    print(f"  {'Ⅶ. 영업외비용':<36}  {w(non_op_exp):>15}  {pct(non_op_exp, rev_total)}")
    if non_op_exp_int:
        print(f"    {'이자비용':<34}  {w(non_op_exp_int):>15}  {pct(non_op_exp_int, rev_total)}")
    if non_op_exp_etc:
        print(f"    {'기타 영업외비용':<34}  {w(non_op_exp_etc):>15}  {pct(non_op_exp_etc, rev_total)}")
    print(SEP2)

    # 법인세비용차감전손익
    sign = "▲" if ebt < 0 else " "
    print(f"  {'Ⅷ. 법인세비용차감전손익':<36}  {sign}{w(abs(ebt)):>14}  {pct(ebt, rev_total)}")
    if tax:
        print(f"  {'Ⅸ. 법인세비용':<36}  {w(tax):>15}")
    print(SEP2)

    # 당기순손익
    sign = "▲" if net_income < 0 else " "
    print(SEP1)
    print(f"  {'Ⅹ. 당기순이익':<36}  {sign}{w(abs(net_income)):>14}  {pct(net_income, rev_total)}")
    print(SEP1)

    # 요약 박스
    print()
    print("  ┌─ 핵심 지표 요약 ──────────────────────────────────────┐")
    print(f"  │  매출액           {w(rev_total):>18}원             │")
    print(f"  │  매출총이익       {w(gross_profit):>18}원  ({pct(gross_profit,rev_total)})  │")
    print(f"  │  영업이익         {w(operating_inc):>18}원  ({pct(operating_inc,rev_total)})  │")
    print(f"  │  당기순이익       {w(net_income):>18}원  ({pct(net_income,rev_total)})  │")
    print(f"  │  국내매출 비중    {pct(rev_domestic,rev_total):>6}  /  수출 비중 {pct(rev_export,rev_total):>6}             │")
    print(f"  │  외주용역비 비중  {pct(sga_outsrc,rev_total):>6}  (판관비 중 최대 항목)           │")
    print("  └──────────────────────────────────────────────────────┘")
    print()


if __name__ == '__main__':
    main()
