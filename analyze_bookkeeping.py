"""
2025 복식장부 데이터 분석 스크립트
파일: 2025 복식장부.xlsm > '2025 장부' 시트
"""

import openpyxl
from collections import defaultdict
from datetime import datetime

FILE_PATH = '2025 복식장부.xlsm'
SHEET_NAME = '2025 장부'

# 컬럼 인덱스 (0-based)
COL_DATE       = 0   # A: 날짜
COL_COMPANY    = 1   # B: 상호
COL_DESC       = 2   # C: 거래내역
COL_CLASS      = 3   # D: 구분
COL_VAT_FLAG   = 4   # E: 부가세 (o/x)
COL_TOTAL      = 5   # F: 합계금액
COL_SUPPLY     = 6   # G: 공급가액
COL_VAT        = 7   # H: 부가세금액
COL_PAYMENT    = 9   # J: 결제방법
COL_EVIDENCE   = 10  # K: 증빙 (세계/통장/카드)
COL_PARTNER    = 11  # L: 거래처
COL_DR1_ACC    = 12  # M: 차변Ⅰ 계정과목
COL_DR1_AMT    = 14  # O: 차변Ⅰ 금액
COL_DR1_TYPE   = 15  # P: 차변Ⅰ 거래요소
COL_DR2_ACC    = 16  # Q: 차변Ⅱ 계정과목
COL_DR2_AMT    = 18  # S: 차변Ⅱ 금액
COL_DR2_TYPE   = 19  # T: 차변Ⅱ 거래요소
COL_CR1_ACC    = 20  # U: 대변Ⅰ 계정과목
COL_CR1_AMT    = 22  # W: 대변Ⅰ 금액
COL_CR1_TYPE   = 23  # X: 대변Ⅰ 거래요소
COL_CR2_ACC    = 24  # Y: 대변Ⅱ 계정과목
COL_CR2_AMT    = 26  # AA: 대변Ⅱ 금액
COL_CR2_TYPE   = 27  # AB: 대변Ⅱ 거래요소
COL_BALANCE    = 28  # AC: 대차차액


def to_num(val):
    """값을 숫자로 변환. None이거나 수식 문자열이면 0 반환."""
    if val is None:
        return 0
    if isinstance(val, (int, float)):
        return float(val)
    if isinstance(val, str) and val.startswith('='):
        return 0
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0


def format_num(n):
    """숫자를 천 단위 구분 포맷으로 반환."""
    return f"{int(n):,}"


def main():
    print("=" * 60)
    print("  2025 복식장부 데이터 분석")
    print("=" * 60)

    wb = openpyxl.load_workbook(FILE_PATH, read_only=True, keep_vba=True, data_only=True)
    ws = wb[SHEET_NAME]

    rows = []
    for i, row in enumerate(ws.iter_rows(values_only=True), 1):
        if i <= 3:
            continue  # 헤더/설명 행 스킵
        date = row[COL_DATE]
        if date is None:
            continue
        # 2025년 데이터만 필터링 (이월 항목 제외)
        if isinstance(date, datetime) and date.year != 2025:
            continue
        rows.append(row)

    print(f"\n[기본 통계]")
    print(f"  총 거래 건수 (2025년): {len(rows):,}건")

    # 월별 집계
    monthly_total   = defaultdict(float)
    monthly_debit   = defaultdict(float)
    monthly_credit  = defaultdict(float)
    monthly_vat     = defaultdict(float)

    # 계정과목별 집계 (차변/대변 각각)
    debit_accounts  = defaultdict(float)
    credit_accounts = defaultdict(float)

    # 거래처별 집계
    partner_totals  = defaultdict(float)

    # 증빙 유형별
    evidence_count  = defaultdict(int)

    # 거래요소별
    element_debit   = defaultdict(float)
    element_credit  = defaultdict(float)

    total_supply = 0
    total_vat = 0

    for row in rows:
        date = row[COL_DATE]
        month = date.month if isinstance(date, datetime) else 0

        total_amt  = to_num(row[COL_TOTAL])
        supply_amt = to_num(row[COL_SUPPLY])
        vat_amt    = to_num(row[COL_VAT])

        dr1_acc  = row[COL_DR1_ACC]
        dr1_amt  = to_num(row[COL_DR1_AMT])
        dr1_type = row[COL_DR1_TYPE]
        dr2_acc  = row[COL_DR2_ACC]
        dr2_amt  = to_num(row[COL_DR2_AMT])
        dr2_type = row[COL_DR2_TYPE]
        cr1_acc  = row[COL_CR1_ACC]
        cr1_amt  = to_num(row[COL_CR1_AMT])
        cr1_type = row[COL_CR1_TYPE]
        cr2_acc  = row[COL_CR2_ACC]
        cr2_amt  = to_num(row[COL_CR2_AMT])
        cr2_type = row[COL_CR2_TYPE]

        partner  = row[COL_PARTNER] or '(미입력)'
        evidence = row[COL_EVIDENCE] or '(미입력)'

        monthly_total[month]  += total_amt
        monthly_debit[month]  += dr1_amt + dr2_amt
        monthly_credit[month] += cr1_amt + cr2_amt
        monthly_vat[month]    += vat_amt

        total_supply += supply_amt
        total_vat    += vat_amt

        if dr1_acc and dr1_amt:
            debit_accounts[dr1_acc] += dr1_amt
        if dr2_acc and dr2_amt:
            debit_accounts[dr2_acc] += dr2_amt
        if cr1_acc and cr1_amt:
            credit_accounts[cr1_acc] += cr1_amt
        if cr2_acc and cr2_amt:
            credit_accounts[cr2_acc] += cr2_amt

        if dr1_type and dr1_amt:
            element_debit[dr1_type] += dr1_amt
        if dr2_type and dr2_amt:
            element_debit[dr2_type] += dr2_amt
        if cr1_type and cr1_amt:
            element_credit[cr1_type] += cr1_amt
        if cr2_type and cr2_amt:
            element_credit[cr2_type] += cr2_amt

        partner_totals[partner] += total_amt
        evidence_count[evidence] += 1

    print(f"  총 공급가액 합계: {format_num(total_supply)}원")
    print(f"  총 부가세 합계:   {format_num(total_vat)}원")

    # 월별 집계 출력
    print(f"\n[월별 거래 현황]")
    print(f"  {'월':>3}  {'합계금액':>15}  {'차변합계':>15}  {'대변합계':>15}  {'부가세':>12}")
    print(f"  {'-'*3}  {'-'*15}  {'-'*15}  {'-'*15}  {'-'*12}")
    for m in sorted(monthly_total.keys()):
        print(f"  {m:>2}월  {format_num(monthly_total[m]):>15}  "
              f"{format_num(monthly_debit[m]):>15}  "
              f"{format_num(monthly_credit[m]):>15}  "
              f"{format_num(monthly_vat[m]):>12}")

    # 차변 계정과목 상위 10
    print(f"\n[차변 계정과목 상위 10]")
    sorted_dr = sorted(debit_accounts.items(), key=lambda x: x[1], reverse=True)[:10]
    for rank, (acc, amt) in enumerate(sorted_dr, 1):
        print(f"  {rank:>2}. {acc:<20} {format_num(amt):>15}원")

    # 대변 계정과목 상위 10
    print(f"\n[대변 계정과목 상위 10]")
    sorted_cr = sorted(credit_accounts.items(), key=lambda x: x[1], reverse=True)[:10]
    for rank, (acc, amt) in enumerate(sorted_cr, 1):
        print(f"  {rank:>2}. {acc:<20} {format_num(amt):>15}원")

    # 거래처별 상위 10
    print(f"\n[거래처별 거래금액 상위 10]")
    sorted_pt = sorted(partner_totals.items(), key=lambda x: x[1], reverse=True)[:10]
    for rank, (partner, amt) in enumerate(sorted_pt, 1):
        print(f"  {rank:>2}. {partner:<25} {format_num(amt):>15}원")

    # 증빙 유형별
    print(f"\n[증빙 유형별 건수]")
    for ev, cnt in sorted(evidence_count.items(), key=lambda x: x[1], reverse=True):
        print(f"  {ev:<10}  {cnt:>5}건")

    # 거래요소별 (차변)
    if element_debit:
        print(f"\n[거래요소별 차변 합계]")
        for elem, amt in sorted(element_debit.items(), key=lambda x: x[1], reverse=True):
            print(f"  {elem:<10}  {format_num(amt):>15}원")

    # 거래요소별 (대변)
    if element_credit:
        print(f"\n[거래요소별 대변 합계]")
        for elem, amt in sorted(element_credit.items(), key=lambda x: x[1], reverse=True):
            print(f"  {elem:<10}  {format_num(amt):>15}원")

    print("\n" + "=" * 60)
    print("  분석 완료")
    print("=" * 60)


if __name__ == '__main__':
    main()
