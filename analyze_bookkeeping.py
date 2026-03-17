"""
2025 복식장부 데이터 분석 스크립트
파일: 2025 복식장부.xlsm > '2025 장부' 시트
"""

import openpyxl
from collections import defaultdict
from datetime import datetime

FILE_PATH = '2025 복식장부.xlsm'
SHEET_NAME = '2025 장부'

# 거래처 명칭 정규화 매핑 (다양한 표기 → 대표 명칭)
PARTNER_NORMALIZE = {
    # 티에이케이텍스타일
    '티에이케이텍스타일 주식회사': '티에이케이텍스타일(주)',
    '티에이케이텍스타일(': '티에이케이텍스타일(주)',
    '티에이케이텍스타일(?': '티에이케이텍스타일(주)',
    '티에이케이텍스타일': '티에이케이텍스타일(주)',
    # 실론
    '㈜실론': '(주)실론',
    '주식회사 실론': '(주)실론',
    '주식회사\u3000실론': '(주)실론',
    '(주)실론': '(주)실론',
    # 하이멜
    '하이멜': '하이멜',
    '문정환(하이멜)': '하이멜',
    '문정환하이멜': '하이멜',
    '하이멜문정환': '하이멜',
    # 에프에스
    '（주）에프에스예약': '（주）에프에스',
    # 은호섬유
    '（주）은호섬유예약': '（주）은호섬유',
    # 한국윈텍
    '(주) 한국윈텍': '(주)한국윈텍',
    # 성일티앤씨
    '성일티앤씨 주식회사': '성일티앤씨(주)',
    '성일티앤씨（주）예약': '성일티앤씨(주)',
    '성일티앤씨（주）': '성일티앤씨(주)',
    '성일티앤씨(주)': '성일티앤씨(주)',
    # 미래심지
    '（주）미래심지예약': '（주）미래심지',
    # 원창머티리얼
    '원창머티리얼(주)비산공장': '원창머티리얼(주)',
    '원창머티리얼（주）': '원창머티리얼(주)',
    # 성원텍스
    '（주）성원텍스（예약': '（주）성원텍스',
    # 유진트레이딩
    '유진트레이딩（주예약': '유진트레이딩(주)',
    '유진트레이딩 주식회사': '유진트레이딩(주)',
    # 디케이티앤씨
    '(주)디케이티앤씨': '(주)디케이티앤씨',
    '（주）디케이티앤예약': '(주)디케이티앤씨',
    '（주）디케이티앤씨': '(주)디케이티앤씨',
    '주식회사 디케이티앤씨': '(주)디케이티앤씨',
    # 에스제이인터내셔널
    '에스제이인터내': '에스제이인터내셔널(주)',
    '에스제이인터내셔널(주)': '에스제이인터내셔널(주)',
    # 유진폴리텍크
    '(주) 유진폴리텍크': '(주)유진폴리텍크',
    '（주）유진폴리텍크': '(주)유진폴리텍크',
    # 두성섬유
    '두성섬유(주)': '두성섬유(주)',
    '두성섬유（주）예약': '두성섬유(주)',
    # 승진섬유
    '황경모（승진섬유예약': '승진섬유산업',
    # 풍아산업
    '풍아산업（주）예약': '풍아산업(주)',
    '풍아산업㈜': '풍아산업(주)',
    '풍아산업 주식회사': '풍아산업(주)',
    # 에스티
    '주식회사 에스티': '(주)에스티',
    '（주）에스티예약': '(주)에스티',
    # 에스티유
    '주식회사에스티유': '(주)에스티유',
    '（주）에스티유예약': '(주)에스티유',
    # 무역대전송금 (공백 제거)
    '무역대전송금 ': '무역대전송금',
    # 유남부직포
    '류영철(유남부직포)': '유남부직포',
    # 아이온인더스트리
    '(주)아이온  인더': '(주)아이온인더스트리',
    # 아이디에프엘
    '아이디에프엘래버러': '아이디에프엘',
    '아이디에프엘 래버러토리 앤드 인스티튜트 인크 (영업소)': '아이디에프엘',
    '아이디에프엘래예약': '아이디에프엘',
    # 아주피피얀
    '주식회사 아주피피얀': '(주)아주피피얀',
    '（주）아주피피얀예약': '(주)아주피피얀',
    # 보현엠앤티
    '보현엠앤티 주식회사': '보현엠앤티(주)',
    '보현엠앤티주식회예약': '보현엠앤티(주)',
    # 아진티앤엘
    '（주）아진티앤엘': '(주)아진티앤엘',
    # 대출이자 (공백 제거)
    '대출이자    ': '대출이자',
    # 제이더블유관세법인
    '제이더블유 관세법인': '제이더블유관세법인',
    # SF익스프레스코리아
    '에스에프익스프레스코리아 주식회사( SF Express Korea Co.,Ltd.)': 'SF익스프레스코리아(주)',
    'ＳＦ익스프레': 'SF익스프레스코리아(주)',
    # 태광텍스타일
    '（주）태광텍스타예약': '（주）태광텍스타일',
    '（주）태광텍스타일（': '（주）태광텍스타일',
    # 일신항공해운
    '（주）일신항공해예약': '(주)일신항공해운',
    '（주）일신항공해운': '(주)일신항공해운',
    # 건승연사
    '주식회사 건승연사': '건승연사',
    # 미래지오
    '주식회사 미래지오': '(주)미래지오',
    '（주）미래지오예약': '(주)미래지오',
    # 유정인더스트리
    '공남순(유정인더스트': '유정인더스트리',
    # 세안물류
    '（주）세안물류예약': '(주)세안물류',
    # 우전섬유
    '주식회사 우전섬유': '(주)우전섬유',
    '（주）우전섬유예약': '(주)우전섬유',
    # 진코퍼레이션
    '진코퍼레이션（주）': '진코퍼레이션(주)',
    '진코퍼레이션（주예약': '진코퍼레이션(주)',
    # 강일섬유
    '김강웅（강일섬유예약': '강일섬유',
    # 삼덕섬유
    '최만식（삼덕섬유예약': '삼덕섬유',
    # 안국섬유
    '차후영（안국섬유예약': '안국섬유',
    # 대한실업
    '김종호（대한실업예약': '대한실업',
    # P.I.P
    '（주）Ｐ．Ｉ．Ｐ예약': '(주)P.I.P',
    '（주）P.I.P': '(주)P.I.P',
    # 소망법무사
    '소망 법무사사무소': '소망법무사',
    '김효근（소망법무': '소망법무사',
    # 제이앤제이
    '제이엔제이(J&J)': '제이앤제이',
    # 관세법인세종
    '관세법인 세종 인천공항지사': '관세법인세종인천',
    # 삼영빌딩
    '백영길삼영빌딩예약': '삼영빌딩',
    # 기보
    '기보＿홉아이엔씨（주': '기보_홉아이엔씨(주)',
    '기보＿홉아이엔씨': '기보_홉아이엔씨(주)',
    # 진흥
    '주식회사진흥': '주식회사 진흥',
    # 비씨카드
    '비씨카드 ': '비씨카드',
    '비씨카드출금': '비씨카드',
    # 한국에스지에스
    '한국에스지에스(주)': '한국에스지에스(주)',
    '한국에스지에스（예약': '한국에스지에스(주)',
    '한국에스지에스（주）': '한국에스지에스(주)',
    # KB/국민카드 (동일 카드사)
    'KB카드출금': 'KB국민카드',
    '국민카드': 'KB국민카드',
    # 서현운수
    '서현운수신동호': '서현운수',
    # 니즈텔레콤
    '주식회사 니즈텔레콤': '니즈텔레콤',
    # 합산보험 (월별 코드 통합)
    '합산보험2501': '합산보험', '합산보험2502': '합산보험', '합산보험2503': '합산보험',
    '합산보험2504': '합산보험', '합산보험2505': '합산보험', '합산보험2506': '합산보험',
    '합산보험2507': '합산보험', '합산보험2508': '합산보험', '합산보험2509': '합산보험',
    '합산보험2510': '합산보험', '합산보험2511': '합산보험', '합산보험2512': '합산보험',
    # KTNET (월별 코드 통합)
    'KTNET-01': 'KTNET', 'KTNET-02': 'KTNET', 'KTNET-03': 'KTNET',
    'KTNET-04': 'KTNET', 'KTNET-05': 'KTNET', 'KTNET-06': 'KTNET',
    'KTNET-07': 'KTNET', 'KTNET-08': 'KTNET', 'KTNET-09': 'KTNET',
    'KTNET-10': 'KTNET', 'KTNET-11': 'KTNET', 'KTNET-12': 'KTNET',
    # KT 통신비 (월별 코드 통합)
    'KT0529748901': 'KT통신비', 'KT0529748902': 'KT통신비', 'KT0529748903': 'KT통신비',
    'KT0529748904': 'KT통신비', 'KT0529748905': 'KT통신비', 'KT0529748906': 'KT통신비',
    'KT0529748907': 'KT통신비', 'KT0529748908': 'KT통신비', 'KT0529748909': 'KT통신비',
    'KT0529748910': 'KT통신비', 'KT0529748911': 'KT통신비', 'KT0529748912': 'KT통신비',
    # 도명섬유
    '이상용（도명섬유예약': '도명섬유',
    # 보람섬유
    '민경욱（보람섬유）': '보람섬유',
}


def normalize_partner(name):
    """거래처명을 정규화하여 동일 거래처를 통합."""
    if name is None:
        return '(미입력)'
    name = str(name).strip()
    return PARTNER_NORMALIZE.get(name, name)


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

        partner  = normalize_partner(row[COL_PARTNER])
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
    print(f"\n[거래처별 거래금액 상위 10] (통합 후 고유 거래처 수: {len(partner_totals):,}개)")
    sorted_pt = sorted(partner_totals.items(), key=lambda x: x[1], reverse=True)[:10]
    for rank, (partner, amt) in enumerate(sorted_pt, 1):
        print(f"  {rank:>2}. {partner:<30} {format_num(amt):>15}원")

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
