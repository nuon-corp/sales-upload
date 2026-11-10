"""채널별 매출 정산 엑셀 → 아마란스 '데이터 업로드 양식' 변환기

사용법
  1) '원본' 폴더에 채널에서 받은 정산 엑셀을 넣고
  2) 변환실행.bat 더블클릭 (또는 python convert_sales.py)
  → '결과' 폴더에 채널별 업로드 파일이 생깁니다.

채널은 파일 '내용'을 보고 자동 인식합니다(파일명 무관).
지원 채널: SK스토아, NS홈쇼핑, 넛지헬스케어(캐시딜), 오늘의집, 공영홈쇼핑, 우리홈쇼핑(롯데홈쇼핑), CJ온스타일,
          롯데ON, 11번가(정산확정건 + 부가세신고내역상세 2개 파일)

규칙은 2026년 2월·9월 데이터와 경영지원팀 업로드 파일을 대조해 정했습니다.
  - 매출/수수료는 원본의 거래(행) 단위로 한 줄씩 올림
  - 부가세 포함 금액은 행마다 ÷1.1 후 원 단위 반올림 (부가세 = 합계 - 공급가)
  - 원본에 공급가/부가세가 따로 있으면 그 값을 그대로 사용
"""
import sys, os, re, shutil, glob, traceback
from collections import defaultdict
from decimal import Decimal, ROUND_HALF_UP
import openpyxl

# 규칙을 고치면 이 날짜도 바꿔 주세요 — 웹 화면 위쪽에 표시돼서 어떤 버전으로 변환했는지 알 수 있어요.
RULES_VERSION = '2026-11-10c'

BASE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(BASE, '데이터 업로드 양식.xlsx')
IN_DIR = os.path.join(BASE, '원본')
OUT_DIR = os.path.join(BASE, '결과')

# ── 과세구분 규칙 (회계 확인 후 여기만 고치면 됨) ─────────────────────────
# 홈쇼핑 위수탁 매출: 경영지원팀 방식대로 전액 14(건별)
HOMESHOP_TAX = '14'
# 넛지 캐시딜 결제수단 → 과세구분 (경영지원팀 업로드 파일로 아직 확인 못 함)
NUDGE_PAY_TO_TAX = {
    '신용카드': '17', '네이버페이': '17', '카카오페이': '17', '토스페이': '17',
    '페이코': '17', 'SSG페이': '17', '엘페이': '17',
    '제로페이': '14', '계좌이체': '14', '적립금': '14',
}
# 오늘의집 부가세신고 화면 결제수단 → 과세구분 (경영지원팀 파일과 일치 확인)
OHOUSE_PAY_TO_TAX = {'카드': '17', '현금': '31', '기타': '14'}
DEFAULT_TAX = '14'
TAX_NAME = {'14': '건별', '17': '카드', '31': '현금'}


# ── 공통 유틸 ──────────────────────────────────────────────────────────
def r0(x):
    return int(Decimal(str(x)).quantize(Decimal('1'), ROUND_HALF_UP))


def split_vat(total, taxable=True):
    """부가세 포함 금액 → (공급가, 부가세)."""
    if not taxable:
        return total, 0
    sup = r0(Decimal(total) / Decimal('1.1'))
    return sup, total - sup


def vat10(sup):
    return r0(Decimal(sup) * Decimal('0.1'))


def num(v):
    if v in (None, '', '-'):
        return 0
    return int(float(str(v).replace(',', '')))


def isdigit(v):
    return str(v if v is not None else '').strip().isdigit()


def find_header(rows, *cols):
    return next(i for i, r in enumerate(rows) if r and all(c in r for c in cols))


def header_index(row):
    return {str(v).strip(): i for i, v in enumerate(row) if v is not None}


def load_book(path):
    """엑셀 파일 열기. .xlsx는 openpyxl, 옛날 .xls는 xlrd로 읽어 같은 모양으로 바꿈."""
    if path.lower().endswith('.xls'):
        try:
            import xlrd
        except ImportError:
            raise RuntimeError('.xls 파일을 읽으려면 xlrd가 필요합니다 (변환기 실행.bat을 다시 실행하면 자동 설치)')
        src = xlrd.open_workbook(path)
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        for sh in src.sheets():
            ws = wb.create_sheet(sh.name[:31])
            for i in range(sh.nrows):
                row = [None if v == '' else v for v in sh.row_values(i)]
                ws.append(row if any(v is not None for v in row) else [None])
        return wb
    return openpyxl.load_workbook(path, data_only=True)


def period_of(words):
    """제목에 적힌 조회기간 '(2026/10/01~2026/10/31)' → 짝 맞추기 키."""
    for w in words:
        m = re.search(r'(\d{4}[./-]\d{2}[./-]\d{2})\s*~\s*(\d{4}[./-]\d{2}[./-]\d{2})', w)
        if m:
            return m.group(1).replace('.', '/').replace('-', '/') + '~' + m.group(2).replace('.', '/').replace('-', '/')
    return ''


def sheet(wb, prefix):
    return next((s for s in wb if s.title.replace(' ', '').startswith(prefix.replace(' ', ''))), None)


class Book:
    """업로드 파일 한 개 분량 (시트별 행 목록)."""
    def __init__(self):
        self.rows = {'매출': [], '광고': [], '수수료': []}
        self.tax = defaultdict(lambda: [0, 0, 0])   # 과세구분 → [공급가, 부가세, 부가세포함 누적] (화면 합계용)
        self.tax_rows = []                           # 부가세 시트에 건별로 쓸 행 [과세구분, 공급가, 부가세]
        self.notes = []   # 확인이 필요한 사항 (가정, 미반영 시트)
        self.info = []    # 참고 안내 (확인된 규칙대로 제외한 것 등)
        self.companions = {}   # 함께 올린 보조 파일 (예: 11번가 부가세신고내역) {키: [(파일명, wb)]}

    def add(self, sheet_name, code, name, sup, vat):
        if sup == 0 and vat == 0:
            return
        code = '' if code in (None, '-') else str(code).strip()
        self.rows[sheet_name].append([code, str(name or '').strip(), int(sup), int(vat)])

    def add_gross(self, sheet_name, code, name, total, taxable=True):
        self.add(sheet_name, code, name, *split_vat(total, taxable))

    def add_tax(self, tax_code, sup, vat):
        self.tax[tax_code][0] += sup; self.tax[tax_code][1] += vat

    def add_tax_gross(self, tax_code, total):
        """결제수단별 합계를 모아 마지막에 한 번 ÷1.1."""
        self.tax[tax_code][2] += total

    def add_tax_row(self, tax_code, sup, vat):
        """부가세 시트에 거래 1건 = 1행으로 기록."""
        if sup == 0 and vat == 0:
            return
        self.tax_rows.append([tax_code, int(sup), int(vat)])
        self.add_tax(tax_code, sup, vat)

    def tax_from_sales(self, tax_code):
        """매출 시트 행마다 같은 금액으로 부가세 시트 행을 만듦 (건별)."""
        for r in self.rows['매출']:
            self.add_tax_row(tax_code, r[2], r[3])

    def finalize(self):
        for d in self.tax.values():
            if d[2]:
                s, v = split_vat(d[2]); d[0] += s; d[1] += v; d[2] = 0


# ── 채널별 변환 ────────────────────────────────────────────────────────
def parse_sk(wb, bk):
    """SK스토아(위탁): 매출 = 고객결제금액, 수수료 = 위수탁수수료(공급가/세액)."""
    ws = sheet(wb, '수수료')
    hdr = None
    for r in ws.iter_rows(values_only=True):
        if r[0] == 'No' and r[1] == '업체코드' and '주문번호' in r:
            hdr = header_index(r); continue
        if hdr and isdigit(r[0]):
            code, name = r[hdr['상품코드']], r[hdr['상품명']]
            bk.add_gross('매출', code, name, num(r[hdr['고객결제금액']]), r[hdr['과세여부']] != '면세')
            bk.add('수수료', code, name, num(r[hdr['위수탁수수료(공급가)']]), num(r[hdr['위수탁수수료(세액)']]))
    bk.tax_from_sales(HOMESHOP_TAX)
    rows = list(ws.iter_rows(values_only=True))
    hi = next((i for i, r in enumerate(rows) if r and '고객부담배송비(+)' in r), None)
    if hi is not None:
        c = list(rows[hi]).index('고객부담배송비(+)')
        ship = sum(num(r[c]) for r in rows[hi + 1:hi + 3] if r and isdigit(r[0]))
        if ship:
            bk.notes.append(f'배송비 확인: 정산종합내역서에 고객부담배송비 {ship:,}원 — 이번 변환에는 미반영')
    bk.notes.append('SK스토아: 경영지원팀 파일로 아직 대조 못 함 (매출 총액 기준·부가세 14 가정)')


def parse_ns(wb, bk):
    """NS홈쇼핑(특약매입): 매출 = 매출금액-반품금액, 부가세 = 단가부가세×수량, 광고 = 쇼핑북광고비."""
    ws = sheet(wb, '품목별')
    rows = list(ws.iter_rows(values_only=True))
    hi = find_header(rows, '상품코드', '매출금액')
    h = header_index(rows[hi])
    for r in rows[hi + 1:]:
        if not r[h['상품코드']]:
            continue
        qty = num(r[h['매출수량']]) - num(r[h['반품수량']])
        sup = num(r[h['매출금액']]) - num(r[h['반품금액']])
        vat = 0 if r[h['과세구분']] == '면세' else num(r[h['부가세']]) * qty
        bk.add('매출', r[h['상품코드']], r[h['상품명']], sup, vat)
    ad = sheet(wb, '쇼핑북광고비')
    if ad:
        rows = list(ad.iter_rows(values_only=True))
        hi = find_header(rows, '상품코드', '광고비')
        h = header_index(rows[hi])
        for r in rows[hi + 1:]:
            if r[h['상품코드']] and num(r[h['광고비']]):
                amt = num(r[h['광고비']])
                bk.add('광고', r[h['상품코드']], r[h['상품명']], amt, vat10(amt))
    ship = sheet(wb, '고객부담배송비')
    if ship:
        srows = list(ship.iter_rows(values_only=True))
        si = next((i for i, r in enumerate(srows) if r and '배송비 합계' in r), None)
        sh = header_index(srows[si]) if si is not None else {}
        srows = srows[si:] if si is not None else []
        amt = sum(num(r[sh['배송비 합계']]) for r in srows[1:] if '배송비 합계' in sh and r and r[sh['배송비 합계']] not in (None, ''))
        n = sum(1 for r in srows[1:] if r and any(v not in (None, '') for v in r))
        if amt or n:
            bk.notes.append(f'배송비 확인: 고객부담배송비 시트에 내역 있음 (합계 {amt:,}원) — 이번 변환에는 미반영')
    fee = sheet(wb, '특약수수료')
    if fee and any(r[0] for r in fee.iter_rows(min_row=2, values_only=True)):
        bk.notes.append('특약수수료 시트에 내역이 있음 — 처리 규칙 확인 필요 (이번 변환에는 미반영)')
    bk.notes.append('NS홈쇼핑: 경영지원팀 파일로 아직 대조 못 함 (세금계산서 매출이라 부가세 시트 비움, 광고비는 VAT 별도 가정)')


def parse_nudge(wb, bk):
    """넛지헬스케어 캐시딜: 매출 = 상품합계(A), 수수료 = 판매수수료(VAT 포함 가정)."""
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    h = header_index(rows[0])
    bigo = len(rows[0]) - 1
    ship = 0
    for r in rows[1:]:
        if r[bigo] != '최초 주문 상품행':
            if r[bigo] and '배송비' in str(r[bigo]):
                ship += num(r[h['상품 합계(A)']]) + num(r[h['정산금액(F=A-C-D-E)']])
            continue
        code, name = r[h['상품 번호']], r[h['상품명']]
        taxable = r[h['과세유형']] != '면세'
        sup, vat = split_vat(num(r[h['상품 합계(A)']]), taxable)
        bk.add('매출', code, name, sup, vat)
        bk.add_tax_row(NUDGE_PAY_TO_TAX.get(r[h['결제수단']], DEFAULT_TAX), sup, vat)
        bk.add_gross('수수료', code, name, num(r[h['판매수수료(C=A*B 또는 C=A-X)']]))
    if ship:
        bk.notes.append(f'배송비 확인: 배송비 행에 금액 {ship:,}원 — 이번 변환에는 미반영')
    bk.notes.append('넛지: 경영지원팀 파일로 아직 대조 못 함 (결제수단→과세구분, 수수료 VAT 포함 가정)')


def parse_ohouse(wb, bk):
    """오늘의집: 매출·수수료 = 품목별 내역 행 단위, 할인 = 부가세신고 화면의 매출할인, 부가세 = 결제수단별 합계."""
    vat_ws = sheet(wb, 'VAT')
    vrows = list(vat_ws.iter_rows(values_only=True))
    vh = header_index(vrows[find_header(vrows, '주문옵션', '결제금액', '결제수단')])
    opt_code = {}
    discounts = []
    for r in vrows:
        kind = r[vh['구분']]
        if kind not in ('주문정산', '매출할인'):
            continue
        amt = num(r[vh['결제금액']])
        if kind == '주문정산' and r[vh['주문옵션']] not in (None, '-'):
            opt_code[str(r[vh['주문옵션']])] = r[vh['상품번호']]
        if kind == '매출할인' and amt:
            discounts.append((r[vh['상품명']] or r[vh['정산구분']], amt))
        if amt:
            bk.add_tax_gross(OHOUSE_PAY_TO_TAX.get(r[vh['결제수단']], DEFAULT_TAX), amt)

    ws = sheet(wb, '품목별')
    rows = list(ws.iter_rows(values_only=True))
    hi = find_header(rows, '주문옵션번호', '상품명')
    h = header_index(rows[hi])
    # 원본 헤더가 '상품 판매가'부터 한 칸씩 밀려 있음 → 값은 각 헤더 오른쪽 칸
    def col(prefix):
        return next(i for k, i in h.items() if k.replace('\n', '').startswith(prefix)) + 1
    c_base, c_ship, c_asm, c_fee = col('정산 기준금액'), col('배송비'), col('조립비'), col('수수료')
    ship_in_items = 0
    for r in rows[hi + 1:]:
        if r[h['구분']] != '주문정산':
            continue
        if r[h['상품명']] != '클레임 배송비':
            ship_in_items += num(r[c_ship])
        pay = num(r[c_base]) + num(r[c_ship]) + num(r[c_asm])
        code = None   # 경영지원팀 방식: 상품코드 비움
        bk.add_gross('매출', code, r[h['상품명']], pay)
        bk.add_gross('수수료', code, r[h['상품명']], num(r[c_fee]))
    if ship_in_items:
        bk.notes.append(f'배송비 확인: 상품 행에 배송비 {ship_in_items:,}원이 있어 해당 상품 매출에 포함함 — 따로 빼야 하는지 확인')
    for name, amt in discounts:   # 오늘의집 부담 할인: 결제금액·수수료 같은 금액만큼 감소
        bk.add_gross('매출', None, name, amt)
        bk.add_gross('수수료', None, name, amt)
    ad = sheet(wb, '광고비')
    if ad and any((isinstance(v, (int, float)) and v) or (isdigit(v) and int(v)) for r in ad.iter_rows(values_only=True) for v in r[3:4]):
        bk.info.append('광고비 시트는 경영지원팀 방식대로 업로드에서 제외')


def parse_gongyoung(wb, bk):
    """공영홈쇼핑: 상품별매출내역(차수별 행) — 매출순금액/VAT, 수수료금액/VAT 그대로. 광고 = 카드무이자 협력사부담."""
    ws = sheet(wb, '품목별')
    rows = list(ws.iter_rows(values_only=True))
    hi = find_header(rows, '상품ID', '매출순금액')
    h = header_index(rows[hi])
    for r in rows[hi + 1:]:
        if not isdigit(r[h['No']]):
            break
        code, name = r[h['상품ID']], r[h['상품명']]
        bk.add('매출', code, name, num(r[h['매출순금액']]), num(r[h['매출순금액VAT']]))
        bk.add('수수료', code, name, num(r[h['수수료금액']]), num(r[h['수수료VAT']]))
    bk.tax_from_sales(HOMESHOP_TAX)
    mu = sheet(wb, '무이자')
    if mu:
        rows = list(mu.iter_rows(values_only=True))
        h = header_index(rows[find_header(rows, '상품번호', '협력사부담금액')])
        total = defaultdict(int); names = {}
        for r in rows:
            if isdigit(r[h['상품번호']]) and r[h['상품명']]:
                k = str(r[h['상품번호']]).strip()
                total[k] += num(r[h['협력사부담금액']]); names[k] = r[h['상품명']]
        for k, amt in total.items():   # 상품별 합계로 1행 (경영지원팀 방식)
            bk.add_gross('광고', k, names[k], amt)
    if sheet(wb, '반품비용'):
        bk.info.append('반품비용 입금 시트는 경영지원팀 방식대로 업로드에서 제외')


def parse_woori(wb, bk):
    """우리홈쇼핑(롯데홈쇼핑): 방송 수수료매출현황(VAT별도) — 매출 = 판매금액, 수수료 = 매입수수료, 부가세 10%."""
    ws = sheet(wb, '품목별')
    rows = list(ws.iter_rows(values_only=True))
    hi = find_header(rows, '주문번호', '판매금액', '매입수수료')
    h = header_index(rows[hi])
    for r in rows[hi + 1:]:
        if r[h['순번']] == '합계':
            break
        if not isdigit(r[h['순번']]):
            continue
        code, name = r[h['상품코드']], r[h['상품명']]
        taxable = r[h['과세']] != '면세'
        sup, fee = num(r[h['판매금액']]), num(r[h['매입수수료']])
        bk.add('매출', code, name, sup, vat10(sup) if taxable else 0)
        bk.add('수수료', code, name, fee, vat10(fee))
    bk.tax_from_sales(HOMESHOP_TAX)
    bk.notes.append('우리홈쇼핑: 경영지원팀 파일로 아직 대조 못 함 (부가세 행별 10%·부가세 14 가정, 인터넷 매출·ARS 미반영)')


def parse_cj(wb, bk):
    """CJ온스타일: 수수료 매출 일별내역 — 소비자공급가/부가세, CJ수수료/부가세 그대로. 광고 = 공제내역의 광고비."""
    ws = sheet(wb, '건별매출')
    rows = list(ws.iter_rows(values_only=True))
    hi = find_header(rows, 'No', '상품코드', '판매상품명')
    h = header_index(rows[hi])
    base, cj = h['협력사매출'], h['CJ온스타일매출(수수료매출)']
    for r in rows[hi + 2:]:
        if not isdigit(r[h['No']]):
            continue
        code, name = r[h['상품코드']], r[h['판매상품명']]
        bk.add('매출', code, name, num(r[base]), num(r[base + 1]))
        bk.add('수수료', code, name, num(r[cj]), num(r[cj + 1]))
    bk.tax_from_sales(HOMESHOP_TAX)
    gj = sheet(wb, '공제내역(')
    if gj:
        rows = list(gj.iter_rows(values_only=True))
        h = header_index(rows[find_header(rows, '공제명', '공제금액')])
        for r in rows:
            if r[h['공제명']] and '광고' in str(r[h['공제명']]) and num(r[h['공제금액']]):
                bk.add_gross('광고', None, r[h['공제명']], num(r[h['공제금액']]))
    if gj:
        grows = list(gj.iter_rows(values_only=True))
        start = next((i for i, r in enumerate(grows) if any(isinstance(v, str) and '저단가배송비' in v for v in r)), None)
        if start is not None:
            gh = header_index(grows[start + 1]) if start + 1 < len(grows) else {}
            amt = sum(num(r[gh['공제액']]) for r in grows[start + 2:] if '공제액' in gh and isdigit(r[1]))
            if amt:
                bk.notes.append(f'배송비 확인: 저단가배송비 공제 {amt:,}원 — 이번 변환에는 미반영')
    ti = sheet(wb, '세금계산서매출')
    if ti and any(isdigit(r[1]) for r in ti.iter_rows(values_only=True) if len(r) > 1):
        bk.notes.append('세금계산서매출(입고) 시트에 내역이 있음 — 처리 규칙 확인 필요 (미반영)')


LOTTEON_PAY_TO_TAX = {'현금영수증': '31', '신용카드': '17', '휴대폰': '14', '기타': '14'}


def parse_lotteon(wb, bk):
    """롯데ON(롯데쇼핑 중개): 매출 = 정산대상판매금액(배송비 포함), 수수료 = 총 수수료 합계, 둘 다 행별 ÷1.1.
    부가세 = 부가세신고자료조회의 결제수단별 금액 (현금영수증 31 / 신용카드 17 / 휴대폰·기타 14)."""
    ws = sheet(wb, '품목별')
    rows = list(ws.iter_rows(values_only=True))
    hi = find_header(rows, '판매자상품명', '정산대상판매금액')
    h = header_index(rows[hi])
    for r in rows[hi + 1:]:
        if not isdigit(r[h['번호']]):
            break
        code, name = None, r[h['판매자상품명']]   # 경영지원팀 방식: 상품코드 비움
        taxable = r[h['과세유형']] != '면세'
        bk.add_gross('매출', code, name, num(r[h['정산대상판매금액']]), taxable)
        bk.add_gross('수수료', code, name, num(r[h['총 수수료 합계']]))
    vs = sheet(wb, 'VAT')
    vrows = list(vs.iter_rows(values_only=True))
    vi = find_header(vrows, '총매출금액', '신용카드', '현금영수증')
    vh = header_index(vrows[vi])
    for r in vrows[vi + 1:]:
        if r[vh['매출월']] in (None, ''):
            break
        for pay, code in LOTTEON_PAY_TO_TAX.items():   # 결제수단마다 ÷1.1 후 과세구분별로 합침 (경영지원팀 방식)
            if pay in vh:
                bk.add_tax(code, *split_vat(num(r[vh[pay]])))
    for title in ('광고비', '셀러포인트'):
        ws2 = sheet(wb, title)
        if ws2 and any((isinstance(v, (int, float)) and v) for r in ws2.iter_rows(values_only=True) for v in r):
            bk.notes.append(f'{title} 시트에 금액이 있음 — 업로드 규칙 확인 필요 (이번 변환에는 미반영)')


# 11번가 수수료 = 서비스이용료(상품)~서비스이용료(선결제배송비) − 11번가할인 (+ '추가' 시트의 서버이용료) → 11번가 수수료 세금계산서 합계와 일치
ST11_DISC_IN_FEE = -1         # 11번가할인을 수수료에서 뺌(-1) — 11번가 세금계산서 서비스이용료가 할인 차감 후 금액 (2026-11-10 확인)
ST11_SERVER_FEE = 0           # 자동 추가 안 함 — 서버이용료는 원본의 '추가' 시트에 적기 (경영지원팀 결정 2026-11-10)
ST11_ADJ_TAX = '14'   # 11번가 매출조정(11번가할인)을 차감할 과세구분 — 14(건별) (경영지원팀 확인 2026-11-10)
ST11_PAY_TO_TAX = {'신용카드결제': '17', '현금영수증(소득공제용)': '31', '현금영수증(지출증빙용)': '31',
                   '휴대폰결제': '14', '기타결제금액': '14'}


def parse_11st(wb, bk):
    """11번가: 정산확정건 파일 + (별도 파일) 부가세신고내역상세.
    매출 = 판매금액합계 - 할인쿠폰이용료(판매자할인) - [공제]스토어묶음할인 - 11번가할인(11번가 부담)
           → 부가세신고 '과세매출금액 + 매출조정금액'과 일치 (11번가할인은 매출에서 제외: 경영지원팀 확인 2026-11-03)
    수수료 = 서비스이용료(상품·선결제배송비) + 제휴마케팅대행 + 기타 이용료, 광고 = 후불광고비 (모두 행별 ÷1.1)
    부가세 = 부가세신고내역의 결제수단별 합계 (신용카드 17 / 현금영수증 31 / 휴대폰·기타 14)"""
    ws = next(w for w in wb if has(sheet_words(w), '판매금액합계') and has(sheet_words(w), '공제금액합계'))
    rows = list(ws.iter_rows(values_only=True))
    hi = find_header(rows, '주문번호', '판매금액합계', '공제금액합계')
    h = {str(k).replace('\n', ''): v for k, v in header_index(rows[hi]).items()}
    def g(r, key):
        return num(r[h[key]]) if key in h else 0
    fee_cols = ['서비스이용료(상품)', '서비스이용료(선결제배송비)', '제휴마케팅대행(상품)', '제휴마케팅대행(선결제배송비)',
                '포인트이용료', '칩이용료', '무이자할부이용료', '지정택배이용료', '물류이용수수료', '반품보험료',
                '수출대행수수료', '전세계배송 판매자책임반품']
    known = set(fee_cols) | {'할인쿠폰이용료', '[공제]스토어묶음할인', '후불광고비',
                             '판매자기본할인', '판매자추가할인', '11번가할인', '복수구매할인비용', '판매자부담카드추가할인'}
    sale_t = fee_t = ad_t = ded_t = st11_disc = 0
    fee_svc = fee_ship = 0   # 11번가 수수료 세금계산서 품목별 합계 (서비스이용료 / 서비스이용료(선결제배송비))
    for r in rows[hi + 1:]:
        if not isdigit(r[h['NO']]):
            continue
        code, name = r[h['상품번호']], r[h['상품명']]
        sale = g(r, '판매금액합계') - g(r, '할인쿠폰이용료') - g(r, '[공제]스토어묶음할인') - g(r, '11번가할인')
        fee = sum(g(r, c) for c in fee_cols) + ST11_DISC_IN_FEE * g(r, '11번가할인')
        ad = g(r, '후불광고비')
        bk.add_gross('매출', code, name, sale)
        bk.add_gross('수수료', code, name, fee)
        bk.add_gross('광고', code, name, ad)
        fee_svc += g(r, '서비스이용료(상품)') + g(r, '제휴마케팅대행(상품)') + ST11_DISC_IN_FEE * g(r, '11번가할인')
        fee_ship += g(r, '서비스이용료(선결제배송비)') + g(r, '제휴마케팅대행(선결제배송비)')
        sale_t += sale; fee_t += fee; ad_t += ad; ded_t += g(r, '공제금액합계'); st11_disc += g(r, '11번가할인')
    # 부가세대급금이 11번가 세금계산서와 같도록: 세금계산서는 품목 합계마다 ÷1.1 하므로 행별 반올림 차이를 가장 큰 행에서 보정
    others = fee_t - fee_svc - fee_ship
    target_vat = sum(split_vat(x)[1] for x in (fee_svc, fee_ship, others) if x)
    fee_rows = bk.rows['수수료']
    diff = target_vat - sum(x[3] for x in fee_rows)
    if diff and fee_rows and abs(diff) <= len(fee_rows):
        big = max(fee_rows, key=lambda x: x[2])
        big[3] += diff; big[2] -= diff
        bk.info.append(f'수수료 부가세 반올림 차이 {diff:+}원을 가장 큰 행에서 보정 (세금계산서 기준)')
    ex = extra_sheet(wb)
    if ex is not None and any('서버이용료' in str(v) for r in ex.iter_rows(values_only=True) for v in r if v):
        pass   # '추가' 시트에 적어 둔 서버이용료를 씀 (자동 추가 안 함)
    elif ST11_SERVER_FEE:
        bk.add_gross('수수료', None, '11번가 서버이용료', ST11_SERVER_FEE)
        bk.info.append(f'서버이용료 {ST11_SERVER_FEE:,}원을 수수료에 1행으로 추가')
    elif ex is None:
        bk.info.append("서버이용료 등 원본에 없는 금액은 원본 파일의 '추가' 시트에 적어야 반영돼요")
    # 공제금액합계에 모르는 항목이 섞였는지 검사 (11번가할인은 공제 항목이 아니라 제외하고 비교)
    known_ded = fee_t - ST11_DISC_IN_FEE * st11_disc + ad_t + sum(g(r, '할인쿠폰이용료') + g(r, '[공제]스토어묶음할인')
                                   for r in rows[hi + 1:] if isdigit(r[h['NO']]))
    if known_ded != ded_t:
        bk.notes.append(f'공제금액합계({ded_t:,})와 반영한 공제 항목 합({known_ded:,})이 다름 — 새 공제 항목 확인 필요')
    if st11_disc:
        bk.info.append(f'11번가할인(11번가 부담) {st11_disc:,}원은 매출에서 제외')

    vat_files = bk.companions.get('11번가부가세', [])
    if not vat_files:
        bk.notes.append('11번가 부가세신고내역상세가 없어 부가세 시트를 비움 — 정산 파일에 부가세 시트를 붙여 넣거나 같은 기간 파일을 같이 추가하세요')
        return
    def read_vat(vws):
        vrows = list(vws.iter_rows(values_only=True))
        vi = find_header(vrows, '정산확정처리일', '과세매출금액')
        vh = header_index(vrows[vi])
        tot = {'taxable': 0, 'adj': 0, 'pay': defaultdict(int), 'free': False}
        for r in vrows[vi + 1:]:
            if not (isinstance(r[0], str) and '/' in r[0]):
                continue
            tot['taxable'] += num(r[vh['과세매출금액']])
            if '매출조정금액' in vh:
                tot['adj'] += num(r[vh['매출조정금액']])
            for pay, code in ST11_PAY_TO_TAX.items():
                if pay in vh:
                    tot['pay'][code] += num(r[vh[pay]])
            if num(r[vh['면세매출금액']]) or num(r[vh['영세매출금액']]):
                tot['free'] = True
        return tot

    # 같은 기간 부가세 파일이 여러 개면(예: 11번가 계정이 여러 개) 금액이 맞는 파일을 고름
    cands = [(fname, read_vat(vws)) for fname, vws in vat_files]
    match = [c for c in cands if c[1]['taxable'] + c[1]['adj'] == sale_t]
    fname, tot = (match or cands)[0]
    if not match:
        bk.notes.append(f'정산파일 매출({sale_t:,})과 금액이 맞는 부가세신고내역이 없음 — 같은 계정·기간 파일인지 확인')
    elif len(cands) > 1:
        bk.info.append(f'부가세신고내역 {len(cands)}개 중 금액이 맞는 파일을 사용')
    for code, amt in tot['pay'].items():
        bk.add_tax_gross(code, amt)
    if tot['free']:
        bk.notes.append('부가세신고에 면세/영세 매출이 있음 — 처리 규칙 확인 필요')
    adj_total = tot['adj']
    if adj_total:   # 11번가할인(매출조정)은 결제수단 구분이 없어 한 과세구분에서 차감
        bk.add_tax_gross(ST11_ADJ_TAX, adj_total)
        bk.info.append(f'매출조정금액 {adj_total:,}원을 부가세 시트 {ST11_ADJ_TAX}({TAX_NAME.get(ST11_ADJ_TAX, "")})에서 차감')
    bk.info.append(f'부가세: {fname} 사용')
    if -adj_total != st11_disc:
        bk.notes.append(f'정산파일 11번가할인({st11_disc:,})과 부가세신고 매출조정({adj_total:,})이 다름 — 확인 필요')


# ── 수기 추가 시트 ('추가') ─────────────────────────────────────────
# 원본에 없는 금액(예: 월 서버이용료)을 원본 파일에 '추가' 시트를 만들어 적으면 업로드 파일에 한 줄씩 넣음.
#   열: 구분(매출/수수료/광고) | 상품명 | 금액(부가세 포함) | 상품코드(선택) | 과세구분(선택, 매출일 때 14/17/31, 기본 14)
#   금액을 부가세 별도로 적으려면 '공급가' 열에 적고 금액은 비워 둠.
EXTRA_SHEET = '추가'


def extra_sheet(wb):
    return next((ws for ws in wb if ws.title.replace(' ', '').startswith(EXTRA_SHEET)), None)


def apply_extra(wb, bk):
    ws = extra_sheet(wb)
    if ws is None:
        return
    rows = list(ws.iter_rows(values_only=True))
    try:
        hi = find_header(rows, '구분', '상품명')
    except StopIteration:
        bk.notes.append("'추가' 시트에 '구분', '상품명' 제목 줄이 없어 반영하지 못함")
        return
    h = header_index(rows[hi])
    n = 0
    for r in rows[hi + 1:]:
        kind = str(r[h['구분']] or '').strip()
        name = str(r[h['상품명']] or '').strip()
        if not kind and not name:
            continue
        if kind not in ('매출', '수수료', '광고') or not name:
            bk.notes.append(f"'추가' 시트 줄을 건너뜀 (구분은 매출/수수료/광고, 상품명 필수): {kind} {name}")
            continue
        code = r[h['상품코드']] if '상품코드' in h else None
        gross = num(r[h['금액']]) if '금액' in h else 0
        sup_only = num(r[h['공급가']]) if '공급가' in h else 0
        if gross:
            sup, vat = split_vat(gross)
        elif sup_only:
            sup, vat = sup_only, vat10(sup_only)
        else:
            bk.notes.append(f"'추가' 시트 줄에 금액이 없음: {kind} {name}")
            continue
        bk.add(kind, code, name, sup, vat)
        if sup + vat < 0:
            bk.notes.append(f"'추가' 시트 {kind} '{name}'이 마이너스({sup + vat:,})로 적힘 — {kind}을(를) 줄이는 금액이 맞는지 확인 (나가는 수수료·광고비는 플러스로 적기)")
        if kind == '매출':
            tax = str(r[h['과세구분']]).strip() if '과세구분' in h and r[h['과세구분']] not in (None, '') else '14'
            if bk.tax_rows:
                bk.add_tax_row(tax, sup, vat)
            else:
                bk.add_tax(tax, sup, vat)
        n += 1
        bk.info.append(f"'추가' 시트: {kind} '{name}' {sup + vat:,}원 반영")


# ── 설정형 채널 (채널설정.xlsx) ───────────────────────────────────────
CONFIG = os.path.join(BASE, '채널설정.xlsx')
CONFIG_COLS = ['채널명', '인식문구', '시트', '상품코드열', '상품명열', '매출열', '매출부가세열',
               '수수료열', '수수료부가세열', '금액구분', '결제수단열', '기본과세구분', '포함조건열', '포함조건값']


def load_config():
    """채널설정.xlsx → [(채널명, 판별함수, 변환함수)]. 파일이 없으면 빈 목록."""
    if not os.path.exists(CONFIG):
        return [], {}
    wb = openpyxl.load_workbook(CONFIG, data_only=True)
    ws = wb['채널']
    rows = list(ws.iter_rows(values_only=True))
    h = header_index(rows[0])
    pay_map = {}
    if '결제수단' in wb.sheetnames:
        for r in list(wb['결제수단'].iter_rows(values_only=True))[1:]:
            if r[0] and r[1]:
                pay_map[str(r[0]).strip()] = str(r[1]).strip()
    out = []
    for r in rows[1:]:
        cfg = {k: (str(r[h[k]]).strip() if k in h and r[h[k]] not in (None, '') else '') for k in CONFIG_COLS}
        if (not cfg['채널명'] or cfg['채널명'].startswith('예시')
                or not cfg['인식문구'] or not cfg['상품명열'] or not cfg['매출열']):
            continue
        keys = [k.strip() for k in cfg['인식문구'].split('|') if k.strip()]
        out.append((cfg['채널명'], (lambda w, keys=keys: all(has(w, k) for k in keys)),
                    (lambda wb, bk, cfg=cfg: parse_config(wb, bk, cfg, pay_map))))
    return out, pay_map


def parse_config(wb, bk, cfg, pay_map):
    """한 시트에 '헤더 1줄 + 거래 행'이 있는 단순한 채널."""
    ws = sheet(wb, cfg['시트']) if cfg['시트'] else wb.worksheets[0]
    if ws is None:
        raise ValueError(f"시트 '{cfg['시트']}'를 찾지 못했습니다")
    rows = list(ws.iter_rows(values_only=True))
    hi = find_header(rows, cfg['상품명열'], cfg['매출열'])
    h = header_index(rows[hi])
    gross = cfg['금액구분'] != '공급가'          # 기본: 부가세 포함 금액
    def amt(r, key):
        return num(r[h[cfg[key]]]) if cfg[key] else 0
    for r in rows[hi + 1:]:
        name = r[h[cfg['상품명열']]]
        code = r[h[cfg['상품코드열']]] if cfg['상품코드열'] else None
        cells = {str(v).strip() for v in (code, name) if v not in (None, '')}
        if not cells or cells <= {'-'} or cells & {'합계', '소계', '총계', '계', 'Σ'}:
            continue   # 빈 행·합계 행
        if cfg['포함조건열'] and str(r[h[cfg['포함조건열']]] or '').strip() != cfg['포함조건값']:
            continue
        try:
            sale = amt(r, '매출열')
        except ValueError:
            continue
        if cfg['매출부가세열']:
            sup, vat = sale, amt(r, '매출부가세열')
        elif gross:
            sup, vat = split_vat(sale)
        else:
            sup, vat = sale, vat10(sale)
        bk.add('매출', code, name, sup, vat)
        tax = cfg['기본과세구분'] or DEFAULT_TAX
        if cfg['결제수단열']:
            tax = pay_map.get(str(r[h[cfg['결제수단열']]] or '').strip(), tax)
        bk.add_tax_row(tax, sup, vat)
        if cfg['수수료열']:
            fee = amt(r, '수수료열')
            if cfg['수수료부가세열']:
                bk.add('수수료', code, name, fee, amt(r, '수수료부가세열'))
            elif gross:
                bk.add_gross('수수료', code, name, fee)
            else:
                bk.add('수수료', code, name, fee, vat10(fee))


# ── 채널 인식 (파일 내용 기준) ─────────────────────────────────────────
def sheet_words(ws):
    words = {'#' + ws.title.replace(' ', '')}
    for r in ws.iter_rows(max_row=40, values_only=True):
        for v in r:
            if isinstance(v, str):
                words.add(v.strip())
    return words


def fingerprint(wb):
    words = set()
    for ws in wb:
        words |= sheet_words(ws)
    return words


def has(words, text):
    return any(text in w for w in words)


CHANNELS = [
    # (이름, 판별 함수, 변환 함수)
    ('SK스토아', lambda w: has(w, '위수탁수수료(공급가)'), parse_sk),
    ('NS홈쇼핑', lambda w: has(w, '#쇼핑북광고비') or (has(w, '제휴사') and has(w, '특약매입')), parse_ns),
    ('넛지헬스케어캐시딜', lambda w: has(w, '판매수수료(C=A*B'), parse_nudge),
    ('오늘의집', lambda w: has(w, '주문옵션번호') and has(w, '부가세신고'), parse_ohouse),
    ('공영홈쇼핑', lambda w: has(w, '매출차감적립금'), parse_gongyoung),
    ('우리홈쇼핑', lambda w: has(w, '매입수수료') and has(w, '판매금액'), parse_woori),
    ('CJ온스타일', lambda w: has(w, 'CJ온스타일매출'), parse_cj),
    ('롯데ON', lambda w: has(w, '정산대상판매금액') and has(w, '판매자상품명'), parse_lotteon),
    ('11번가', lambda w: has(w, '판매금액합계') and has(w, '공제금액합계') and has(w, '할인쿠폰이용료'), parse_11st),
]

# 단독으로는 변환하지 않고, 같은 기간의 본 파일에 붙여 쓰는 보조 파일
# (키, 판별 함수, 설명, 붙는 채널)
COMPANIONS = [
    ('11번가부가세', lambda w: has(w, '부가세신고내역상세') and has(w, '[정산]스토어묶음할인'), '11번가 부가세신고내역', '11번가'),
]


DATA_START = 4   # 양식 1~3행(한글 제목 / 영문 필드명 / 타입 설명)은 그대로 두고 4행부터 데이터


def write(out, bk):
    """아마란스 업로드 양식 그대로: 1~3행 유지, 샘플 행 지우고 4행부터 데이터. 합계 행은 넣지 않음(데이터로 읽힐 수 있어서)."""
    shutil.copyfile(TEMPLATE, out)
    wb = openpyxl.load_workbook(out)
    for title in ('매출', '광고', '수수료'):
        ws = wb[title]
        if ws.max_row >= DATA_START:
            ws.delete_rows(DATA_START, ws.max_row)
        for i, (code, name, sup, vat) in enumerate(bk.rows[title], start=DATA_START):
            ws.cell(i, 1, code[:30] or None).number_format = '@'
            ws.cell(i, 2, name[:100])
            ws.cell(i, 3, sup)
            ws.cell(i, 4, vat)
    ws = wb['부가세']
    if ws.max_row >= DATA_START:
        ws.delete_rows(DATA_START, ws.max_row)
    if bk.tax_rows:                      # 건별 (홈쇼핑·넛지·설정형 채널)
        tax_lines = bk.tax_rows
    else:                                # 결제수단별 합계 (오늘의집·롯데ON·11번가)
        tax_lines = [[c, *bk.tax[c][:2]] for c in sorted(bk.tax, key=lambda c: {'17': 0, '31': 1}.get(c, 2))]
    i = DATA_START
    for code, s_, v in tax_lines:
        if s_ or v:
            ws.cell(i, 1, code).number_format = '@'
            ws.cell(i, 2, s_); ws.cell(i, 3, v); i += 1
    wb.save(out)


def new_result(path):
    return {'file': os.path.basename(path), 'status': '', 'channel': '', 'out': '',
            'sums': {}, 'tax': {}, 'notes': [], 'info': [], 'message': ''}


def scan(path):
    """파일을 열어 내용 지문·기간을 읽음."""
    wb = load_book(path)
    words = fingerprint(wb)
    return wb, words, period_of(words)


def run_one(path, out_dir, extra, preloaded=None, companions=None):
    """파일 1개 변환. 결과를 dict로 돌려줌 (화면/창 표시용).
    status: '완료' | '건너뜀' | '오류' | '보조'"""
    res = new_result(path)
    try:
        wb, words, _ = preloaded or scan(path)
        match = [c for c in CHANNELS if c[1](words)] or [c for c in extra if c[1](words)]
        if len(match) != 1:
            res['status'] = '건너뜀'
            res['message'] = (f'여러 채널과 겹침: {", ".join(c[0] for c in match)}' if match
                              else '등록되지 않은 채널 — 채널설정.xlsx에 추가하거나 전산 담당자에게 원본 파일 전달')
            return res
        ch, _, parse = match[0]
        bk = Book()
        bk.companions = {k: list(v) for k, v in (companions or {}).items()}
        for ws in wb:   # 한 파일 안에 붙여 넣은 보조 시트가 있으면 그걸 우선 사용
            sw = sheet_words(ws)
            for key, det, label, target in COMPANIONS:
                if target == ch and det(sw):
                    bk.companions[key] = [(f"같은 파일의 '{ws.title}' 시트", ws)]
        parse(wb, bk)
        apply_extra(wb, bk)
        bk.finalize()
        base = os.path.splitext(os.path.basename(path))[0]
        out = os.path.join(out_dir, f'업로드_{ch}_{base}.xlsx')
        write(out, bk)
        res.update(status='완료', channel=ch, out=out, notes=bk.notes, info=bk.info)
        for t in ('매출', '광고', '수수료'):
            rs = bk.rows[t]
            res['sums'][t] = (len(rs), sum(r[2] for r in rs), sum(r[3] for r in rs))
        res['tax'] = {c: tuple(v[:2]) for c, v in sorted(bk.tax.items()) if v[0] or v[1]}
    except Exception as e:
        res['status'] = '오류'
        res['message'] = f'변환 중 문제가 생겼습니다 (원본 양식이 바뀌었을 수 있음): {e}'
        res['trace'] = traceback.format_exc()
    return res


def run_batch(paths, out_dir, extra, on_result=None, on_progress=None):
    """여러 파일을 한 번에. 보조 파일(예: 11번가 부가세신고내역)은 같은 기간의 본 파일에 자동으로 붙임."""
    scanned, results = {}, {}
    for i, p in enumerate(paths, 1):
        if on_progress:
            on_progress(i, len(paths), p, '읽는 중')
        try:
            scanned[p] = scan(p)
        except Exception as e:
            r = new_result(p)
            r.update(status='오류', message=f'파일을 열 수 없습니다: {e}', trace=traceback.format_exc())
            results[p] = r
    comp = {}   # (키, 기간) → [(파일명, wb, 경로)]
    comp_info = {}
    for p, (wb, words, period) in scanned.items():
        if any(c[1](words) for c in CHANNELS):
            continue   # 본 파일 (보조 시트가 안에 붙어 있을 수 있음)
        for key, det, label, target in COMPANIONS:
            ws = next((w for w in wb if det(sheet_words(w))), None)
            if ws is not None:
                comp.setdefault((key, period), []).append((os.path.basename(p), ws, p))
                comp_info[p] = (key, period, label, target)
                break
    used = set()
    for i, p in enumerate(paths, 1):
        if p in results or p in comp_info:
            continue
        if on_progress:
            on_progress(i, len(paths), p, '변환 중')
        wb, words, period = scanned[p]
        mine = {}
        for (key, per), lst in comp.items():
            if per == period or (not per or not period):
                mine.setdefault(key, []).extend((n, w) for n, w, _ in lst)
                used.update(cp for _, _, cp in lst)
        results[p] = run_one(p, out_dir, extra, preloaded=(wb, words, period), companions=mine)
        if on_result:
            on_result(p, results[p])
    for p, (key, period, label, target) in comp_info.items():
        r = new_result(p)
        r['channel'] = target
        if p in used:
            r.update(status='보조', message=f'{label} — {target} 정산 파일과 함께 처리됨')
        else:
            r.update(status='건너뜀', message=f'{label} — 같은 기간({period})의 {target} 정산 파일을 같이 추가하세요')
        results[p] = r
        if on_result:
            on_result(p, r)
    for p in paths:   # 열지 못한 파일
        if on_result and p in results and p not in scanned:
            on_result(p, results[p])
    return results


def print_result(res):
    """콘솔 출력용."""
    if res['status'] != '완료':
        print(f"[{res['status']}] {res['file']} — {res['message']}")
        if res.get('trace'):
            print(res['trace'])
        print()
        return
    print(f"[완료] {res['file']} → {os.path.basename(res['out'])}")
    for t, (n, s, v) in res['sums'].items():
        if n:
            print(f'    {t:3s} {n:4d}행  공급가 {s:>15,}  부가세 {v:>13,}  합계 {s + v:>15,}')
    for code, (s, v) in res['tax'].items():
        print(f'    부가세 {code}      공급가 {s:>15,}  부가세 {v:>13,}')
    for n in res['notes']:
        print('    ※', n)
    for n in res['info']:
        print('    ·', n)
    print()


def main(argv):
    open_folder = '--open' in argv
    args = [a for a in argv[1:] if not a.startswith('--')]
    files = args or sorted(glob.glob(os.path.join(IN_DIR, '*.xls*')))
    out_dir = OUT_DIR
    files = [f for f in files if not os.path.basename(f).startswith('~$')]
    if not files:
        print(f"'원본' 폴더에 엑셀 파일이 없습니다: {IN_DIR}")
        return 1
    os.makedirs(out_dir, exist_ok=True)
    try:
        extra, _ = load_config()
    except Exception as e:
        print(f'[주의] 채널설정.xlsx를 읽지 못했습니다 — {e}')
        extra = []
    results = run_batch(files, out_dir, extra, on_result=lambda p, r: print_result(r))
    ok = sum(r['status'] == '완료' for r in results.values())
    failed = [r['file'] for r in results.values() if r['status'] not in ('완료', '보조')]
    print(f'총 {len(files)}개 중 {ok}개 변환. 결과 폴더: {out_dir}')
    if failed:
        print(f'변환 못 한 파일 {len(failed)}개:')
        for f in failed:
            print('   -', f)
    print('※ 업로드 전, 위 합계를 채널 정산서 합계와 한 번 대조하세요.')
    if open_folder and ok and hasattr(os, 'startfile'):
        os.startfile(out_dir)   # 윈도우에서 결과 폴더 열기
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv))
