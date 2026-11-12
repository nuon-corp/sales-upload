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

# 규칙을 고치면 이 날짜·시간도 바꿔 주세요 — 웹 화면 위쪽에 표시돼서 어떤 버전으로 변환했는지 알 수 있어요.
# 바뀐 내용은 깃헙 커밋 기록(convert_sales.py → History)에 남아요.
RULES_VERSION = '2026-11-12 13:28'

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
    if path.lower().endswith('.csv'):
        import csv, io
        raw = open(path, 'rb').read()
        for enc in ('utf-8-sig', 'cp949', 'euc-kr'):
            try:
                text = raw.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        else:
            raise RuntimeError('CSV 글자 인코딩을 읽지 못했습니다')
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = os.path.splitext(os.path.basename(path))[0][:31] or 'csv'
        for row in csv.reader(io.StringIO(text)):
            ws.append([None if v == '' else v for v in row] or [None])
        return wb
    try:
        return openpyxl.load_workbook(path, data_only=True)
    except ValueError:
        # 일부 사이트 다운로드는 셀 주소에 행 번호가 빠져 있음(r="A") → 고쳐서 다시 읽기
        return openpyxl.load_workbook(_fix_cell_refs(path), data_only=True)


def _fix_cell_refs(path):
    import io, zipfile
    src = zipfile.ZipFile(path)
    out = io.BytesIO()
    dst = zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED)
    row_re = re.compile(r'<row\b([^>]*)>(.*?)</row>', re.S)
    for info in src.infolist():
        data = src.read(info.filename)
        if info.filename.startswith('xl/worksheets/') and info.filename.endswith('.xml'):
            text = data.decode('utf-8')
            n = [0]
            def fix_row(m):
                rn = re.search(r'\br="(\d+)"', m.group(1))
                n[0] = int(rn.group(1)) if rn else n[0] + 1
                body = re.sub(r'(<c\b[^>]*?\br=")([A-Z]+)(")', lambda c: f'{c.group(1)}{c.group(2)}{n[0]}{c.group(3)}', m.group(2))
                return f'<row{m.group(1)}>{body}</row>'
            data = row_re.sub(fix_row, text).encode('utf-8')
        dst.writestr(info, data)
    dst.close()
    out.seek(0)
    return out


def period_of(words):
    """제목에 적힌 조회기간 '(2026/10/01~2026/10/31)' → 짝 맞추기 키."""
    for w in words:
        m = re.search(r'(\d{4}[./-]\d{2}[./-]\d{2})\s*~\s*(\d{4}[./-]\d{2}[./-]\d{2})', w)
        if m:
            return m.group(1).replace('.', '/').replace('-', '/') + '~' + m.group(2).replace('.', '/').replace('-', '/')
    return ''


def first_sheet(wb):
    """첫 번째 데이터 시트 ('추가' 시트는 건너뜀)."""
    return next((ws for ws in wb if not ws.title.replace(' ', '').startswith('추가')), wb.worksheets[0])


def sheet(wb, prefix, *need):
    """시트 찾기: 이름이 prefix로 시작하는 시트 → 없으면 열 제목(need)이 모두 들어 있는 시트.
    그래서 원본 시트 이름이 바뀌어도 열 제목으로 찾아감."""
    ws = next((s for s in wb if s.title.replace(' ', '').startswith(prefix.replace(' ', ''))), None)
    if ws is not None or not need:
        return ws
    return next((s for s in wb if not s.title.replace(' ', '').startswith('추가')
                 and all(has(sheet_words(s), k) for k in need)), None)


class Book:
    """업로드 파일 한 개 분량 (시트별 행 목록)."""
    def __init__(self):
        self.rows = {'매출': [], '광고': [], '수수료': []}
        self.tax = defaultdict(lambda: [0, 0, 0])   # 과세구분 → [공급가, 부가세, 부가세포함 누적] (화면 합계용)
        self.tax_rows = []                           # 부가세 시트에 건별로 쓸 행 [과세구분, 공급가, 부가세]
        self.notes = []   # 확인이 필요한 사항 (가정, 미반영 시트)
        self.info = []    # 참고 안내 (확인된 규칙대로 제외한 것 등)
        self.companions = {}   # 함께 올린 보조 파일 (예: 11번가 부가세신고내역) {키: [(파일명, 시트)]}
        self.used = set()      # 실제로 사용한 보조 파일 이름
        self.label = ''        # 채널 안 구분(브랜드 등) — 결과 파일명·화면에 붙음

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
    ws = sheet(wb, '수수료', '위수탁수수료(공급가)', '고객결제금액')
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


def parse_ns(wb, bk):
    """NS홈쇼핑(특약매입): 매출 = 매출금액-반품금액, 부가세 = 단가부가세×수량, 광고 = 쇼핑북광고비."""
    ws = sheet(wb, '품목별', '상품코드', '매출금액')
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
    ad = sheet(wb, '쇼핑북광고비', '광고비', '상품코드')
    if ad:
        rows = list(ad.iter_rows(values_only=True))
        hi = find_header(rows, '상품코드', '광고비')
        h = header_index(rows[hi])
        for r in rows[hi + 1:]:
            if r[h['상품코드']] and num(r[h['광고비']]):
                amt = num(r[h['광고비']])
                bk.add('광고', r[h['상품코드']], r[h['상품명']], amt, vat10(amt))
    ship = sheet(wb, '고객부담배송비', '배송비 합계')
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
    bk.info.append('세금계산서 매출이라 부가세 시트는 비움')


def parse_nudge(wb, bk):
    """넛지헬스케어 캐시딜: 매출 = 상품합계(A), 수수료 = 판매수수료(VAT 포함 가정)."""
    ws = first_sheet(wb)
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


def parse_ohouse(wb, bk):
    """오늘의집: 매출·수수료 = 품목별 내역 행 단위, 할인 = 부가세신고 화면의 매출할인, 부가세 = 결제수단별 합계."""
    vat_ws = sheet(wb, 'VAT', '주문옵션', '결제금액', '결제수단')
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

    ws = sheet(wb, '품목별', '주문옵션번호', '상품명')
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
    ad = sheet(wb, '광고비', '광고유형', '유상 광고비')
    if ad and any((isinstance(v, (int, float)) and v) or (isdigit(v) and int(v)) for r in ad.iter_rows(values_only=True) for v in r[3:4]):
        bk.info.append('광고비 시트는 경영지원팀 방식대로 업로드에서 제외')


def parse_gongyoung(wb, bk):
    """공영홈쇼핑: 상품별매출내역(차수별 행) — 매출순금액/VAT, 수수료금액/VAT 그대로. 광고 = 카드무이자 협력사부담."""
    ws = sheet(wb, '품목별', '상품ID', '매출순금액')
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
    mu = sheet(wb, '무이자', '상품번호', '협력사부담금액')
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
    ws = sheet(wb, '품목별', '판매금액', '매입수수료')
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


def parse_cj(wb, bk):
    """CJ온스타일: 수수료 매출 일별내역 — 소비자공급가/부가세, CJ수수료/부가세 그대로. 광고 = 공제내역의 광고비."""
    ws = sheet(wb, '건별매출', '판매상품명', 'CJ온스타일매출(수수료매출)')
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
    gj = sheet(wb, '공제내역(', '공제명', '공제금액')
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
    ws = sheet(wb, '품목별', '판매자상품명', '정산대상판매금액')
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
    vs = sheet(wb, 'VAT', '총매출금액', '신용카드', '현금영수증')
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
    bk.used.add(fname)
    if -adj_total != st11_disc:
        bk.notes.append(f'정산파일 11번가할인({st11_disc:,})과 부가세신고 매출조정({adj_total:,})이 다름 — 확인 필요')


# ── 카카오 (톡스토어·선물하기) ───────────────────────────────────────
# 카카오 판매자번호 → 브랜드 (화면 채널명에 표시). 새 판매자 계정이 생기면 추가
KAKAO_SELLERS = {'10094': '살림백서', '205510': '아우페', '207416': '헤디안', '230621': '프로클'}
KAKAO_PAY_TO_TAX = {'카카오페이 카드결제': '17', '신용카드': '17', '현금(카카오머니)': '31', '무통장입금': '31',
                    '휴대폰 결제': '14', '포인트': '31'}   # 톡스토어 부가세 파일 분류로 확인 (포인트는 대부분 '현금'으로 잡힘)
KAKAO_VAT_COLS = {'카드': '17', '현금': '31', '휴대폰': '14', '기타': '14'}


def parse_kakao(wb, bk):
    """카카오 정산(판매자번호별) + 부가세 파일(같은 판매자번호).
    톡스토어: 매출 = 정산기준금액 + 선불배송비 + 반품배송비 − 수수료차감금액(부가세 파일) → 부가세 파일 과세매출과 일치
              수수료 = 수수료합계 + 큐레이터 서비스수수료 − 수수료차감금액
    선물하기: 부가세 파일에 없음 → 매출 = 정산기준금액 + 배송비, 수수료 = 수수료합계 (확인 필요)"""
    ws = next(w for w in wb if has(sheet_words(w), '판매정산금액'))
    rows = list(ws.iter_rows(values_only=True))
    hi = find_header(rows, '판매정산금액', '정산기준금액', '판매자번호')
    h = header_index(rows[hi])
    g = lambda r, k: num(r[h[k]]) if k in h else 0
    data = [r for r in rows[hi + 1:] if r and r[h['주문번호']] not in (None, '') or (r and r[h['정산구분']] == '배송비')]
    sellers = {str(r[h['판매자번호']]).split('.')[0] for r in data if r[h['판매자번호']] not in (None, '')}
    bk.info.append(f"판매자번호 {', '.join(sorted(sellers))}")
    bk.label = '·'.join(KAKAO_SELLERS.get(x, x) for x in sorted(sellers))

    # 같은 판매자번호의 부가세 파일 찾기
    vat = None
    for fname, vws in bk.companions.get('카카오부가세', []):
        vrows = list(vws.iter_rows(values_only=True))
        try:
            vi = find_header(vrows, '수수료차감금액', '과세매출금액', '판매자번호')
        except StopIteration:
            continue
        vh = header_index(vrows[vi])
        vdata = [r for r in vrows[vi + 1:] if r and r[vh['판매자번호']] not in (None, '')]
        if {str(r[vh['판매자번호']]).split('.')[0] for r in vdata} & sellers:
            vat = (fname, vh, vdata)
            break
    deduct = defaultdict(int)
    if vat:
        fname, vh, vdata = vat
        for r in vdata:
            if r[vh['주문번호']] not in (None, ''):
                deduct[str(r[vh['주문번호']]).split('.')[0]] += num(r[vh['수수료차감금액']])
        bk.used.add(fname)
        bk.info.append(f'톡스토어 부가세: {fname} 사용')
    else:
        bk.notes.append('톡스토어 부가세 자료 없음 — 결제수단으로 과세구분을 추정함. 부가세 자료를 같은 파일에 시트로 넣어 주세요')

    tok_sale = gift_disc = gift_sale = gift_n = 0
    done_orders = set()
    for r in data:
        is_ship = r[h['정산구분']] == '배송비'
        code = None if is_ship else r[h['상품번호']]
        name = '배송비' if is_ship or not r[h['상품명']] else r[h['상품명']]
        order = str(r[h['주문번호']] or '').split('.')[0]
        sale = g(r, '정산기준금액') + g(r, '선불배송비') + g(r, '반품배송비')
        fee = g(r, '수수료합계') + g(r, '큐레이터 서비스수수료')
        gift = r[h['채널']] != '카카오톡 스토어'
        if not gift and order and order not in done_orders:
            sale -= deduct.get(order, 0)
            fee -= deduct.get(order, 0)
            done_orders.add(order)
        bk.add_gross('매출', code, name, sale)
        bk.add_gross('수수료', code, name, fee)
        if gift or not vat:
            bk.add_tax_gross(KAKAO_PAY_TO_TAX.get(str(r[h['결제수단']]).strip(), DEFAULT_TAX), sale)
        if gift:
            gift_disc += g(r, '카카오할인금액합계')
            gift_sale += sale; gift_n += 1
        else:
            tok_sale += sale
    if vat:
        fname, vh, vdata = vat
        for col, tcode in KAKAO_VAT_COLS.items():
            if col in vh:
                bk.add_tax_gross(tcode, sum(num(r[vh[col]]) for r in vdata))
        vtot = sum(num(r[vh['과세매출금액']]) for r in vdata)
        if any(num(r[vh['면세매출금액']]) for r in vdata):
            bk.notes.append('카카오 부가세 파일에 면세매출이 있음 — 처리 규칙 확인 필요')
        if vtot != tok_sale:
            bk.notes.append(f'톡스토어 매출({tok_sale:,})과 부가세 파일 과세매출({vtot:,})이 다름 — 같은 판매자·기간 파일인지 확인')
    if gift_n:
        msg = (f'선물하기 부가세 자료 없음 ({gift_n}건, {gift_sale:,}원) — 결제수단으로 과세구분 추정 '
               f'(카드 17 / 현금·무통장·포인트 31 / 휴대폰 14)')
        if gift_disc:
            msg += f', 카카오할인 {gift_disc:,}원은 빼지 않음'
        bk.notes.append(msg + '. 선물하기 부가세 자료를 같은 파일에 시트로 넣어 주세요')


# ── 토스쇼핑 ──────────────────────────────────────────────────────────
TOSS_VAT_COLS = {'신용카드 과세 공급가액': '17', '현금영수증 발행 과세 공급가액': '31', '기타 매출 과세 공급가액': '14'}
TOSS_ADJ_TAX = '14'   # 매출 에누리(토스 부담 할인) 차감 과세구분 — 토스 설명상 '기타 매출'에서 차감


def parse_toss(wb, bk):
    """토스쇼핑 정산 CSV + 부가세 '상세 내역'.
    매출 = 수수료 기준금액 − 토스 부담 할인 금액 (토스 설명: 매출 에누리는 수수료 내에서 공제)
    수수료 = 결제 수수료 + 상품 판매 수수료(각 부가세 포함) − 토스 부담 할인 금액
    부가세 = 상세 내역의 신용카드 17 / 현금영수증 31 / 기타 14 공급가, 매출 에누리는 14에서 차감"""
    ws = next(w for w in wb if has(sheet_words(w), '수수료 기준금액'))
    rows = list(ws.iter_rows(values_only=True))
    hi = find_header(rows, '수수료 기준금액', '상품 주문번호')
    h = header_index(rows[hi])
    g = lambda r, k: num(r[h[k]]) if k in h and len(r) > h[k] else 0
    base_t = disc_t = 0
    for r in rows[hi + 1:]:
        if not r or r[0] in (None, ''):
            continue
        base = g(r, '수수료 기준금액')
        fsup = -(g(r, '결제 수수료') + g(r, '상품 판매 수수료'))
        fvat = -(g(r, '결제 수수료 부가세') + g(r, '상품 판매 수수료 부가세'))
        disc = g(r, '토스 부담 할인 금액') + g(r, '배송비 토스 부담 할인 금액')
        if not (base or fsup or fvat or disc):
            continue
        pid = str(r[h['상품ID']] or '').split('.')[0]
        code = None if pid in ('', '0') else pid
        name = r[h['상품명']] or '배송비'
        bk.add_gross('매출', code, name, base - disc)
        ds, dv = split_vat(disc)
        bk.add('수수료', code, name, fsup - ds, fvat - dv)
        base_t += base; disc_t += disc
    if disc_t:
        bk.info.append(f'토스 부담 할인 {disc_t:,}원은 매출·수수료에서 차감')

    cands = []
    for fname, vws in bk.companions.get('토스부가세', []):
        vrows = list(vws.iter_rows(values_only=True))
        vi = find_header(vrows, '주문상품번호', '매출 에누리')
        vh = header_index(vrows[vi])
        vdata = [r for r in vrows[vi + 1:] if r and str(r[0] or '').startswith('20')]
        sup = {k: sum(num(r[vh[k]]) for r in vdata) for k in TOSS_VAT_COLS if k in vh}
        vat_t = sum(num(r[vh['[참고] 부가가치세(VAT)']]) for r in vdata)
        enuri = sum(num(r[vh['매출 에누리']]) for r in vdata)
        free = sum(num(r[vh['면세 공급가액']]) for r in vdata) if '면세 공급가액' in vh else 0
        cands.append((fname, sup, vat_t, enuri, free))
    if not cands:
        bk.notes.append('토스 부가세 상세 내역 파일이 없어 부가세 시트를 비움 — 같은 달 파일을 같이 올려 주세요')
        return
    match = [c for c in cands if sum(c[1].values()) + c[2] + c[4] == base_t]
    fname, sup, vat_t, enuri, free = (match or cands)[0]
    if not match:
        bk.notes.append(f'정산 수수료기준금액({base_t:,})과 금액이 맞는 토스 부가세 파일이 없음 — 같은 달 파일인지 확인')
    if free:
        bk.notes.append('토스 부가세 파일에 면세 매출이 있음 — 처리 규칙 확인 필요')
    # 과세구분별 공급가 + 부가세(10%, 합계가 VAT 총액과 같도록 가장 큰 구분에서 보정)
    parts = {TOSS_VAT_COLS[k]: [v, vat10(v)] for k, v in sup.items() if v}
    if parts:
        big = max(parts.values(), key=lambda x: x[0])
        big[1] += vat_t - sum(x[1] for x in parts.values())
    for code, (sp, vt) in parts.items():
        bk.add_tax(code, sp, vt)
    if enuri:
        bk.add_tax_gross(TOSS_ADJ_TAX, enuri)
        bk.info.append(f'매출 에누리 {enuri:,}원을 부가세 시트 {TOSS_ADJ_TAX}({TAX_NAME.get(TOSS_ADJ_TAX, "")})에서 차감')
    if -enuri != disc_t:
        bk.notes.append(f'정산 토스 부담 할인({disc_t:,})과 부가세 파일 매출 에누리({enuri:,})가 다름 — 확인 필요')
    bk.used.add(fname)
    bk.info.append(f'부가세: {fname} 사용')


# ── 포인트·PG 결제 입점몰 (요약/목록 시트) ─────────────────────────────
POINTMALL_NAME = '(주) 엘지씨엔에스'   # 복지몰(포인트·PG 결제) 입점 정산
POINTMALL_TAX = {'포인트총액': '14', 'PG총액': '17'}


def parse_pointmall(wb, bk):
    """입점사 정산 '목록' 시트: 매출 = 판매가총액 + 배송비총액 − 할인총액(= 포인트총액 + PG총액)
    수수료 = 판매사수수료 − (할인총액 − 쿠폰(입점사부담))  → 매출 − 수수료 = 대금지급액"""
    ws = sheet(wb, '목록', '주문상품요약', '판매사수수료(총액)') or first_sheet(wb)
    rows = list(ws.iter_rows(values_only=True))
    hi = find_header(rows, '판매가총액', '판매사수수료(총액)')
    h = header_index(rows[hi])
    g = lambda r, k: num(r[h[k]]) if k in h else 0
    plat_disc = 0
    for r in rows[hi + 1:]:
        if not r or r[h['입점사']] in (None, ''):
            continue
        name = r[h['주문상품요약']] or '상품'
        sale = g(r, '판매가총액') + g(r, '배송비총액') - g(r, '할인총액')
        pdisc = g(r, '할인총액') - g(r, '쿠폰(입점사부담)')
        bk.add_gross('매출', None, name, sale)
        bk.add_gross('수수료', None, name, g(r, '판매사수수료(총액)') - pdisc)
        for col, code in POINTMALL_TAX.items():
            if g(r, col):
                bk.add_tax_gross(code, g(r, col))
        if abs(sale - g(r, '포인트총액') - g(r, 'PG총액')) > 0:
            bk.notes.append(f"결제금액(포인트+PG)과 매출이 다른 행: {str(name)[:20]}")
        plat_disc += pdisc
    if plat_disc:
        bk.info.append(f'운용사 부담 할인 {plat_disc:,}원은 매출·수수료에서 차감')


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
    ws = sheet(wb, cfg['시트']) if cfg['시트'] else first_sheet(wb)
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
    for r in ws.iter_rows(max_row=60, values_only=True):
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


# ── 2026-11-11 추가 매출처 ──────────────────────────────────────────────
# 알리익스프레스 정산 기간 ID 안의 판매자 번호 → 브랜드 (결과 파일·화면 표시용)
ALI_ACCOUNTS = {'4866425162': '살림백서', '6114755911': '아우페', '6472674923': '프로클'}
ALI_TAX = '14'          # 결제수단 정보가 없어 건별로 넣음 (경영지원팀 확인 필요)


def _won(v):
    """'₩ 46,691' / '-1,082' → 정수."""
    t = str(v if v is not None else '').replace('₩', '').replace(',', '').strip()
    try:
        return int(float(t)) if t not in ('', '-') else 0
    except ValueError:
        return 0


def parse_ali(wb, bk):
    """알리익스프레스: 거래명세서 행 단위. 매출 = '판매'(상품 결제액·환불), 수수료 = 'AliExpress 비용'
    (세금 포함 금액은 ÷1.1, 제휴 수수료는 공급가·부가세 행을 주문별로 묶음), 광고 = 광고 시트 비용 합계."""
    ws = sheet(wb, 'Aliexpress', '비용 항목명', '수수료 항목명', '총액(KRW)')
    rows = list(ws.iter_rows(values_only=True))
    hi = find_header(rows, '비용 항목명', '수수료 항목명', '총액(KRW)')
    h = header_index(rows[hi])
    accounts = set()
    aff = defaultdict(lambda: [0, 0, ''])   # (주문 ID, 제품 ID) → [제휴 수수료, 제휴 수수료 부가세, 상품명]
    for r in rows[hi + 1:]:
        kind, item = r[h['비용 항목명']], str(r[h['수수료 항목명']] or '')
        if not kind:
            continue
        m = re.search(r'SETTLE_(\d+)_', str(r[h['정산 기간 ID']] or ''))
        if m:
            accounts.add(m.group(1))
        code, name, amt = r[h['제품 ID']], r[h['아이템 이름']], _won(r[h['총액(KRW)']])
        if kind == '판매':
            bk.add_gross('매출', code, name, amt)
        elif item.startswith('제휴 수수료'):
            k = (r[h['주문 ID']], code)
            aff[k][1 if '부가세' in item else 0] += -amt
            aff[k][2] = name
        else:
            bk.add_gross('수수료', code, name, -amt)
    for (_, code), (sup, vat, name) in aff.items():
        bk.add('수수료', code, name, sup, vat)
    bk.label = '·'.join(ALI_ACCOUNTS.get(a, a) for a in sorted(accounts))
    bk.tax_from_sales(ALI_TAX)
    ad = sheet(wb, '광고', '비용', '클릭 수')
    if ad:
        arows = list(ad.iter_rows(values_only=True))
        ah = header_index(arows[find_header(arows, '비용', '클릭 수')])
        total = sum(_won(r[ah['비용']]) for r in arows if re.match(r'\d{4}-\d{2}-\d{2}', str(r[ah['날짜']] or '')))
        if total:
            bk.add_gross('광고', None, '알리익스프레스 광고비', total)
    bk.info.append(f'부가세 시트: 결제수단 정보가 없어 {TAX_NAME[ALI_TAX]}({ALI_TAX})로 넣음')


def parse_welstory(wb, bk):
    """삼성웰스토리(웰스토리몰, 매입): 매출 = 최종매입공급금액/VAT, 수수료 없음. 세금계산서 매출."""
    ws = sheet(wb, '§', '최종매입공급금액', '최종매입VAT')
    rows = list(ws.iter_rows(values_only=True))
    hi = find_header(rows, '최종매입공급금액', '최종매입VAT')
    h = header_index(rows[hi])
    for r in rows[hi + 1:]:
        if not r[h['주문번호']]:
            continue
        taxable = r[h['과면세']] != '면세'
        bk.add('매출', r[h['업체상품번호']], r[h['상품명']], num(r[h['최종매입공급금액']]),
               num(r[h['최종매입VAT']]) if taxable else 0)
    bk.info.append('세금계산서 매출이라 부가세 시트는 비움')


def parse_welfare(wb, bk):
    """복지샵 확정상세내역(매입): 매출 = 실정산금액(부가세 포함, ÷1.1), 수수료 없음. 세금계산서 매출."""
    ws = sheet(wb, '확정상세', '실정산금액')
    rows = list(ws.iter_rows(values_only=True))
    hi = next(i for i, r in enumerate(rows) if r and '상품명' in r and '주문번호' in r)
    top = [str(v or '').replace('\n', '') for v in rows[hi]]
    sub = [str(v or '') for v in rows[hi + 1]]
    c_name, c_code, c_tax = top.index('상품명'), top.index('상품코드'), next(i for i, v in enumerate(top) if v.startswith('면과세'))
    c_amt = sub.index('실정산금액')
    for r in rows[hi + 2:]:
        if r[0] == '합계' or not r[c_name]:
            continue
        bk.add_gross('매출', r[c_code], r[c_name], num(r[c_amt]), r[c_tax] != '면세')
    bk.info.append('세금계산서 매출이라 부가세 시트는 비움')


CONSIGN_PAY_TO_TAX = {'카드': '17', '현금': '31', '포인트': '14', '기타': '14'}


def parse_consign(wb, bk):
    """위탁몰(매출신고액 양식): 매출 = 총 결제금액 공급가/세액, 수수료 = 위탁수수료 공급가액/세액,
    부가세 = 카드 17 / 현금 31 / 포인트·기타 14."""
    ws = sheet(wb, '§', '총 결제금액(매출신고액)', '위탁수수료 공급가액')
    rows = list(ws.iter_rows(values_only=True))
    hi = find_header(rows, '총 결제금액(매출신고액)', '위탁수수료 공급가액')
    h = header_index(rows[hi])
    names = {}   # 상세 시트에 상품명이 비어 있으면 다른 시트의 '품목 명칭'으로 채움
    for other in wb.worksheets:
        orows = list(other.iter_rows(values_only=True))
        try:
            oh = header_index(orows[find_header(orows, '주문번호', '품목 명칭')])
        except StopIteration:
            continue
        for r in orows:
            if isdigit(r[oh['주문번호']]):
                names[str(r[oh['주문번호']])] = r[oh['품목 명칭']]
    for r in rows[hi + 1:]:
        if not isdigit(r[h['No']]):
            continue
        name = r[h['상품명']] or names.get(str(r[h['주문번호']]), '')
        bk.add('매출', None, name, num(r[h['결제금액 공급가']]), num(r[h['결제금액 세액']]))
        bk.add('수수료', None, name, num(r[h['위탁수수료 공급가액']]), num(r[h['위탁수수료 세액']]))
        pays = [(code, num(r[h[pay]])) for pay, code in CONSIGN_PAY_TO_TAX.items() if pay in h and num(r[h[pay]])]
        if len(pays) == 1:   # 결제수단 하나면 매출 행 공급가·세액 그대로 (매출 시트와 원 단위 일치)
            bk.add_tax_row(pays[0][0], num(r[h['결제금액 공급가']]), num(r[h['결제금액 세액']]))
        else:
            for code, amt in pays:
                bk.add_tax_row(code, *split_vat(amt))


ALWAYS_PAY_TO_TAX = {'카드': '17', '현금': '31'}


def parse_always(wb, bk):
    """올웨이즈(레브잇) 세금계산서 상세내역: 매출 = 수수료대상 거래금액 − 매출에누리(올웨이즈 부담), 수수료 = 세금계산서 발급금액
    (= 수수료 + 특별수수료 − 매출에누리(올웨이즈 부담)). 결제수단 없는 행(정산제외 - 서버이용료·올팜 사진후기·기타)은 수수료로 넣음.
    부가세 = 카드 17 / 현금 31."""
    ws = sheet(wb, '세금계산서', '수수료대상 거래금액', '세금계산서 발급금액')
    rows = list(ws.iter_rows(values_only=True))
    hi = find_header(rows, '수수료대상 거래금액', '세금계산서 발급금액')
    h = header_index(rows[hi])
    etc = defaultdict(lambda: [0, 0])   # 정산제외 항목명 → [건수, 금액]
    for r in rows[hi + 1:]:
        pay, fee = r[h['결제 수단']], num(r[h['세금계산서 발급금액']])
        if pay:
            name = r[h['상품명']]
            amt = num(r[h['수수료대상 거래금액']]) - num(r[h['매출에누리(올웨이즈 부담)']])   # 경영지원팀 확인(11/12)
            bk.add_gross('매출', None, name, amt)
            bk.add_gross('수수료', None, name, fee)
            bk.add_tax_gross(ALWAYS_PAY_TO_TAX.get(str(pay).strip(), '14'), amt)
        elif fee:
            # 상품 없는 청구(모두 수수료): 주문아이디 칸에 '정산제외 - 서버이용료 / 올팜 사진후기 / 기타' 같은 항목명이 있음
            item = re.sub(r'^정산제외\s*-\s*', '', str(r[h['주문아이디']] or '').strip()) or '기타'
            bk.add_gross('수수료', None, f'올웨이즈 {item}', fee)
            etc[item][0] += 1; etc[item][1] += fee
    if etc:   # 정산제외 항목은 모두 수수료 (경영지원팀 확인 11/12)
        bk.info.append('정산제외 항목은 수수료로 넣음: ' + ', '.join(f'{k} {n}건 {a:,}원' for k, (n, a) in etc.items()))


JASON_PAY_TO_TAX = {'신용카드': '17', '계좌이체': '31'}   # 그 밖(포인트·휴대폰 등)은 14


def parse_jason(wb, bk):
    """제이슨딜(제이슨그룹): 매출 = 결제금액(배송비포함) 행별 (쿠폰·적립금은 제이슨 부담이라 이미 빠져 있음),
    수수료 = 제이슨 수익 = 매출 합계 − 총 지급 금액(공급가 + 배송비) 한 줄 (원본에 상품별 수수료 없음).
    부가세 = 신용카드 17 / 계좌이체 31 / 그 밖 14."""
    ws = sheet(wb, '§', '결제금액(배송비포함)', '공급가', '업체ID')
    rows = list(ws.iter_rows(values_only=True))
    hi = find_header(rows, '결제금액(배송비포함)', '공급가', '업체ID')
    h = header_index(rows[hi])
    sales = pay_out = 0
    for r in rows[hi + 1:]:
        if not r[h['업체ID']] or not r[h['주문번호']]:
            continue   # 맨 아래 합계·설명 행
        amt = num(r[h['결제금액(배송비포함)']])
        bk.add_gross('매출', r[h['통합상품코드']], r[h['상품명(전시)']], amt)
        bk.add_tax_gross(JASON_PAY_TO_TAX.get(str(r[h['결제방법']] or '').strip(), '14'), amt)
        sales += amt
        pay_out += num(r[h['공급가']]) + num(r[h['배송비']]) + num(r[h['CS배송비(공급처)']]) \
            + num(r[h['고객보상']]) + num(r[h['추가결제금(공급처)']])
    if sales:
        bk.add_gross('수수료', None, '제이슨딜 수수료(제이슨 수익)', sales - pay_out)
    bk.info.append(f'지급 금액 {pay_out:,}원 = 공급가 + 배송비, 수수료 = 매출 − 지급 금액')


# ── ESM(G마켓·옥션) ──────────────────────────────────────────────────────
# G마켓은 주문일 기준(매출용)과 매출기준일 기준(수수료용) 자료를 같이 받음 — 부가세 자료 총액과 맞는 쌍이 매출용.
ESM_VAT_TO_TAX = {'신용카드결제': '17', '신용카드': '17', '현금결제': '31', '현금영수증': '31',
                  '휴대폰결제': '14', '기타결제': '14', '기타': '14', '선불전자결제': '14', '선불전자': '14',
                  'G마켓쿠폰': '14'}


def _sheets_with(wb, *need):
    out = []
    for ws in wb.worksheets:
        if ws.title.startswith(EXTRA_SHEET):
            continue
        rows = list(ws.iter_rows(values_only=True))
        try:
            hi = find_header(rows, *need)
        except StopIteration:
            continue
        out.append((ws.title, header_index(rows[hi]), rows[hi + 1:]))
    return out


def _ship_row(r, h, col):
    v = r[h[col]] if r and len(r) > h[col] else None
    return v not in (None, '') and str(v).strip() != col   # 빈 줄·반복된 제목 줄 제외


def _esm_pick(bk, prods, ships, total, extra_sales, key_pay, key_ship):
    """부가세 총액 = 상품 결제금액 + 배송비 + extra_sales 인 (상품, 배송비) 시트 쌍을 고름."""
    def psum(p):
        return sum(num(r[p[1][key_pay]]) for r in p[2] if r and r[p[1][key_pay]] not in (None, ''))
    def ssum(sh):
        return sum(num(r[sh[1][key_ship]]) for r in sh[2] if _ship_row(r, sh[1], key_ship))
    ship_opts = ships or [None]
    pairs = [(p, sh) for p in prods for sh in ship_opts]
    pick = None
    if total is not None:
        for p, sh in pairs:
            if psum(p) + (ssum(sh) if sh else 0) + extra_sales == total:
                pick = (p, sh); break
    if pick is None:
        pick = pairs[0]
        if len(prods) > 1 or len(ships) > 1:
            bk.notes.append('상품·배송비 시트가 여러 벌인데 부가세 총액과 맞는 조합이 없어 첫 번째 시트로 변환함 — 시트 확인 필요')
        elif total is not None:
            got = psum(pick[0]) + (ssum(pick[1]) if pick[1] else 0) + extra_sales
            bk.notes.append(f'매출 {got:,}원이 부가세 자료 총액 {total:,}원과 다름 ({got - total:+,}원)')
    return pick


ESM_BRANDS = ('살림백서', '아우페', '헤디안', '프로클', '튼튼백서')


def _esm_unused(wb, bk, *groups):
    used = {x[0] for g in groups for x in (g if isinstance(g, list) else [g]) if x}
    left = [ws.title for ws in wb.worksheets if ws.title not in used and not ws.title.startswith(EXTRA_SHEET)]
    if left:   # 취소위약금·환불금차감 등 아직 규칙이 없는 시트
        bk.notes.append('변환에 안 쓴 시트가 있음: ' + ', '.join(left) + ' — 처음 보는 항목이면 규칙 추가가 필요해요')


def _esm_label(bk, prod):
    _, h, rows = prod
    names = ' '.join(str(r[h['상품명']] or '') for r in rows if r)
    found = [b for b in ESM_BRANDS if b in names]
    if found:
        bk.label = '·'.join(found)


def _esm_rows(bk, prod, ship, pay_col, ship_pay, ship_net, sales=True, fee=True):
    if prod:
        _, h, rows = prod
        for r in rows:
            if not r or r[h['주문번호']] in (None, '') or r[h[pay_col]] in (None, ''):
                continue
            pay, net = num(r[h[pay_col]]), num(r[h['판매자 최종정산금']])
            code, name = r[h['상품번호']], r[h['상품명']]
            if sales:
                bk.add_gross('매출', code, name, pay)
            if fee:
                bk.add_gross('수수료', code, name, pay - net)     # = 서비스이용료(감면 반영), 마이너스일 수 있음
    if ship:
        _, hs, srows = ship
        for r in srows:
            if not _ship_row(r, hs, ship_pay):
                continue
            if sales:
                bk.add_gross('매출', None, '배송비', num(r[hs[ship_pay]]))
            if fee:
                bk.add_gross('수수료', None, '배송비 서비스이용료', num(r[hs[ship_pay]]) - num(r[hs[ship_net]]))


def parse_gmarket(wb, bk):
    """G마켓 (경영지원팀 11/12): 매출은 주문일 기준 파일, 수수료는 매출기준일(배송완료) 기준 파일로 계산.
    한 파일에 두 기준 시트를 다 넣으면 부가세 자료 총액과 맞는 쪽을 매출용, 나머지를 수수료용으로 씀.
    매출 = 고객결제금 + 배송비 (+ 부가세 자료의 G마켓쿠폰), 수수료 = 고객결제금 − 판매자 최종정산금(= 서비스이용료,
    G마켓 부담 쿠폰이 기본이용료 감면으로 들어와 마이너스일 수 있음) + 배송비 서비스이용료.
    부가세 = 부가세 자료 결제수단별 합계 (신용카드 17 / 현금 31 / 휴대폰·기타·선불전자·G마켓쿠폰 14)."""
    prods = _sheets_with(wb, '고객결제금(구. 구매대금)', '판매자 최종정산금')
    ships = _sheets_with(wb, '배송비정산금액', '서비스이용료(선결제배송비)')
    vat = _sheets_with(wb, '총매출액', 'G마켓쿠폰')
    total, coupon, vrow = None, 0, None
    if vat:
        _, hv, vrows = vat[0]
        vrow = next((r for r in vrows if r and str(r[0] or '').strip() not in ('', '합계')), None)
        if vrow:
            total, coupon = num(vrow[hv['총매출액']]), num(vrow[hv['G마켓쿠폰']])
    prod, ship = _esm_pick(bk, prods, ships, total, coupon, '고객결제금(구. 구매대금)', '배송비')
    fprod = next((p for p in prods if p is not prod), None)
    fship = next((x for x in ships if x is not ship), None)
    _esm_label(bk, prod)
    _esm_rows(bk, prod, ship, '고객결제금(구. 구매대금)', '배송비', '배송비정산금액', fee=False)
    if fprod or fship:
        _esm_rows(bk, fprod or prod, fship or ship, '고객결제금(구. 구매대금)', '배송비', '배송비정산금액', sales=False)
        bk.info.append('매출: ' + ', '.join(t for t in [prod[0], ship[0] if ship else None] if t)
                       + ' (주문일 기준) / 수수료: ' + ', '.join(t for t in [(fprod or prod)[0], (fship or ship)[0] if (fship or ship) else None] if t)
                       + ' (매출기준일 기준)')
    else:
        _esm_rows(bk, prod, ship, '고객결제금(구. 구매대금)', '배송비', '배송비정산금액', sales=False)
        bk.info.append('기준일이 한 벌뿐이라 매출·수수료를 같은 시트로 계산함 (경영지원팀 확인 11/12: 금액 같으면 상관없음)')
    if coupon:
        bk.add_gross('매출', None, 'G마켓쿠폰', coupon)
        bk.info.append(f'부가세 자료의 G마켓쿠폰 {coupon:,}원을 매출 1줄로 넣음 (수수료는 상계하지 않음 — 경영지원팀 확인 11/12)')
    if vrow:
        for col, code in ESM_VAT_TO_TAX.items():
            if col in hv and num(vrow[hv[col]]):
                bk.add_tax_gross(code, num(vrow[hv[col]]))
    else:
        bk.tax_from_sales('14')
        bk.notes.append('G마켓 부가세 자료(총매출액·결제수단별) 시트가 없어 부가세 시트를 14(건별)로 채움 — 같은 파일에 넣어 주세요')
    _esm_unused(wb, bk, prods, ships, vat)


def parse_auction(wb, bk):
    """옥션: 매출 = 결제금액 + 배송비 + 환불금차감(반품 배송비 등), 수수료 = 매출 − 판매자 최종정산금(배송비는 배송수수료).
    부가세 = 부가세 상세 시트의 자료구분별 합계 (신용카드 17 / 현금영수증 31 / 휴대폰·기타·선불전자 14)."""
    prods = _sheets_with(wb, '옥션 구매자 쿠폰 할인', '판매자 최종정산금')
    both = _sheets_with(wb, '배송비 정산액', '배송수수료')
    ships = [x for x in both if '배송비' in x[1]]
    extras = [x for x in both if '배송비' not in x[1] and '금액' in x[1]]     # 환불금차감 등
    vat = _sheets_with(wb, '자료구분', '매출금액')
    total, vrows, hv = None, [], None
    if vat:
        _, hv, vrows = vat[0]
        vrows = [r for r in vrows if r and r[hv['매출금액']] not in (None, '')]
        total = sum(num(r[hv['매출금액']]) for r in vrows)
    extra_total = sum(num(r[x[1]['금액']]) for x in extras for r in x[2] if _ship_row(r, x[1], '금액'))
    prod, ship = _esm_pick(bk, prods, ships, total, extra_total, '결제금액', '배송비')
    _esm_label(bk, prod)
    _esm_rows(bk, prod, ship, '결제금액', '배송비', '배송비 정산액')
    for _, hx, xrows in extras:
        for r in xrows:
            if not _ship_row(r, hx, '금액'):
                continue
            kind = str(r[hx['종류']] or '').strip() if '종류' in hx else ''
            amt = num(r[hx['금액']])
            bk.add_gross('매출', None, f'배송비({kind})' if kind else '배송비', amt)
            if amt - num(r[hx['배송비 정산액']]):
                bk.add_gross('수수료', None, '배송비 서비스이용료', amt - num(r[hx['배송비 정산액']]))
    if vrows:
        for r in vrows:
            bk.add_tax_gross(ESM_VAT_TO_TAX.get(str(r[hv['자료구분']] or '').strip(), '14'), num(r[hv['매출금액']]))
    else:
        bk.tax_from_sales('14')
        bk.notes.append('옥션 부가세 상세 시트가 없어 부가세 시트를 14(건별)로 채움 — 같은 파일에 넣어 주세요')
    _esm_unused(wb, bk, prods, both, vat)


# 아마란스 매출처 전체 목록 — 화면 왼쪽에 표시 (변환 여부 확인용, 인식에는 안 씀)
# (아마란스 매출처명, 변환 규칙 이름 또는 None, 메모)
# 새 매출처 규칙을 만들면 두 번째 칸에 CHANNELS의 이름을 적으면 '변환 가능'으로 바뀜.
CHANNEL_LIST = [
    ('쿠팡(주)', None, ''),
    ('지마켓글로벌 유한책임회사(옥션)', '옥션', '아우페·헤디안 계정으로 추정'),
    ('지마켓', 'G마켓', '아우페·헤디안 계정으로 추정'),
    ('11번가(십일번가)', '11번가', ''),
    ('씨제이이엔엠(CJ오쇼핑)', 'CJ온스타일', ''),
    ('(주)공영홈쇼핑', '공영홈쇼핑', ''),
    ('(주)홈앤쇼핑', None, ''),
    ('쿠팡 주식회사(로켓배송)', None, ''),
    ('롯데닷컴-롯데쇼핑(주)e커머스사업본부', '롯데ON', ''),
    ('(주)우리홈쇼핑', '우리홈쇼핑', ''),
    ('(주)현대홈쇼핑', None, ''),
    ('(주)에스에스지닷컴', None, ''),
    ('(주)카카오', None, ''),
    ('네이버파이낸셜(주)', None, ''),
    ('에스케이스토아 주식회사', 'SK스토아', ''),
    ('엔에이치엔페이코(주)', None, ''),
    ('(주)티알엔', None, ''),
    ('Qoo10(큐텐)', None, ''),
    ('주식회사 레브잇(올웨이즈)', '올웨이즈', ''),
    ('(주)버킷플레이스', '오늘의집', ''),
    ('네이버페이(자사몰)', None, ''),
    ('지마켓-지마켓(프로클)', 'G마켓', ''),
    ('지마켓-옥션(프로클)', '옥션', ''),
    ('(주)지에스리테일홈쇼핑', None, ''),
    ('(주)더블유쇼핑', None, ''),
    ('주식회사 핌아시아', None, ''),
    ('네이버파이낸셜(주)_살림백서', None, ''),
    ('알리익스프레스코리아 유한회사', '알리익스프레스', '살림백서 계정으로 추정'),
    ('알리익스프레스코리아 유한회사_튼튼백서', None, ''),
    ('네이버파이낸셜(주)_튼튼백서', None, ''),
    ('(주)카카오_살림백서(카카오톡스토어)', '카카오', ''),
    ('(주)카카오_살림백서(선물하기)', None, '카카오 파일에 같이 들어옴 · 따로 나누기·부가세 자료 확인 필요'),
    ('네이버페이(살림백서)_np_hsuwx956659', None, ''),
    ('네이버페이(살림백서_np_fvctr791550)', None, ''),
    ('(주) 엘지씨엔에스', '(주) 엘지씨엔에스', ''),
    ('넛지헬스케어(주)', '넛지헬스케어캐시딜', ''),
    ('네이버파이낸셜(주)_아우페', None, ''),
    ('(주)카카오_아우페(카카오톡스토어)', '카카오', ''),
    ('주식회사 현대이지웰', '복지샵', '복지샵 파일로 추정'),
    ('(주)카카오_튼튼백서(카카오톡스토어)', None, ''),
    ('알리익스프레스코리아 유한회사(아우페)_살림백서', '알리익스프레스', '아우페 계정으로 추정'),
    ('(주) 비바리퍼블리카(토스)', '토스쇼핑', ''),
    ('(주)카카오_헤디안(카카오톡스토어)', '카카오', ''),
    ('씨제이이엔엠(CJ오쇼핑)_살림백서', 'CJ온스타일', ''),
    ('(주)제이슨그룹', '제이슨딜', ''),
    ('테무(TEMU)', None, ''),
    ('마켓컬리_살림백서', None, ''),
    ('에스케이엠앤서비스 주식회사(베네피아)', None, ''),
    ('테무(살림백서)', None, ''),
    ('지마켓(newnuon)', None, ''),
    ('지마켓글로벌 유한책임회사(newnuon)', None, ''),
    ('주식회사 티앤케이팩토리', None, ''),
    ('알리익스프레스(프로클)', '알리익스프레스', ''),
    ('테무(프로클)', None, ''),
    ('삼성웰스토리(주)', '삼성웰스토리', ''),
    ('(주)카카오_프로클(카카오톡스토어)', '카카오', ''),
    ('주식회사 화해글로벌', None, ''),
]

# 원본(raw) 파일로 아직 검증 안 한 매출처 — 화면 왼쪽 목록에 '원본 확인 중' 표시.
# 원본으로 확인이 끝나면 여기서 이름을 지우면 됨.
RAW_PENDING = {'오늘의집', '공영홈쇼핑', '우리홈쇼핑', 'CJ온스타일'}

# 원본으로 규칙은 만들었지만 경영지원팀 업로드 기준과 아직 대조 못 한 매출처 — '경영지원팀 확인 중' 표시.
RULE_PENDING = {'알리익스프레스', '삼성웰스토리', '복지샵', '위탁몰', '제이슨딜', '옥션'}

# 변환 규칙은 있지만 화면 목록에서 숨길 것
CHANNEL_HIDE = {'NS홈쇼핑'}

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
    ('카카오', lambda w: has(w, '판매정산금액') and has(w, '카카오할인금액합계'), parse_kakao),
    ('토스쇼핑', lambda w: has(w, '수수료 기준금액') and has(w, '토스 부담 할인 금액'), parse_toss),
    ('알리익스프레스', lambda w: has(w, '비용 항목명') and has(w, '수수료 항목명') and has(w, '총액(KRW)'), parse_ali),
    ('삼성웰스토리', lambda w: has(w, '최종매입공급금액') and has(w, '최종매입VAT'), parse_welstory),
    ('복지샵', lambda w: has(w, '복지샵') and has(w, '실정산금액'), parse_welfare),
    ('올웨이즈', lambda w: has(w, '매출에누리(올웨이즈 부담)') and has(w, '세금계산서 발급금액'), parse_always),
    ('제이슨딜', lambda w: has(w, '결제금액(배송비포함)') and has(w, '업체ID') and has(w, '공급가'), parse_jason),
    ('G마켓', lambda w: has(w, 'G마켓 구매자 쿠폰 할인') and has(w, '판매자 최종정산금'), parse_gmarket),
    ('옥션', lambda w: has(w, '옥션 구매자 쿠폰 할인') and has(w, '판매자 최종정산금'), parse_auction),
    ('위탁몰', lambda w: has(w, '총 결제금액(매출신고액)') and has(w, '위탁수수료 공급가액'), parse_consign),
    (POINTMALL_NAME, lambda w: has(w, '판매사수수료(총액)') and has(w, '대금지급액'), parse_pointmall),
]

# 단독으로는 변환하지 않고, 같은 기간의 본 파일에 붙여 쓰는 보조 파일
# (키, 판별 함수, 설명, 붙는 채널)
COMPANIONS = [
    ('11번가부가세', lambda w: has(w, '부가세신고내역상세') and has(w, '[정산]스토어묶음할인'), '11번가 부가세신고내역', '11번가'),
    ('카카오부가세', lambda w: has(w, '수수료차감금액') and has(w, '과세매출금액') and has(w, '판매자번호'), '카카오 부가세 파일', '카카오'),
    ('토스부가세', lambda w: has(w, '토스쇼핑 매출 에누리') or (has(w, '기타 매출 과세 공급가액') and has(w, '매출 에누리')), '토스 부가세 상세 내역', '토스쇼핑'),
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
            'sums': {}, 'tax': {}, 'notes': [], 'info': [], 'message': '', 'used': []}


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
        chl = f'{ch}({bk.label})' if bk.label else ch   # 화면 표시용 (카카오는 판매자번호로 브랜드 표시)
        out = os.path.join(out_dir, f'{base}_변환.xlsx')  # 결과 파일명 = 원본 파일명 + _변환
        n = 2
        while os.path.exists(out):   # 같은 이름이 있으면 (2), (3)… 붙임
            out = os.path.join(out_dir, f'{base}_변환({n}).xlsx'); n += 1
        write(out, bk)
        res.update(status='완료', channel=chl, out=out, notes=bk.notes, info=bk.info, used=sorted(bk.used))
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
    name_to_path = {os.path.basename(cp): cp for lst in comp.values() for _, _, cp in lst}
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
        results[p] = run_one(p, out_dir, extra, preloaded=(wb, words, period), companions=mine)
        used.update(name_to_path[n] for n in results[p].get('used', []) if n in name_to_path)
        if on_result:
            on_result(p, results[p])
    for p, (key, period, label, target) in comp_info.items():
        r = new_result(p)
        r['channel'] = target
        if p in used:
            r.update(status='보조', message=f'{label} — {target} 정산 파일과 함께 처리됨')
        else:
            r.update(status='건너뜀', message=f'{label} — 짝이 되는 {target} 정산 파일(같은 계정·기간)을 찾지 못함, 같이 추가하세요')
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
