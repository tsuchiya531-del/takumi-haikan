#!/usr/bin/env python3
"""配管洗浄専門店 匠 - 月次収支表"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

NAVY = '1A2A4A'
BLUE = '2C5F8A'
LIGHT_BLUE = 'E8EEF5'
WHITE = 'FFFFFF'
GRAY = 'F5F5F5'
CREAM = 'FFF5E6'
GREEN = '2E7D32'
RED = 'C62828'

thin = Side(style='thin', color='C8C8C8')
border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
border_bottom = Border(bottom=Side(style='medium', color=NAVY))

MONTHS = ['4月', '5月', '6月', '7月', '8月', '9月',
          '10月', '11月', '12月', '1月', '2月', '3月']


def styled_header(cell, bg=NAVY, color=WHITE):
    cell.font = Font(size=10, bold=True, color=color)
    cell.fill = PatternFill('solid', fgColor=bg)
    cell.alignment = Alignment(horizontal='center', vertical='center')
    cell.border = border_all


def build():
    wb = Workbook()

    # ═══════════════════════════════════════
    # シート1: 月次収支表（12ヶ月分）
    # ═══════════════════════════════════════
    ws = wb.active
    ws.title = '月次収支表'
    ws.sheet_view.showGridLines = False

    # 列幅
    ws.column_dimensions['A'].width = 3
    ws.column_dimensions['B'].width = 22  # 項目名
    for i, _ in enumerate(MONTHS):
        ws.column_dimensions[get_column_letter(3 + i)].width = 12
    ws.column_dimensions[get_column_letter(3 + len(MONTHS))].width = 14  # 年間合計
    ws.column_dimensions[get_column_letter(3 + len(MONTHS) + 1)].width = 2

    # ─── タイトル ───
    ws.merge_cells('B2:P2')
    ws.cell(2, 2, '配管洗浄専門店 匠 - 月次収支表（2026年度）').font = Font(size=16, bold=True, color=NAVY)
    ws.cell(2, 2).alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[2].height = 28

    # ─── 月ヘッダー行 ───
    row = 4
    styled_header(ws.cell(row, 2, '項目'))
    for i, month in enumerate(MONTHS):
        styled_header(ws.cell(row, 3 + i, month))
    styled_header(ws.cell(row, 3 + len(MONTHS), '年間合計'), bg=BLUE)
    ws.row_dimensions[row].height = 22

    last_col_letter = get_column_letter(3 + len(MONTHS) - 1)  # O
    total_col_letter = get_column_letter(3 + len(MONTHS))     # P

    # ヘルパー: 行を追加する関数
    def add_row(r, label, formula=None, indent=0, bold=False, fill=None, font_color=None,
                number_format='#,##0', is_sum=False):
        cell = ws.cell(r, 2, label)
        cell.alignment = Alignment(horizontal='left', vertical='center', indent=indent)
        cell.font = Font(size=10, bold=bold, color=font_color or '333333')
        cell.border = border_all
        if fill:
            cell.fill = PatternFill('solid', fgColor=fill)

        # 12ヶ月の金額セル
        for i in range(len(MONTHS)):
            c = ws.cell(r, 3 + i)
            c.border = border_all
            c.number_format = number_format
            if fill:
                c.fill = PatternFill('solid', fgColor=fill)
            if bold:
                c.font = Font(size=10, bold=True, color=font_color or '333333')
            if formula:
                col_letter = get_column_letter(3 + i)
                c.value = formula.replace('{col}', col_letter)

        # 年間合計
        total_cell = ws.cell(r, 3 + len(MONTHS))
        total_cell.border = border_all
        total_cell.number_format = number_format
        total_cell.font = Font(size=10, bold=True, color=NAVY)
        if fill:
            total_cell.fill = PatternFill('solid', fgColor=fill)
        # 年間合計 = SUM(月次)
        total_cell.value = f'=SUM(C{r}:{last_col_letter}{r})'

        ws.row_dimensions[r].height = 20

    # ─── 売上 ───
    row += 1  # 5
    add_row(row, '【売上】', bold=True, fill=LIGHT_BLUE, font_color=NAVY)
    row += 1
    add_row(row, '受注件数', indent=1, number_format='0"件"')
    件数行 = row
    row += 1
    add_row(row, '売上高（税込）', indent=1)
    売上行 = row

    # ─── 変動費 ───
    row += 1
    add_row(row, '【変動費】', bold=True, fill=LIGHT_BLUE, font_color=NAVY)
    row += 1
    add_row(row, '材料費', indent=1)
    材料費行 = row
    row += 1
    add_row(row, '車両費（ガソリン・駐車場）', indent=1)
    車両費行 = row
    row += 1
    add_row(row, '外注費', indent=1)
    外注費行 = row
    row += 1
    add_row(row, '変動費 計',
            formula=f'=SUM({{col}}{材料費行}:{{col}}{外注費行})',
            indent=1, bold=True, fill=GRAY)
    変動費計行 = row

    # ─── 売上総利益（粗利） ───
    row += 1
    add_row(row, '売上総利益（粗利）',
            formula=f'={{col}}{売上行}-{{col}}{変動費計行}',
            bold=True, fill=CREAM, font_color=NAVY)
    粗利行 = row
    row += 1
    add_row(row, '粗利率', indent=1,
            formula=f'=IFERROR({{col}}{粗利行}/{{col}}{売上行},"")',
            number_format='0.0%')

    # ─── 固定費 ───
    row += 1
    add_row(row, '【固定費】', bold=True, fill=LIGHT_BLUE, font_color=NAVY)

    row += 1
    add_row(row, 'Google広告費', indent=1)
    広告行 = row
    row += 1
    add_row(row, 'その他広告（チラシ等）', indent=1)
    row += 1
    add_row(row, '通信費（IVRy/LINE）', indent=1)
    row += 1
    add_row(row, 'ドメイン・サーバー', indent=1)
    row += 1
    add_row(row, '保険', indent=1)
    row += 1
    add_row(row, '水道光熱費（事務所）', indent=1)
    row += 1
    add_row(row, '消耗品・備品', indent=1)
    row += 1
    add_row(row, '交際費', indent=1)
    row += 1
    add_row(row, 'その他', indent=1)
    固定費末尾行 = row
    row += 1
    add_row(row, '固定費 計',
            formula=f'=SUM({{col}}{広告行}:{{col}}{固定費末尾行})',
            indent=1, bold=True, fill=GRAY)
    固定費計行 = row

    # ─── 営業利益 ───
    row += 1
    add_row(row, '営業利益',
            formula=f'={{col}}{粗利行}-{{col}}{固定費計行}',
            bold=True, fill='FFE5CC', font_color=NAVY)
    営業利益行 = row
    row += 1
    add_row(row, '営業利益率', indent=1,
            formula=f'=IFERROR({{col}}{営業利益行}/{{col}}{売上行},"")',
            number_format='0.0%')

    # ─── KPI ───
    row += 2
    add_row(row, '【KPI】', bold=True, fill=LIGHT_BLUE, font_color=NAVY)
    row += 1
    add_row(row, '平均客単価', indent=1,
            formula=f'=IFERROR({{col}}{売上行}/{{col}}{件数行},"")')
    row += 1
    add_row(row, 'CPA（広告費÷成約件数）', indent=1,
            formula=f'=IFERROR({{col}}{広告行}/{{col}}{件数行},"")')
    row += 1
    add_row(row, '1件あたり粗利', indent=1,
            formula=f'=IFERROR({{col}}{粗利行}/{{col}}{件数行},"")')

    # ─── 注釈 ───
    row += 3
    ws.merge_cells(start_row=row, start_column=2, end_row=row+3, end_column=15)
    note = (
        '【記入ルール】\n'
        '・「受注件数」「売上高」は『顧客管理.xlsx』の当月実績を転記\n'
        '・各費用は実費を税込みで記入\n'
        '・粗利/営業利益/合計/KPIは自動計算\n'
        '・営業利益がマイナスの月は固定費の見直しを検討（CFOに相談）'
    )
    cell = ws.cell(row, 2, note)
    cell.font = Font(size=9, color='666666')
    cell.alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')

    # ウィンドウ固定
    ws.freeze_panes = 'C5'

    # ═══════════════════════════════════════
    # シート2: 開業コスト・損益分岐点
    # ═══════════════════════════════════════
    ws2 = wb.create_sheet('損益分岐点')
    ws2.sheet_view.showGridLines = False

    ws2.column_dimensions['A'].width = 3
    ws2.column_dimensions['B'].width = 24
    ws2.column_dimensions['C'].width = 16
    ws2.column_dimensions['D'].width = 4
    ws2.column_dimensions['E'].width = 24
    ws2.column_dimensions['F'].width = 16

    # タイトル
    ws2.merge_cells('B2:F2')
    ws2.cell(2, 2, '損益分岐点シミュレーション').font = Font(size=16, bold=True, color=NAVY)
    ws2.cell(2, 2).alignment = Alignment(horizontal='center', vertical='center')
    ws2.row_dimensions[2].height = 28

    # 左：前提数値
    ws2.merge_cells('B4:C4')
    ws2.cell(4, 2, '前提数値（月次）').font = Font(size=11, bold=True, color=WHITE)
    ws2.cell(4, 2).fill = PatternFill('solid', fgColor=BLUE)
    ws2.cell(4, 2).alignment = Alignment(horizontal='center', vertical='center')
    ws2.row_dimensions[4].height = 22

    inputs = [
        ('平均客単価（税込）', 33000),
        ('1件あたり材料費', 2500),
        ('1件あたり車両費', 500),
        ('固定費 月額（広告費除く）', 30000),
        ('Google広告費 月額', 30000),
        ('目標営業利益 月額', 200000),
    ]
    for i, (label, val) in enumerate(inputs):
        r = 5 + i
        ws2.cell(r, 2, label).font = Font(size=10)
        ws2.cell(r, 2).fill = PatternFill('solid', fgColor=GRAY)
        ws2.cell(r, 2).border = border_all
        ws2.cell(r, 2).alignment = Alignment(horizontal='left', vertical='center', indent=1)
        ws2.cell(r, 3, val).font = Font(size=11, bold=True)
        ws2.cell(r, 3).number_format = '#,##0"円"'
        ws2.cell(r, 3).border = border_all
        ws2.cell(r, 3).alignment = Alignment(horizontal='right', vertical='center')
        ws2.row_dimensions[r].height = 22

    # 右：試算結果
    ws2.merge_cells('E4:F4')
    ws2.cell(4, 5, '試算結果').font = Font(size=11, bold=True, color=WHITE)
    ws2.cell(4, 5).fill = PatternFill('solid', fgColor=BLUE)
    ws2.cell(4, 5).alignment = Alignment(horizontal='center', vertical='center')

    # セル参照
    # C5 = 客単価, C6 = 材料費, C7 = 車両費, C8 = 固定費, C9 = 広告費, C10 = 目標利益
    outputs = [
        ('1件あたり粗利', '=C5-C6-C7', '#,##0"円"'),
        ('粗利率', '=IFERROR((C5-C6-C7)/C5,"")', '0.0%'),
        ('固定費合計（月）', '=C8+C9', '#,##0"円"'),
        ('損益分岐点（件数）', '=IFERROR(ROUNDUP((C8+C9)/(C5-C6-C7),0),"")', '0"件"'),
        ('損益分岐点（売上）', '=IFERROR(ROUNDUP((C8+C9)/(C5-C6-C7),0)*C5,"")', '#,##0"円"'),
        ('目標達成に必要な件数', '=IFERROR(ROUNDUP((C8+C9+C10)/(C5-C6-C7),0),"")', '0"件"'),
        ('目標達成に必要な売上', '=IFERROR(ROUNDUP((C8+C9+C10)/(C5-C6-C7),0)*C5,"")', '#,##0"円"'),
    ]
    for i, (label, formula, fmt) in enumerate(outputs):
        r = 5 + i
        ws2.cell(r, 5, label).font = Font(size=10)
        ws2.cell(r, 5).fill = PatternFill('solid', fgColor=GRAY)
        ws2.cell(r, 5).border = border_all
        ws2.cell(r, 5).alignment = Alignment(horizontal='left', vertical='center', indent=1)
        ws2.cell(r, 6, formula).font = Font(size=11, bold=True, color=NAVY)
        ws2.cell(r, 6).number_format = fmt
        ws2.cell(r, 6).border = border_all
        ws2.cell(r, 6).alignment = Alignment(horizontal='right', vertical='center')
        ws2.row_dimensions[r].height = 22

    # 注釈
    ws2.merge_cells('B13:F16')
    note = (
        '【使い方】\n'
        '・左の「前提数値」を変更すると、右の試算結果が自動で更新されます\n'
        '・広告費を増やした時に何件取れば成り立つか、客単価を上げた時の効果などをシミュレーション\n'
        '・損益分岐点: 毎月これだけの件数をこなさないと赤字になる水準\n'
        '・目標営業利益: 月の「取りたい利益」を入れると、必要な件数が出ます'
    )
    cell = ws2.cell(13, 2, note)
    cell.font = Font(size=9, color='666666')
    cell.alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')

    # ═══════════════════════════════════════
    # シート3: 広告費管理
    # ═══════════════════════════════════════
    ws3 = wb.create_sheet('広告費管理')
    ws3.sheet_view.showGridLines = False

    ws3.column_dimensions['A'].width = 3
    ws3.column_dimensions['B'].width = 14  # 日付
    ws3.column_dimensions['C'].width = 14  # 媒体
    ws3.column_dimensions['D'].width = 14  # 費用
    ws3.column_dimensions['E'].width = 12  # 表示回数
    ws3.column_dimensions['F'].width = 12  # クリック数
    ws3.column_dimensions['G'].width = 10  # CPC
    ws3.column_dimensions['H'].width = 12  # 問合せ数
    ws3.column_dimensions['I'].width = 12  # CPA
    ws3.column_dimensions['J'].width = 20  # メモ

    # タイトル
    ws3.merge_cells('B2:J2')
    ws3.cell(2, 2, '広告費管理（日別）').font = Font(size=14, bold=True, color=NAVY)
    ws3.cell(2, 2).alignment = Alignment(horizontal='center', vertical='center')
    ws3.row_dimensions[2].height = 26

    # ヘッダー
    ad_headers = ['日付', '媒体', '費用', '表示回数', 'クリック数', 'CPC', '問合せ数', 'CPA', 'メモ']
    for i, h in enumerate(ad_headers):
        c = ws3.cell(4, 2 + i, h)
        styled_header(c)
    ws3.row_dimensions[4].height = 22

    # 60行分の入力枠
    for r in range(5, 65):
        for i in range(len(ad_headers)):
            c = ws3.cell(r, 2 + i)
            c.border = border_all
            c.font = Font(size=10)
            c.alignment = Alignment(vertical='center', horizontal='center')
        # CPC = 費用 / クリック数
        ws3.cell(r, 7, f'=IFERROR(D{r}/F{r},"")')
        ws3.cell(r, 7).number_format = '#,##0"円"'
        # CPA = 費用 / 問合せ数
        ws3.cell(r, 9, f'=IFERROR(D{r}/H{r},"")')
        ws3.cell(r, 9).number_format = '#,##0"円"'
        # 書式
        ws3.cell(r, 2).number_format = 'yyyy/mm/dd'
        ws3.cell(r, 4).number_format = '#,##0"円"'
        ws3.cell(r, 5).number_format = '#,##0'
        ws3.cell(r, 6).number_format = '#,##0'
        ws3.cell(r, 8).number_format = '#,##0'

    ws3.freeze_panes = 'B5'

    # サンプル行
    ws3.cell(5, 2, '2026-04-01')
    ws3.cell(5, 3, 'Google広告')
    ws3.cell(5, 4, 1000)
    ws3.cell(5, 5, 120)
    ws3.cell(5, 6, 8)
    ws3.cell(5, 8, 1)
    ws3.cell(5, 10, '初日・キーワード:排水 詰まり').alignment = Alignment(horizontal='left', vertical='center')

    # ═══════════════════════════════════════
    # 保存
    # ═══════════════════════════════════════
    path = '/Users/tsuchiyatakumiki/配管洗浄専門店/docs/月次収支表.xlsx'
    wb.save(path)
    print(f'月次収支表生成完了: {path}')


if __name__ == '__main__':
    build()
