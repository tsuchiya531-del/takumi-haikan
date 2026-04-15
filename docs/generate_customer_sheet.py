#!/usr/bin/env python3
"""配管洗浄専門店 匠 - 顧客管理スプレッドシート"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule
from openpyxl.worksheet.datavalidation import DataValidation

# ─── 色 ───
NAVY = '1A2A4A'
BLUE = '2C5F8A'
LIGHT_BLUE = 'E8EEF5'
WHITE = 'FFFFFF'
GRAY = 'F5F5F5'
GREEN_OK = 'E6F4EA'
RED_NG = 'FCE8E6'
YELLOW = 'FFF8E1'

thin = Side(style='thin', color='C8C8C8')
border_all = Border(left=thin, right=thin, top=thin, bottom=thin)


def build():
    wb = Workbook()

    # ═══════════════════════════════════════
    # シート1: 案件管理
    # ═══════════════════════════════════════
    ws = wb.active
    ws.title = '案件管理'

    headers = [
        ('案件No', 8),
        ('受付日', 11),
        ('受付時刻', 9),
        ('問合せ経路', 10),
        ('広告経由', 8),
        ('検索KW', 12),
        ('お客様名', 14),
        ('電話番号', 13),
        ('住所', 22),
        ('建物種別', 10),
        ('築年数', 7),
        ('症状概要', 20),
        ('訪問希望日', 11),
        ('訪問確定日', 11),
        ('訪問時刻', 9),
        ('担当者', 9),
        ('現地所見', 20),
        ('作業内容', 20),
        ('作業時間(分)', 10),
        ('見積額', 10),
        ('成約', 7),
        ('辞退理由', 14),
        ('実受領額', 10),
        ('材料費', 9),
        ('粗利', 10),
        ('支払方法', 10),
        ('完了日', 11),
        ('保証期限', 11),
        ('アフター', 8),
        ('口コミ依頼', 9),
        ('クレーム', 8),
        ('備考', 25),
    ]

    # ヘッダー行
    f_header = Font(size=9, bold=True, color=WHITE)
    fill_header = PatternFill('solid', fgColor=NAVY)
    for col_idx, (name, width) in enumerate(headers, 1):
        cell = ws.cell(1, col_idx, name)
        cell.font = f_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border_all
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.row_dimensions[1].height = 32
    ws.freeze_panes = 'C2'  # 案件No・受付日を固定

    # データ行（空の50行 + 書式設定）
    f_cell = Font(size=9)
    for r in range(2, 52):
        for c in range(1, len(headers) + 1):
            cell = ws.cell(r, c)
            cell.font = f_cell
            cell.border = border_all
            cell.alignment = Alignment(vertical='center', wrap_text=True)

        # 案件No 自動番号（手動で入力するフォーマット）
        ws.cell(r, 1).alignment = Alignment(horizontal='center', vertical='center')

        # 金額列の書式
        for col_letter_idx in [20, 23, 24, 25]:  # 見積額/実受領額/材料費/粗利
            ws.cell(r, col_letter_idx).number_format = '#,##0'

        # 粗利 = 実受領額 - 材料費（数式）
        ws.cell(r, 25).value = f'=IF(W{r}="","",W{r}-X{r})'

        # 保証期限 = 完了日 + 30日（数式）
        ws.cell(r, 28).value = f'=IF(AA{r}="","",AA{r}+30)'
        ws.cell(r, 28).number_format = 'yyyy/mm/dd'

        # 日付列の書式
        for col_idx in [2, 13, 14, 27, 28]:
            ws.cell(r, col_idx).number_format = 'yyyy/mm/dd'

        ws.row_dimensions[r].height = 22

    # データバリデーション（プルダウン）
    dv_route = DataValidation(type='list', formula1='"電話,LINE,Web,紹介,その他"', allow_blank=True)
    dv_route.add(f'D2:D100')
    ws.add_data_validation(dv_route)

    dv_yn = DataValidation(type='list', formula1='"YES,NO"', allow_blank=True)
    dv_yn.add('E2:E100')
    dv_yn.add('U2:U100')  # 成約
    dv_yn.add('AC2:AC100')  # アフター
    dv_yn.add('AD2:AD100')  # 口コミ依頼
    dv_yn.add('AE2:AE100')  # クレーム
    ws.add_data_validation(dv_yn)

    dv_building = DataValidation(type='list', formula1='"戸建て,マンション,アパート,その他"', allow_blank=True)
    dv_building.add('J2:J100')
    ws.add_data_validation(dv_building)

    dv_payment = DataValidation(type='list', formula1='"現金,銀行振込,カード,その他"', allow_blank=True)
    dv_payment.add('Z2:Z100')
    ws.add_data_validation(dv_payment)

    # 条件付き書式（成約=YES/NO で色分け）
    ws.conditional_formatting.add(
        'U2:U100',
        CellIsRule(operator='equal', formula=['"YES"'],
                   fill=PatternFill('solid', fgColor=GREEN_OK))
    )
    ws.conditional_formatting.add(
        'U2:U100',
        CellIsRule(operator='equal', formula=['"NO"'],
                   fill=PatternFill('solid', fgColor=RED_NG))
    )

    # クレームが"YES"なら赤
    ws.conditional_formatting.add(
        'AE2:AE100',
        CellIsRule(operator='equal', formula=['"YES"'],
                   fill=PatternFill('solid', fgColor=RED_NG))
    )

    # 記入例（2行目）
    sample = [
        '001', '2026-04-15', '10:30', '電話', 'YES', '排水 詰まり 名古屋',
        '山田 太郎', '090-1234-5678', '名古屋市中川区○○町1-2-3', '戸建て', 25,
        'キッチン排水が流れない。ゴボゴボ音あり',
        '2026-04-15', '2026-04-15', '14:00', '土屋',
        '排水管の油脂蓄積。屋外升も汚泥多め',
        'キッチン排水管 高圧洗浄 + 屋外升清掃', 90,
        33000, 'YES', '', 33000, 2500, None,  # 粗利は数式
        '現金', '2026-04-15', None, 'YES', 'NO', 'NO',  # 保証期限は数式
        '口コミ投稿依頼済（LINEで依頼）'
    ]
    for i, val in enumerate(sample, 1):
        if val is not None:
            ws.cell(2, i, val)

    # ═══════════════════════════════════════
    # シート2: ダッシュボード
    # ═══════════════════════════════════════
    ws2 = wb.create_sheet('ダッシュボード')

    # 列幅
    for col, w in {'A': 4, 'B': 22, 'C': 16, 'D': 16, 'E': 4, 'F': 22, 'G': 16}.items():
        ws2.column_dimensions[col].width = w

    # タイトル
    ws2.merge_cells('B2:G2')
    ws2.cell(2, 2, '匠 - 経営ダッシュボード').font = Font(size=16, bold=True, color=NAVY)
    ws2.cell(2, 2).alignment = Alignment(horizontal='center', vertical='center')
    ws2.row_dimensions[2].height = 28

    # サマリーブロック1: 問合せ・成約
    ws2.merge_cells('B4:D4')
    ws2.cell(4, 2, '問合せ・成約').font = Font(size=11, bold=True, color=WHITE)
    ws2.cell(4, 2).fill = PatternFill('solid', fgColor=BLUE)
    ws2.cell(4, 2).alignment = Alignment(horizontal='center', vertical='center')
    ws2.row_dimensions[4].height = 22

    kpis_left = [
        ('今月の受付件数', '=COUNTA(案件管理!B2:B200)-COUNTIF(案件管理!B2:B200,"<>"&"")+COUNTIFS(案件管理!B2:B200,">="&EOMONTH(TODAY(),-1)+1,案件管理!B2:B200,"<="&EOMONTH(TODAY(),0))', '件'),
        ('今月の成約件数', '=COUNTIFS(案件管理!B2:B200,">="&EOMONTH(TODAY(),-1)+1,案件管理!B2:B200,"<="&EOMONTH(TODAY(),0),案件管理!U2:U200,"YES")', '件'),
        ('成約率', '=IFERROR(COUNTIFS(案件管理!B2:B200,">="&EOMONTH(TODAY(),-1)+1,案件管理!B2:B200,"<="&EOMONTH(TODAY(),0),案件管理!U2:U200,"YES")/COUNTIFS(案件管理!B2:B200,">="&EOMONTH(TODAY(),-1)+1,案件管理!B2:B200,"<="&EOMONTH(TODAY(),0)),"")', '%'),
        ('電話経由 比率', '=IFERROR(COUNTIFS(案件管理!B2:B200,">="&EOMONTH(TODAY(),-1)+1,案件管理!B2:B200,"<="&EOMONTH(TODAY(),0),案件管理!D2:D200,"電話")/COUNTIFS(案件管理!B2:B200,">="&EOMONTH(TODAY(),-1)+1,案件管理!B2:B200,"<="&EOMONTH(TODAY(),0)),"")', '%'),
        ('LINE経由 比率', '=IFERROR(COUNTIFS(案件管理!B2:B200,">="&EOMONTH(TODAY(),-1)+1,案件管理!B2:B200,"<="&EOMONTH(TODAY(),0),案件管理!D2:D200,"LINE")/COUNTIFS(案件管理!B2:B200,">="&EOMONTH(TODAY(),-1)+1,案件管理!B2:B200,"<="&EOMONTH(TODAY(),0)),"")', '%'),
    ]

    for i, (label, formula, unit) in enumerate(kpis_left):
        r = 5 + i
        ws2.cell(r, 2, label).font = Font(size=10)
        ws2.cell(r, 2).fill = PatternFill('solid', fgColor=GRAY)
        ws2.cell(r, 2).border = border_all
        ws2.cell(r, 2).alignment = Alignment(horizontal='left', vertical='center', indent=1)

        ws2.cell(r, 3, formula).font = Font(size=11, bold=True, color=NAVY)
        ws2.cell(r, 3).alignment = Alignment(horizontal='right', vertical='center')
        ws2.cell(r, 3).border = border_all
        if unit == '%':
            ws2.cell(r, 3).number_format = '0.0%'
        else:
            ws2.cell(r, 3).number_format = '#,##0'

        ws2.cell(r, 4, unit).font = Font(size=9, color='666666')
        ws2.cell(r, 4).alignment = Alignment(horizontal='left', vertical='center')
        ws2.cell(r, 4).border = border_all
        ws2.row_dimensions[r].height = 22

    # サマリーブロック2: 売上・粗利
    ws2.merge_cells('F4:G4')
    ws2.cell(4, 6, '売上・粗利').font = Font(size=11, bold=True, color=WHITE)
    ws2.cell(4, 6).fill = PatternFill('solid', fgColor=BLUE)
    ws2.cell(4, 6).alignment = Alignment(horizontal='center', vertical='center')

    kpis_right = [
        ('今月売上', '=SUMIFS(案件管理!W2:W200,案件管理!B2:B200,">="&EOMONTH(TODAY(),-1)+1,案件管理!B2:B200,"<="&EOMONTH(TODAY(),0))', '円'),
        ('今月粗利', '=SUMIFS(案件管理!Y2:Y200,案件管理!B2:B200,">="&EOMONTH(TODAY(),-1)+1,案件管理!B2:B200,"<="&EOMONTH(TODAY(),0))', '円'),
        ('平均客単価', '=IFERROR(AVERAGEIFS(案件管理!W2:W200,案件管理!B2:B200,">="&EOMONTH(TODAY(),-1)+1,案件管理!B2:B200,"<="&EOMONTH(TODAY(),0),案件管理!U2:U200,"YES"),"")', '円'),
        ('平均粗利', '=IFERROR(AVERAGEIFS(案件管理!Y2:Y200,案件管理!B2:B200,">="&EOMONTH(TODAY(),-1)+1,案件管理!B2:B200,"<="&EOMONTH(TODAY(),0),案件管理!U2:U200,"YES"),"")', '円'),
        ('粗利率', '=IFERROR(SUMIFS(案件管理!Y2:Y200,案件管理!B2:B200,">="&EOMONTH(TODAY(),-1)+1,案件管理!B2:B200,"<="&EOMONTH(TODAY(),0))/SUMIFS(案件管理!W2:W200,案件管理!B2:B200,">="&EOMONTH(TODAY(),-1)+1,案件管理!B2:B200,"<="&EOMONTH(TODAY(),0)),"")', '%'),
    ]

    for i, (label, formula, unit) in enumerate(kpis_right):
        r = 5 + i
        ws2.cell(r, 6, label).font = Font(size=10)
        ws2.cell(r, 6).fill = PatternFill('solid', fgColor=GRAY)
        ws2.cell(r, 6).border = border_all
        ws2.cell(r, 6).alignment = Alignment(horizontal='left', vertical='center', indent=1)

        ws2.cell(r, 7, formula).font = Font(size=11, bold=True, color=NAVY)
        ws2.cell(r, 7).alignment = Alignment(horizontal='right', vertical='center')
        ws2.cell(r, 7).border = border_all
        if unit == '%':
            ws2.cell(r, 7).number_format = '0.0%'
        else:
            ws2.cell(r, 7).number_format = '#,##0"円"'
        ws2.row_dimensions[r].height = 22

    # アラート
    ws2.merge_cells('B12:G12')
    ws2.cell(12, 2, '要フォロー').font = Font(size=11, bold=True, color=WHITE)
    ws2.cell(12, 2).fill = PatternFill('solid', fgColor='C44A4A')
    ws2.cell(12, 2).alignment = Alignment(horizontal='center', vertical='center')

    alerts = [
        ('クレーム件数（今月）', '=COUNTIFS(案件管理!B2:B200,">="&EOMONTH(TODAY(),-1)+1,案件管理!B2:B200,"<="&EOMONTH(TODAY(),0),案件管理!AE2:AE200,"YES")'),
        ('アフター未対応', '=COUNTIFS(案件管理!U2:U200,"YES",案件管理!AC2:AC200,"NO")'),
        ('保証期間内の案件数', '=COUNTIFS(案件管理!AB2:AB200,">="&TODAY())'),
        ('口コミ依頼未実施', '=COUNTIFS(案件管理!U2:U200,"YES",案件管理!AD2:AD200,"NO")'),
    ]
    for i, (label, formula) in enumerate(alerts):
        r = 13 + i
        ws2.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        ws2.cell(r, 2, label).font = Font(size=10)
        ws2.cell(r, 2).fill = PatternFill('solid', fgColor='FFF0F0')
        ws2.cell(r, 2).border = border_all
        ws2.cell(r, 2).alignment = Alignment(horizontal='left', vertical='center', indent=1)
        ws2.cell(r, 7, formula).font = Font(size=11, bold=True, color='C44A4A')
        ws2.cell(r, 7).alignment = Alignment(horizontal='right', vertical='center')
        ws2.cell(r, 7).border = border_all
        ws2.cell(r, 7).number_format = '0"件"'
        ws2.row_dimensions[r].height = 22

    # 注釈
    ws2.merge_cells('B18:G20')
    note = ('※ データは「案件管理」シートに入力するとここに自動集計されます。\n'
            '※ 「受付日」を入れると当月集計に反映されます。\n'
            '※ KPI詳細分析は「月次収支表.xlsx」と併せてご確認ください。')
    ws2.cell(18, 2, note).font = Font(size=8, color='666666')
    ws2.cell(18, 2).alignment = Alignment(wrap_text=True, vertical='top')

    # ウィンドウ設定
    ws2.sheet_view.showGridLines = False

    # ═══════════════════════════════════════
    # 保存
    # ═══════════════════════════════════════
    path = '/Users/tsuchiyatakumiki/配管洗浄専門店/docs/顧客管理.xlsx'
    wb.save(path)
    print(f'顧客管理シート生成完了: {path}')


if __name__ == '__main__':
    build()
