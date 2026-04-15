#!/usr/bin/env python3
"""配管洗浄専門店 匠 - お客様用見積書テンプレート（Excel / Googleスプレッドシート対応）"""

from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side,
    numbers
)
from openpyxl.utils import get_column_letter

def build_xlsx():
    wb = Workbook()
    ws = wb.active
    ws.title = '御見積書'

    # ─── 色定義 ───
    NAVY = '1A2A4A'
    BLUE = '2C5F8A'
    WHITE = 'FFFFFF'
    LIGHT_GRAY = 'F5F5F5'
    MID_GRAY = 'C8C8C8'
    DARK_TEXT = '333333'
    CREAM = 'FFF5E6'

    # ─── フォント定義 ───
    f_title = Font(size=20, bold=True, color=WHITE)
    f_subtitle = Font(size=9, color=WHITE)
    f_estimate_title = Font(size=18, bold=True, color=WHITE)
    f_header_small = Font(size=8, color=WHITE)
    f_bold_navy = Font(size=10, bold=True, color=NAVY)
    f_bold_9 = Font(size=9, bold=True, color=DARK_TEXT)
    f_normal = Font(size=9, color=DARK_TEXT)
    f_normal_s = Font(size=8, color=DARK_TEXT)
    f_small = Font(size=8, color='808080')
    f_label = Font(size=8, color=DARK_TEXT)
    f_table_header = Font(size=8, bold=True, color=WHITE)
    f_total_label = Font(size=10, bold=True, color=WHITE)
    f_total_amount = Font(size=13, bold=True, color=NAVY)
    f_promise_title = Font(size=8, bold=True, color=NAVY)
    f_promise = Font(size=7.5, color=DARK_TEXT)
    f_note = Font(size=7, color='808080')
    f_section = Font(size=8, bold=True, color=WHITE)
    f_seal_label = Font(size=8, color=DARK_TEXT)
    f_footer = Font(size=7, color='888888')

    # ─── 塗りつぶし定義 ───
    fill_navy = PatternFill('solid', fgColor=NAVY)
    fill_blue = PatternFill('solid', fgColor=BLUE)
    fill_gray = PatternFill('solid', fgColor=LIGHT_GRAY)
    fill_white = PatternFill('solid', fgColor=WHITE)
    fill_cream = PatternFill('solid', fgColor=CREAM)
    fill_light_blue = PatternFill('solid', fgColor='F0F4F8')

    # ─── 罫線定義 ───
    thin = Side(style='thin', color=MID_GRAY)
    thin_navy = Side(style='thin', color=NAVY)
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    border_bottom = Border(bottom=Side(style='medium', color=NAVY))
    border_none = Border()

    # ─── 列幅設定 ───
    col_widths = {
        'A': 4,    # No.
        'B': 6,    # ラベル等
        'C': 14,   # ラベル
        'D': 18,   # 値
        'E': 14,   # 値
        'F': 12,   # 数量
        'G': 14,   # 単価
        'H': 14,   # 金額
        'I': 10,   # 区分
        'J': 8,    # 印鑑
    }
    for col, w in col_widths.items():
        ws.column_dimensions[col].width = w

    # ─── 印刷設定 ───
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = 'portrait'
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_area = 'A1:J58'

    row = 1

    # ═══════════════════════════════════════
    # ヘッダー（紺色バー）
    # ═══════════════════════════════════════
    for r in range(row, row + 3):
        for c in range(1, 11):
            ws.cell(r, c).fill = fill_navy

    ws.merge_cells(f'A{row}:E{row}')
    ws.cell(row, 1, '配管洗浄・詰まり抜き専門店').font = f_subtitle
    ws.cell(row, 1).alignment = Alignment(vertical='center')

    ws.merge_cells(f'F{row}:J{row + 2}')
    ws.cell(row, 6, '御 見 積 書').font = f_estimate_title
    ws.cell(row, 6).alignment = Alignment(horizontal='right', vertical='center')

    row += 1  # 2
    ws.merge_cells(f'A{row}:E{row}')
    ws.cell(row, 1, '匠').font = f_title
    ws.cell(row, 1).alignment = Alignment(vertical='center')
    ws.row_dimensions[row].height = 28

    row += 1  # 3
    ws.merge_cells(f'A{row}:E{row}')
    ws.cell(row, 1, 'TEL: 050-1725-6168  |  takumi-haikan.com').font = f_header_small
    ws.cell(row, 1).alignment = Alignment(vertical='center')

    row += 1  # 4 空行
    ws.row_dimensions[row].height = 6

    # ═══════════════════════════════════════
    # お客様情報（左） + 見積番号等（右）
    # ═══════════════════════════════════════
    row += 1  # 5

    # 左: お客様名
    ws.merge_cells(f'A{row}:D{row}')
    c = ws.cell(row, 1)
    c.font = Font(size=11, bold=True, color=DARK_TEXT)
    c.border = border_bottom
    c.alignment = Alignment(vertical='bottom')
    for cc in range(2, 5):
        ws.cell(row, cc).border = border_bottom
    # 「様」
    ws.cell(row, 5, '様').font = Font(size=11, bold=True, color=DARK_TEXT)
    ws.cell(row, 5).alignment = Alignment(horizontal='left', vertical='bottom')
    ws.cell(row, 5).border = border_bottom

    # 右: 見積番号
    ws.merge_cells(f'F{row}:G{row}')
    ws.cell(row, 6, '見積番号:').font = f_label
    ws.cell(row, 6).alignment = Alignment(horizontal='right', vertical='center')
    ws.merge_cells(f'H{row}:J{row}')
    ws.cell(row, 8).border = Border(bottom=thin)

    row += 1  # 6 ご住所
    ws.cell(row, 1, 'ご住所').font = f_label
    ws.cell(row, 1).fill = fill_gray
    ws.cell(row, 1).border = border_all
    ws.merge_cells(f'A{row}:B{row}')
    ws.cell(row, 2).fill = fill_gray
    ws.cell(row, 2).border = border_all
    ws.merge_cells(f'C{row}:E{row}')
    ws.cell(row, 3).border = border_all

    ws.merge_cells(f'F{row}:G{row}')
    ws.cell(row, 6, '発行日:').font = f_label
    ws.cell(row, 6).alignment = Alignment(horizontal='right', vertical='center')
    ws.merge_cells(f'H{row}:J{row}')
    ws.cell(row, 8).border = Border(bottom=thin)
    ws.cell(row, 8).font = f_normal_s

    row += 1  # 7 お電話
    ws.cell(row, 1, 'お電話').font = f_label
    ws.cell(row, 1).fill = fill_gray
    ws.cell(row, 1).border = border_all
    ws.merge_cells(f'A{row}:B{row}')
    ws.cell(row, 2).fill = fill_gray
    ws.cell(row, 2).border = border_all
    ws.merge_cells(f'C{row}:E{row}')
    ws.cell(row, 3).border = border_all

    ws.merge_cells(f'F{row}:G{row}')
    ws.cell(row, 6, '有効期限:').font = f_label
    ws.cell(row, 6).alignment = Alignment(horizontal='right', vertical='center')
    ws.merge_cells(f'H{row}:J{row}')
    ws.cell(row, 8, '発行日より14日間').font = f_normal_s
    ws.cell(row, 8).border = Border(bottom=thin)

    row += 1  # 8 建物情報
    ws.cell(row, 1, '建物情報').font = f_label
    ws.cell(row, 1).fill = fill_gray
    ws.cell(row, 1).border = border_all
    ws.merge_cells(f'A{row}:B{row}')
    ws.cell(row, 2).fill = fill_gray
    ws.cell(row, 2).border = border_all
    ws.merge_cells(f'C{row}:E{row}')
    ws.cell(row, 3, '戸建て ／ 築　　年').font = f_normal_s
    ws.cell(row, 3).border = border_all

    # 右: 事業者情報
    ws.merge_cells(f'F{row}:J{row}')
    ws.cell(row, 6, '配管洗浄専門店 匠 / 代表: 土屋 拓幹').font = f_small
    ws.cell(row, 6).alignment = Alignment(horizontal='right')

    row += 1  # 9
    ws.merge_cells(f'F{row}:J{row}')
    ws.cell(row, 6, '愛知県名古屋市中川区 / TEL: 050-1725-6168').font = f_small
    ws.cell(row, 6).alignment = Alignment(horizontal='right')

    ws.row_dimensions[row].height = 6

    # ═══════════════════════════════════════
    # 現地確認の所見
    # ═══════════════════════════════════════
    row += 1  # 10
    ws.merge_cells(f'A{row}:J{row}')
    ws.cell(row, 1, '  現地確認の所見').font = f_section
    ws.cell(row, 1).fill = fill_blue
    for c in range(1, 11):
        ws.cell(row, c).fill = fill_blue
    ws.row_dimensions[row].height = 18

    findings = [('症状', 20), ('原因（推定）', 20), ('推奨作業', 24)]
    for label, height in findings:
        row += 1
        ws.merge_cells(f'A{row}:B{row}')
        ws.cell(row, 1, label).font = f_label
        ws.cell(row, 1).fill = fill_gray
        ws.cell(row, 1).border = border_all
        ws.cell(row, 1).alignment = Alignment(vertical='top', wrap_text=True)
        ws.cell(row, 2).fill = fill_gray
        ws.cell(row, 2).border = border_all
        ws.merge_cells(f'C{row}:J{row}')
        ws.cell(row, 3).border = border_all
        ws.cell(row, 3).alignment = Alignment(vertical='top', wrap_text=True)
        ws.row_dimensions[row].height = height

    # ═══════════════════════════════════════
    # 作業内容・金額テーブル
    # ═══════════════════════════════════════
    row += 1  # 14 空行
    ws.row_dimensions[row].height = 4

    row += 1  # 15 ヘッダー
    headers = [('A', 'No.'), ('B', ''), ('C', ''), ('D', '作業内容'),
               ('E', ''), ('F', '数量'), ('G', '単価'), ('H', '金額（税込）'), ('I', '区分')]
    ws.merge_cells(f'B{row}:E{row}')
    ws.cell(row, 1, 'No.').font = f_table_header
    ws.cell(row, 2, '作業内容').font = f_table_header
    ws.cell(row, 6, '数量').font = f_table_header
    ws.cell(row, 7, '単価').font = f_table_header
    ws.cell(row, 8, '金額（税込）').font = f_table_header
    ws.cell(row, 9, '区分').font = f_table_header

    for c in range(1, 10):
        ws.cell(row, c).fill = fill_navy
        ws.cell(row, c).border = border_all
        ws.cell(row, c).alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[row].height = 18

    # 7行の作業行
    for i in range(1, 8):
        row += 1
        ws.cell(row, 1, i).font = f_normal_s
        ws.cell(row, 1).alignment = Alignment(horizontal='center', vertical='center')
        ws.merge_cells(f'B{row}:E{row}')
        for c in range(1, 10):
            ws.cell(row, c).border = border_all
            ws.cell(row, c).font = f_normal_s
        ws.cell(row, 6).alignment = Alignment(horizontal='center')
        ws.cell(row, 7).number_format = '#,##0'
        ws.cell(row, 8).number_format = '#,##0'
        ws.cell(row, 9).font = Font(size=7, color='999999')
        ws.cell(row, 9).alignment = Alignment(horizontal='center')
        ws.row_dimensions[row].height = 20

    # 出張費行
    row += 1
    ws.cell(row, 1).fill = fill_gray
    ws.cell(row, 1).border = border_all
    ws.merge_cells(f'B{row}:E{row}')
    ws.cell(row, 2, '出張費（料金に含まれています）').font = f_normal_s
    ws.cell(row, 2).fill = fill_gray
    ws.cell(row, 6, '──').font = f_normal_s
    ws.cell(row, 6).alignment = Alignment(horizontal='center')
    ws.cell(row, 7, '──').font = f_normal_s
    ws.cell(row, 7).alignment = Alignment(horizontal='center')
    ws.cell(row, 8, 0).font = Font(size=9, bold=True, color=DARK_TEXT)
    ws.cell(row, 8).number_format = '#,##0"円"'
    ws.cell(row, 8).alignment = Alignment(horizontal='center')
    for c in range(1, 10):
        ws.cell(row, c).fill = fill_gray
        ws.cell(row, c).border = border_all
    ws.row_dimensions[row].height = 16

    # 小計行
    first_amount_row = row - 7  # 最初の作業行
    last_amount_row = row - 1   # 最後の作業行

    row += 1
    ws.merge_cells(f'A{row}:G{row}')
    ws.cell(row, 1, '小計（税抜）').font = f_bold_9
    ws.cell(row, 1).alignment = Alignment(horizontal='right', vertical='center')
    ws.merge_cells(f'H{row}:I{row}')
    ws.cell(row, 8).font = f_bold_9
    ws.cell(row, 8).number_format = '#,##0"円"'
    ws.cell(row, 8).alignment = Alignment(horizontal='center')
    for c in range(1, 10):
        ws.cell(row, c).border = border_all
    ws.row_dimensions[row].height = 18
    subtotal_row = row

    # 消費税行
    row += 1
    ws.merge_cells(f'A{row}:G{row}')
    ws.cell(row, 1, '消費税（10%）').font = f_normal_s
    ws.cell(row, 1).alignment = Alignment(horizontal='right', vertical='center')
    ws.merge_cells(f'H{row}:I{row}')
    ws.cell(row, 8).font = f_normal_s
    ws.cell(row, 8).number_format = '#,##0"円"'
    ws.cell(row, 8).alignment = Alignment(horizontal='center')
    for c in range(1, 10):
        ws.cell(row, c).border = border_all
    ws.row_dimensions[row].height = 16

    # 合計行
    row += 1
    ws.merge_cells(f'A{row}:G{row}')
    ws.cell(row, 1, '合計金額（税込）').font = f_total_label
    ws.cell(row, 1).fill = fill_navy
    ws.cell(row, 1).alignment = Alignment(horizontal='right', vertical='center')
    for c in range(1, 8):
        ws.cell(row, c).fill = fill_navy
    ws.merge_cells(f'H{row}:I{row}')
    ws.cell(row, 8).font = f_total_amount
    ws.cell(row, 8).fill = fill_cream
    ws.cell(row, 8).number_format = '#,##0"円"'
    ws.cell(row, 8).alignment = Alignment(horizontal='center', vertical='center')
    ws.cell(row, 9).fill = fill_cream
    for c in range(1, 10):
        ws.cell(row, c).border = border_all
    ws.row_dimensions[row].height = 26

    # ═══════════════════════════════════════
    # お支払い方法 + 保証
    # ═══════════════════════════════════════
    row += 1
    ws.row_dimensions[row].height = 4

    row += 1
    ws.merge_cells(f'A{row}:J{row}')
    ws.cell(row, 1, 'お支払い方法:  □ 現金　　□ 銀行振込　　□ クレジットカード　　□ その他（　　　　　）').font = f_normal_s
    ws.row_dimensions[row].height = 16

    row += 1
    ws.merge_cells(f'A{row}:B{row}')
    ws.cell(row, 1, '保証内容').font = f_label
    ws.cell(row, 1).fill = fill_gray
    ws.cell(row, 1).border = border_all
    ws.cell(row, 2).fill = fill_gray
    ws.cell(row, 2).border = border_all
    ws.merge_cells(f'C{row}:J{row}')
    ws.cell(row, 3, '作業完了後　　日以内に同一箇所で同一症状が再発した場合、無償で再対応いたします。').font = f_normal_s
    ws.cell(row, 3).border = border_all
    ws.row_dimensions[row].height = 18

    # ═══════════════════════════════════════
    # 匠のお約束
    # ═══════════════════════════════════════
    row += 1
    ws.row_dimensions[row].height = 4

    row += 1
    ws.merge_cells(f'A{row}:J{row}')
    ws.cell(row, 1, '  匠のお約束').font = f_promise_title
    ws.cell(row, 1).fill = fill_light_blue
    for c in range(1, 11):
        ws.cell(row, c).fill = fill_light_blue
    ws.row_dimensions[row].height = 16

    promises = [
        '1. この見積書の内容にご納得いただいてから作業を開始します。',
        '2. 作業中に追加作業が必要になった場合は、必ず事前にご説明し、ご了承を得てから進めます。',
        '3. 作業後、実施内容と今後の注意点をお伝えします。',
    ]
    for p in promises:
        row += 1
        ws.merge_cells(f'A{row}:J{row}')
        ws.cell(row, 1, f'  {p}').font = f_promise
        ws.row_dimensions[row].height = 13

    # ═══════════════════════════════════════
    # 注意事項
    # ═══════════════════════════════════════
    row += 1
    ws.merge_cells(f'A{row}:J{row}')
    ws.cell(row, 1, '【注意事項】').font = Font(size=7, bold=True, color='808080')
    ws.row_dimensions[row].height = 12

    notes = [
        '・本見積りの有効期限は発行日より14日間です。・お見積り金額は税込表示です。',
        '・作業中に追加対応が必要と判断した場合は、作業を中断し、追加内容と費用をご説明のうえ、お客様のご了承を得てから再開します。',
        '・建物の構造・配管の老朽化等により、作業をお受けできない場合がございます。・キャンセル: 作業着手前は無料です。',
    ]
    for n in notes:
        row += 1
        ws.merge_cells(f'A{row}:J{row}')
        ws.cell(row, 1, f'  {n}').font = f_note
        ws.row_dimensions[row].height = 11

    # ═══════════════════════════════════════
    # お客様確認・署名・印鑑
    # ═══════════════════════════════════════
    row += 1
    ws.row_dimensions[row].height = 4

    row += 1
    ws.merge_cells(f'A{row}:J{row}')
    ws.cell(row, 1, '  お客様ご確認欄').font = f_section
    ws.cell(row, 1).fill = fill_blue
    for c in range(1, 11):
        ws.cell(row, c).fill = fill_blue
    ws.row_dimensions[row].height = 16

    row += 1
    ws.merge_cells(f'A{row}:J{row}')
    ws.cell(row, 1, '上記の見積り内容を確認し、作業を依頼します。').font = f_normal_s
    ws.row_dimensions[row].height = 14

    row += 1
    # 日付
    ws.cell(row, 1, '日付').font = f_label
    ws.cell(row, 1).fill = fill_gray
    ws.cell(row, 1).border = border_all
    ws.cell(row, 1).alignment = Alignment(horizontal='center', vertical='center')
    ws.merge_cells(f'B{row}:C{row}')
    ws.cell(row, 2).border = border_all
    ws.cell(row, 2).font = f_normal_s
    ws.cell(row, 3).border = border_all
    # 署名
    ws.cell(row, 4, 'ご署名').font = f_label
    ws.cell(row, 4).fill = fill_gray
    ws.cell(row, 4).border = border_all
    ws.cell(row, 4).alignment = Alignment(horizontal='center', vertical='center')
    ws.merge_cells(f'E{row}:H{row}')
    ws.cell(row, 5).border = border_all
    for cc in range(5, 9):
        ws.cell(row, cc).border = border_all
    # 印鑑
    ws.cell(row, 9, 'ご印').font = f_seal_label
    ws.cell(row, 9).fill = fill_gray
    ws.cell(row, 9).border = border_all
    ws.cell(row, 9).alignment = Alignment(horizontal='center', vertical='center')
    ws.cell(row, 10).border = border_all
    ws.cell(row, 10).alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[row].height = 32

    # ═══════════════════════════════════════
    # フッター
    # ═══════════════════════════════════════
    row += 1
    ws.row_dimensions[row].height = 4

    row += 1
    # 担当者
    ws.merge_cells(f'A{row}:C{row}')
    ws.cell(row, 1, '配管洗浄専門店 匠  |  TEL: 050-1725-6168').font = f_footer
    ws.cell(row, 4, '担当者:').font = f_label
    ws.cell(row, 4).alignment = Alignment(horizontal='right', vertical='center')
    ws.merge_cells(f'E{row}:H{row}')
    ws.cell(row, 5).border = Border(bottom=thin)
    for cc in range(5, 9):
        ws.cell(row, cc).border = Border(bottom=thin)
    ws.cell(row, 9, '印').font = f_seal_label
    ws.cell(row, 9).alignment = Alignment(horizontal='center', vertical='center')
    ws.cell(row, 10).border = border_all
    ws.cell(row, 10).alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[row].height = 28

    row += 1
    ws.merge_cells(f'A{row}:J{row}')
    ws.cell(row, 1, '※ご不明な点は消費者ホットライン（188）でもご相談いただけます。').font = f_footer

    # ─── 出力 ───
    output_path = '/Users/tsuchiyatakumiki/配管洗浄専門店/docs/estimate-template.xlsx'
    wb.save(output_path)
    print(f'Excel生成完了: {output_path}')
    print('→ Googleドライブにアップロードしてスプレッドシートとして開けます')

if __name__ == '__main__':
    build_xlsx()
