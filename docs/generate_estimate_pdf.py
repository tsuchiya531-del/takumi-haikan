#!/usr/bin/env python3
"""配管洗浄専門店 匠 - お客様用見積書PDF（1ページ版）"""

from fpdf import FPDF
import math

# ─── レイアウト定数（全セクションのy座標を事前決定して重なりを防ぐ） ───
# 各セクションの実高さを計算し、3mm以上の余白を確保して配置
#   テーブル実高: ヘッダー5.5 + 作業7×7(49) + 出張費6 + 小計5.5 + 消費税5.5 + 合計9 = 80.5mm
#   テーブル: 64 → 144.5
#   支払い: 148 → 153 (高さ5)
#   保証:   156 → 161 (高さ5)
#   お約束: 165 → 182 (header5 + 3行×3.8=11.4 + 余裕 → 17mm)
#   注意:   186 → 199 (header3 + 3行×3=9 + 余裕 → 13mm)
#   署名:   204 → 225 (header5 + 文3.5 + gap1.5 + 枠10 = 20mm)
#   フッター: 275
Y_HEADER_END       = 24
Y_INFO_START       = 28
Y_TABLE_START      = 64
Y_PAYMENT_START    = 148
Y_WARRANTY_START   = 156
Y_PROMISE_START    = 165
Y_NOTES_START      = 186
Y_SIGNATURE_START  = 204
Y_FOOTER_START     = 275


class EstimatePDF(FPDF):
    def __init__(self):
        super().__init__()
        self.add_font('Gothic', '', '/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc')
        self.add_font('GothicBold', '', '/System/Library/Fonts/ヒラギノ角ゴシック W6.ttc')
        self.set_auto_page_break(auto=False)

    def draw_circle(self, cx, cy, r, color=(180, 180, 180)):
        """印鑑用の丸枠を描画"""
        self.set_draw_color(*color)
        pts = []
        for i in range(36):
            angle = 2 * math.pi * i / 36
            pts.append((cx + r * math.cos(angle), cy + r * math.sin(angle)))
        for i in range(len(pts)):
            x1, y1 = pts[i]
            x2, y2 = pts[(i + 1) % len(pts)]
            self.line(x1, y1, x2, y2)
        self.set_draw_color(200, 200, 200)

    # ═══ ヘッダー ═══
    def header_block(self):
        self.set_fill_color(26, 42, 74)
        self.rect(0, 0, 210, Y_HEADER_END, 'F')

        self.set_text_color(255, 255, 255)
        self.set_font('Gothic', size=7)
        self.set_xy(12, 4)
        self.cell(80, 3.5, '配管洗浄・詰まり抜き専門店')

        self.set_font('GothicBold', size=18)
        self.set_xy(12, 8)
        self.cell(80, 10, '匠')

        self.set_font('Gothic', size=6.5)
        self.set_xy(12, 18.5)
        self.cell(80, 3.5, 'TEL: 050-1725-6168  |  takumi-haikan.com')

        self.set_font('GothicBold', size=16)
        self.set_xy(130, 7)
        self.cell(68, 10, '御 見 積 書', align='R')

        self.set_text_color(51, 51, 51)

    # ═══ 見積番号 + お客様情報 + 事業者情報 ═══
    def info_section(self):
        y = Y_INFO_START
        self.set_draw_color(200, 200, 200)

        # 右：見積番号ブロック
        self.set_font('Gothic', size=7)
        self.set_xy(130, y)
        self.cell(26, 4.5, '見積番号:')
        self.cell(42, 4.5, '', border='B')

        self.set_xy(130, y + 5)
        self.cell(26, 4.5, '発行日:')
        self.cell(42, 4.5, '令和　　年　　月　　日', border='B')

        self.set_xy(130, y + 10)
        self.cell(26, 4.5, '有効期限:')
        self.cell(42, 4.5, '発行日より14日間', border='B')

        # 左：お客様名（下線）
        self.set_font('GothicBold', size=10)
        self.set_xy(10, y)
        self.set_draw_color(26, 42, 74)
        self.cell(110, 7, '', border='B')
        self.set_xy(10, y)
        self.cell(80, 7, '様', align='R')

        # お客様情報3行
        self.set_draw_color(200, 200, 200)
        self.set_font('Gothic', size=7)
        labels = ['ご住所', 'お電話', '建物情報']
        for i, label in enumerate(labels):
            self.set_xy(10, y + 9 + i * 5.5)
            self.set_fill_color(245, 245, 245)
            self.cell(18, 5, f' {label}', border=1, fill=True)
            val = '戸建て ／ 築　　年' if label == '建物情報' else ''
            self.cell(92, 5, f' {val}', border=1)

        # 右下：事業者情報（お客様情報と並列・ヘッダーに被らない）
        biz_y = y + 17
        self.set_font('Gothic', size=6)
        self.set_text_color(100, 100, 100)
        self.set_xy(130, biz_y)
        self.cell(70, 3.5, '配管洗浄・詰まり抜き専門店 匠')
        self.set_xy(130, biz_y + 3.5)
        self.cell(70, 3.5, '代表: 土屋 卓幹')
        self.set_xy(130, biz_y + 7)
        self.cell(70, 3.5, '愛知県名古屋市中川区')
        self.set_xy(130, biz_y + 10.5)
        self.cell(70, 3.5, '富田町榎津東新海1430-48')
        self.set_text_color(51, 51, 51)

    # ═══ 作業テーブル ═══
    def estimate_table(self):
        y = Y_TABLE_START
        self.set_xy(10, y)
        self.set_fill_color(26, 42, 74)
        self.set_text_color(255, 255, 255)
        self.set_font('GothicBold', size=7)

        col_w = [10, 72, 18, 22, 22, 16]
        headers = ['No.', '作業内容', '数量', '単価', '金額', '区分']
        for i, h in enumerate(headers):
            self.cell(col_w[i], 5.5, f' {h}', border=1, fill=True)
        self.ln()
        self.set_text_color(51, 51, 51)

        # 7行
        self.set_font('Gothic', size=7)
        for row_num in range(1, 8):
            self.set_x(10)
            self.cell(col_w[0], 7, f' {row_num}', border=1)
            self.cell(col_w[1], 7, '', border=1)
            self.cell(col_w[2], 7, '', border=1)
            self.cell(col_w[3], 7, '', border=1)
            self.cell(col_w[4], 7, '', border=1)
            self.set_font('Gothic', size=5.5)
            self.set_text_color(150, 150, 150)
            self.cell(col_w[5], 7, ' 技術/材料', border=1)
            self.set_font('Gothic', size=7)
            self.set_text_color(51, 51, 51)
            self.ln()

        # 出張費行
        self.set_x(10)
        self.set_fill_color(245, 245, 245)
        self.cell(col_w[0], 6, '', border=1, fill=True)
        self.set_font('Gothic', size=6.5)
        self.cell(col_w[1], 6, ' 出張費（料金に含まれています）', border=1, fill=True)
        self.cell(col_w[2], 6, ' ──', border=1, fill=True, align='C')
        self.cell(col_w[3], 6, ' ──', border=1, fill=True, align='C')
        self.set_font('GothicBold', size=7)
        self.cell(col_w[4], 6, ' 0円', border=1, fill=True)
        self.cell(col_w[5], 6, '', border=1, fill=True)
        self.ln()

        # 小計・消費税・合計
        left_w = col_w[0] + col_w[1] + col_w[2] + col_w[3]
        right_w = col_w[4] + col_w[5]

        self.set_x(10)
        self.set_font('Gothic', size=7)
        self.cell(left_w, 5.5, '小計（税抜）　', border=1, align='R')
        self.cell(right_w, 5.5, '　　　　　　　　　円', border=1)
        self.ln()

        self.set_x(10)
        self.cell(left_w, 5.5, '消費税（10%）　', border=1, align='R')
        self.cell(right_w, 5.5, '　　　　　　　　　円', border=1)
        self.ln()

        # 合計
        self.set_x(10)
        self.set_fill_color(26, 42, 74)
        self.set_text_color(255, 255, 255)
        self.set_font('GothicBold', size=9)
        self.cell(left_w, 9, '合計金額（税込）　', border=1, fill=True, align='R')
        self.set_fill_color(255, 245, 230)
        self.set_text_color(26, 42, 74)
        self.set_font('GothicBold', size=11)
        self.cell(right_w, 9, '　　　　　　　円', border=1, fill=True)
        self.ln()
        self.set_text_color(51, 51, 51)

    # ═══ お支払い方法 ═══
    def payment_section(self):
        y = Y_PAYMENT_START
        self.set_xy(10, y)
        self.set_font('Gothic', size=6.5)
        self.cell(0, 4, 'お支払い方法:  □ 現金　　□ 銀行振込　　□ クレジットカード　　□ その他（　　　　　）')

    # ═══ 保証内容 ═══
    def warranty_section(self):
        y = Y_WARRANTY_START
        self.set_xy(10, y)
        self.set_draw_color(200, 200, 200)
        self.set_fill_color(245, 245, 245)
        self.set_font('Gothic', size=6.5)
        self.cell(18, 5, ' 保証内容', border=1, fill=True)
        self.cell(self.epw - 18, 5, ' 作業完了後　　日以内に同一箇所で同一症状が再発した場合、無償で再対応いたします。', border=1)

    # ═══ 匠のお約束 ═══
    def promise_section(self):
        y = Y_PROMISE_START
        self.set_xy(10, y)
        self.set_fill_color(240, 244, 248)
        self.set_font('GothicBold', size=7)
        self.set_text_color(26, 42, 74)
        self.cell(self.epw, 5, '  匠のお約束', fill=True)

        self.set_font('Gothic', size=6)
        self.set_text_color(51, 51, 51)
        promises = [
            '1. この見積書の内容にご納得いただいてから作業を開始します。',
            '2. 作業中に追加作業が必要になった場合は、必ず事前にご説明し、ご了承を得てから進めます。',
            '3. 作業後、実施内容と今後の注意点をお伝えします。',
        ]
        # 各お約束行は高さ3.8で描画、ヘッダー5 + 3.8*3 = 16.4
        for i, p in enumerate(promises):
            self.set_xy(12, y + 5.5 + i * 3.8)
            self.cell(0, 3.5, p)

    # ═══ 注意事項 ═══
    def notes_section(self):
        y = Y_NOTES_START
        self.set_xy(10, y)
        self.set_font('GothicBold', size=6)
        self.set_text_color(100, 100, 100)
        self.cell(0, 3, '【注意事項】')

        self.set_font('Gothic', size=5.5)
        notes = [
            '・本見積りの有効期限は発行日より14日間です。お見積り金額は税込表示です。',
            '・作業中に追加対応が必要と判断した場合は、作業を中断し、追加内容と費用をご説明のうえ、お客様のご了承を得てから再開します。',
            '・建物の構造・配管の老朽化等により、作業をお受けできない場合がございます。キャンセルは作業着手前まで無料です。',
        ]
        for i, n in enumerate(notes):
            self.set_xy(10, y + 4 + i * 3)
            self.cell(0, 2.8, n)
        self.set_text_color(51, 51, 51)

    # ═══ お客様ご確認欄（署名＋印鑑） ═══
    def signature_section(self):
        y = Y_SIGNATURE_START
        self.set_xy(10, y)
        self.set_fill_color(44, 95, 138)
        self.set_text_color(255, 255, 255)
        self.set_font('GothicBold', size=7)
        self.cell(self.epw, 5, '  お客様ご確認欄', fill=True)

        self.set_text_color(51, 51, 51)
        self.set_font('Gothic', size=6.5)
        self.set_xy(10, y + 6)
        self.cell(0, 3.5, '上記の見積り内容を確認し、作業を依頼します。')

        # 署名行
        self.set_draw_color(200, 200, 200)
        self.set_fill_color(245, 245, 245)
        self.set_font('Gothic', size=6.5)
        sy = y + 11
        self.set_xy(10, sy)
        self.cell(12, 10, ' 日付', border=1, fill=True)
        self.cell(44, 10, '　　　年　　月　　日', border=1)
        self.cell(14, 10, ' ご署名', border=1, fill=True)
        self.cell(66, 10, '', border=1)
        self.cell(12, 10, ' ご印', border=1, fill=True)
        seal_x = self.get_x()
        self.cell(12, 10, '', border=1)
        self.draw_circle(seal_x + 6, sy + 5, 4)

    # ═══ フッター ═══
    def footer_block(self):
        y = Y_FOOTER_START
        self.set_draw_color(26, 42, 74)
        self.line(10, y, 200, y)

        # 左：連絡先
        self.set_font('Gothic', size=6)
        self.set_text_color(100, 100, 100)
        self.set_xy(10, y + 2)
        self.cell(110, 3.5, '配管洗浄・詰まり抜き専門店 匠  |  TEL: 050-1725-6168  |  LINE: lin.ee/zw0OS3N')
        self.set_xy(10, y + 5.5)
        self.cell(110, 3.5, '※ご不明な点は消費者ホットライン（188）でもご相談いただけます。')

        # 右：担当者署名 + 印鑑
        self.set_xy(130, y + 2)
        self.set_font('Gothic', size=6.5)
        self.set_text_color(51, 51, 51)
        self.cell(16, 7, '担当者:', border=0)
        self.set_draw_color(200, 200, 200)
        self.cell(42, 7, '', border='B')
        self.cell(4, 7, '', border=0)
        seal_x = self.get_x()
        self.cell(8, 7, '', border=1)
        self.draw_circle(seal_x + 4, y + 5.5, 3)


def build_pdf():
    pdf = EstimatePDF()
    pdf.add_page()

    pdf.header_block()
    pdf.info_section()
    pdf.estimate_table()
    pdf.payment_section()
    pdf.warranty_section()
    pdf.promise_section()
    pdf.notes_section()
    pdf.signature_section()
    pdf.footer_block()

    output_path = '/Users/tsuchiyatakumiki/配管洗浄専門店/docs/estimate-template.pdf'
    pdf.output(output_path)
    print(f'PDF生成完了: {output_path}')


if __name__ == '__main__':
    build_pdf()
