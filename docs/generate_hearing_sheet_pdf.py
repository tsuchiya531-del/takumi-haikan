#!/usr/bin/env python3
"""受電ヒアリングシートをPDFに変換するスクリプト (fpdf2版 - 2ページ版)"""

from fpdf import FPDF

class HearingSheetPDF(FPDF):
    def __init__(self):
        super().__init__()
        self.add_font('Gothic', '', '/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc')
        self.add_font('GothicBold', '', '/System/Library/Fonts/ヒラギノ角ゴシック W6.ttc')
        self.set_auto_page_break(auto=True, margin=8)

    def section_header(self, text):
        self.set_fill_color(26, 42, 74)
        self.set_text_color(255, 255, 255)
        self.set_font('GothicBold', size=8)
        self.cell(0, 5.5, f'  {text}', new_x="LMARGIN", new_y="NEXT", fill=True)
        self.ln(1)
        self.set_text_color(51, 51, 51)

    def sub_header(self, text):
        self.set_font('GothicBold', size=7)
        self.set_text_color(26, 42, 74)
        self.cell(0, 4, text, new_x="LMARGIN", new_y="NEXT")
        self.set_text_color(51, 51, 51)

    def table_row(self, label, value, label_w=32):
        value_w = self.epw - label_w
        self.set_font('Gothic', size=7)
        h = 5.5
        self.set_fill_color(245, 245, 245)
        self.set_draw_color(200, 200, 200)
        self.cell(label_w, h, f' {label}', border=1, fill=True)
        self.set_fill_color(255, 255, 255)
        self.cell(value_w, h, f' {value}', border=1, new_x="LMARGIN", new_y="NEXT")

    def check_item(self, text):
        self.set_font('Gothic', size=7)
        self.cell(0, 4, f' [ ] {text}', new_x="LMARGIN", new_y="NEXT")

    def check_inline(self, items):
        """チェックボックスを横並びで表示"""
        self.set_font('Gothic', size=7)
        text = '  '.join([f'[ ] {item}' for item in items])
        self.cell(0, 4, f' {text}', new_x="LMARGIN", new_y="NEXT")

    def quote_box(self, text):
        self.set_fill_color(240, 244, 248)
        self.set_font('Gothic', size=6.5)
        self.set_text_color(80, 80, 80)
        x = self.get_x() + 2
        w = self.epw - 4
        self.set_x(x)
        self.multi_cell(w, 3.5, text, fill=True)
        self.ln(1)
        self.set_text_color(51, 51, 51)

    def price_table(self):
        self.set_font('GothicBold', size=6.5)
        col_w = [35, 62, 48]
        headers = ['プラン', '内容', '目安（税込・出張費込）']
        self.set_fill_color(44, 95, 138)
        self.set_text_color(255, 255, 255)
        for i, h in enumerate(headers):
            self.cell(col_w[i], 5, f' {h}', border=1, fill=True)
        self.ln()
        self.set_font('Gothic', size=6.5)
        self.set_text_color(51, 51, 51)
        self.set_fill_color(255, 255, 255)
        rows = [
            ['軽度の詰まり対応', '薬剤処理・部分的な詰まり除去', '16,500〜27,500円'],
            ['高圧洗浄（単系統）', '1系統を高圧洗浄', '33,000〜55,000円'],
            ['屋外排水管を含む洗浄', '屋内＋屋外排水マス・配管', '44,000〜66,000円'],
            ['匠コース', '家全体の排水経路一括点検・洗浄', '66,000〜110,000円'],
        ]
        for row in rows:
            for i, val in enumerate(row):
                self.cell(col_w[i], 4.5, f' {val}', border=1, fill=True)
            self.ln()


def build_pdf():
    pdf = HearingSheetPDF()
    pdf.add_page()

    # タイトル
    pdf.set_font('GothicBold', size=13)
    pdf.set_text_color(26, 42, 74)
    pdf.cell(0, 7, '受電ヒアリングシート', align='C', new_x="LMARGIN", new_y="NEXT")
    pdf.set_font('Gothic', size=7)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(0, 4, '配管洗浄専門店 匠 / 電話・LINE共通 / IVRy 050-1725-6168', align='C', new_x="LMARGIN", new_y="NEXT")
    pdf.set_draw_color(26, 42, 74)
    pdf.line(pdf.l_margin, pdf.get_y(), pdf.w - pdf.r_margin, pdf.get_y())
    pdf.ln(2)

    # 受付日時 + お客様情報を横並び
    pdf.section_header('受付日時・お客様情報')
    pdf.table_row('受付日', '　　年　　月　　日（　曜）　　：　　')
    pdf.table_row('チャネル / 担当', '[ ] 電話  [ ] LINE  /  担当：')
    pdf.table_row('お名前（フリガナ）', '')
    pdf.table_row('電話番号', '')
    pdf.table_row('ご住所', '〒　　-　　　　')
    pdf.ln(1)

    # 2. 対応エリア確認
    pdf.section_header('エリア確認')
    pdf.check_inline(['愛知県内→対応可', '岐阜県（一部）', '三重県（一部）', 'エリア外→お断り'])
    pdf.ln(1)

    # 3. 建物情報
    pdf.section_header('建物情報')
    pdf.table_row('建物種別', '[ ] 戸建て  [ ] マンション・集合住宅  [ ] その他（　）')
    pdf.table_row('築年数 / 階数', '約　　年 /　　階建て / 症状の階：　　階')
    pdf.ln(1)

    # 4. 症状の確認
    pdf.section_header('症状の確認')
    pdf.sub_header('発生場所')
    pdf.check_inline(['キッチン', '浴室', '洗面所', '洗濯機排水', 'トイレ', '屋外排水マス', 'その他（　　）'])
    pdf.ln(0.5)

    pdf.sub_header('症状')
    pdf.check_inline(['流れが悪い', '完全に詰まった', 'ボコボコ音', '悪臭', '逆流', '排水マス溢れ', '複数箇所'])
    pdf.check_item('その他（　　　　　　　　　　　）')
    pdf.ln(0.5)

    pdf.sub_header('時期・頻度')
    pdf.table_row('いつから', '[ ] 今日  [ ] 2〜3日前  [ ] 1週間前  [ ] 1ヶ月以上前  [ ] 前からたまに')
    pdf.table_row('頻度', '[ ] 常時  [ ] 毎日  [ ] 週に数回  [ ] たまに')
    pdf.table_row('悪化傾向', '[ ] だんだん悪化  [ ] 変わらない  [ ] 最近急に')
    pdf.ln(0.5)

    pdf.sub_header('お客様が試したこと')
    pdf.check_inline(['パイプクリーナー使用', 'ラバーカップ使用', '他業者に相談済（　　　　）', '何もしていない'])
    pdf.ln(1)

    # 5. 緊急度の判断
    pdf.section_header('緊急度の判断（1つでも該当→【緊急】）')
    pdf.check_inline(['完全に詰まって流せない', '逆流して溢れている', '排水マス溢れ', 'トイレ使用不可', '強い悪臭'])
    pdf.set_font('GothicBold', size=7.5)
    pdf.cell(0, 5, '  ▶ 判定:  [ ] 緊急　　[ ] 通常', new_x="LMARGIN", new_y="NEXT")
    pdf.ln(1)

    # 6. 料金の目安
    pdf.section_header('料金の目安（電話で確定金額は絶対に伝えない）')
    pdf.quote_box('「料金は現地で確認後、作業内容と金額を事前にご説明します。ご納得いただいてから作業開始。見積り無料です。」')
    pdf.price_table()
    pdf.ln(1)
    pdf.quote_box('「高い」と言われた場合: 「金額の不安はもっともです。だからこそ作業前に必ず金額をお伝えし、ご納得いただいてから始めます。」')

    # ===== ページ2 =====

    # 7. 訪問日時
    pdf.section_header('訪問日時の調整')
    pdf.table_row('希望日 第1', '　　月　　日（　曜）　　時頃')
    pdf.table_row('希望日 第2', '　　月　　日（　曜）　　時頃')
    pdf.table_row('確定日時', '　　月　　日（　曜）　　：　　')
    pdf.table_row('立ち会い者', '[ ] ご本人  [ ] ご家族  [ ] その他（　　）')
    pdf.table_row('駐車スペース', '[ ] あり  [ ] なし（近隣コインP等を確認）')
    pdf.table_row('備考', '')
    pdf.ln(1)

    # 8. 電話を切る前の確認
    pdf.section_header('電話を切る前の確認チェックリスト')
    pdf.check_inline(['名前・住所・電話を復唱', '症状と場所を復唱', '訪問日時を復唱'])
    pdf.check_inline(['「見積り無料・見積り後に作業」を伝えた', '「追加費用は事前説明」を伝えた', '駐車スペース確認'])
    pdf.ln(0.5)
    pdf.quote_box('締め: 「それでは○月○日にお伺いします。作業内容と金額をご説明してから始めます。ご不明な点はいつでもお電話ください。」')

    # 9. 受付後の申し送り
    pdf.section_header('受付後の申し送り')
    pdf.table_row('想定プラン', '[ ] 軽度  [ ] 高圧洗浄  [ ] 屋外含む  [ ] 匠コース  [ ] 不明')
    pdf.table_row('特記事項', '')
    pdf.table_row('作業担当（予定）', '')
    pdf.ln(2)

    # LINE問い合わせ用テンプレート
    pdf.section_header('LINE問い合わせ返信テンプレート')
    pdf.quote_box('お問い合わせありがとうございます。配管洗浄専門店 匠です。\n以下を教えていただけますか。\n1. お名前  2. ご住所  3. 症状のある場所  4. どんな症状か  5. いつ頃からか  6. 症状のわかる写真\n現地で確認し、作業内容と金額をご説明してからの作業となります。見積り無料です。')

    # トークスクリプト集
    pdf.section_header('トークスクリプト集')
    pdf.quote_box('【料金を聞かれた場合】「現地で状況を確認し、作業内容と金額を事前にご説明します。ご納得いただいてから作業します。」')
    pdf.quote_box('【エリア外】「申し訳ございません。現在そちらのエリアは対応しておりません。お力になれず恐縮です。」')
    pdf.quote_box('【マンション】「当店は戸建て専門です。集合住宅は管理会社様への相談をおすすめします。」')
    pdf.quote_box('【緊急時】「今お困りの状況、よくわかります。できる限り早く伺えるよう調整します。」')

    # フッター
    pdf.ln(1)
    pdf.set_draw_color(200, 200, 200)
    pdf.line(pdf.l_margin, pdf.get_y(), pdf.w - pdf.r_margin, pdf.get_y())
    pdf.ln(1)
    pdf.set_font('Gothic', size=6)
    pdf.set_text_color(136, 136, 136)
    pdf.cell(0, 3, '匠の憲法: 「見積り後に作業」「追加費用は事前説明」「安さではなく安心で勝つ」— この3つは電話でも必ずお客様に伝えること。', align='C')

    output_path = '/Users/tsuchiyatakumiki/配管洗浄専門店/docs/hearing-sheet.pdf'
    pdf.output(output_path)
    print(f'PDF生成完了: {output_path}')

if __name__ == '__main__':
    build_pdf()
