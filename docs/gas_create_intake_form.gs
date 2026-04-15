/**
 * 配管洗浄専門店 匠 - 受付フォーム自動生成スクリプト
 *
 * 【使い方】
 * 1. https://script.google.com を開く
 * 2. 「新しいプロジェクト」を作成
 * 3. このファイルの内容を全てコピー＆ペースト
 * 4. 保存（Ctrl+S または Cmd+S）→ プロジェクト名を「匠 受付フォーム生成」等に変更
 * 5. 上部の関数選択で「createIntakeForm」を選び「実行」をクリック
 * 6. 初回のみGoogle認証を求められるので許可
 * 7. 実行ログに生成されたフォームURL・スプレッドシートURLが表示される
 * 8. フォームURLをスマホのホーム画面にブックマーク → 受付時にワンタップで開ける
 */

function createIntakeForm() {
  // ─── フォーム作成 ───
  const form = FormApp.create('匠 - お客様受付フォーム')
    .setDescription(
      'お電話・LINEでの問い合わせ受付時に入力するフォームです。\n' +
      '※「*」は必須項目です。'
    )
    .setCollectEmail(false)
    .setAllowResponseEdits(true)
    .setShowLinkToRespondAgain(true);

  // ─── 質問を追加 ───

  // Q1: 問合せ経路（必須・ラジオボタン）
  form.addMultipleChoiceItem()
    .setTitle('問合せ経路 *')
    .setChoiceValues(['電話', 'LINE', 'Web（LPからの入力）', '紹介', 'その他'])
    .setRequired(true);

  // Q2: 広告経由か
  form.addMultipleChoiceItem()
    .setTitle('広告経由か')
    .setChoiceValues(['はい（Google広告）', 'いいえ', 'わからない'])
    .setRequired(false);

  // Q3: 検索キーワード（広告経由の場合）
  form.addTextItem()
    .setTitle('検索キーワード（分かれば）')
    .setHelpText('例: 排水 詰まり 名古屋')
    .setRequired(false);

  // Q4: お客様名（必須）
  form.addTextItem()
    .setTitle('お客様名 *')
    .setRequired(true);

  // Q5: 電話番号（必須）
  form.addTextItem()
    .setTitle('電話番号 *')
    .setHelpText('例: 090-1234-5678')
    .setRequired(true);

  // Q6: 住所
  form.addParagraphTextItem()
    .setTitle('ご住所')
    .setHelpText('市区町村から記入。番地は訪問時に確認でもOK')
    .setRequired(false);

  // Q7: 建物種別
  form.addMultipleChoiceItem()
    .setTitle('建物種別')
    .setChoiceValues(['戸建て', 'マンション', 'アパート', 'その他'])
    .setRequired(false);

  // Q8: 築年数
  form.addTextItem()
    .setTitle('築年数（分かれば）')
    .setHelpText('例: 25')
    .setRequired(false);

  // Q9: 症状（必須・段落）
  form.addParagraphTextItem()
    .setTitle('症状・お困りごと *')
    .setHelpText('例: キッチン排水が流れない、ゴボゴボ音がする、悪臭がする等')
    .setRequired(true);

  // Q10: 訪問希望日
  form.addTextItem()
    .setTitle('訪問希望日')
    .setHelpText('例: 本日できるだけ早く / 明日午前 / 4/20午後')
    .setRequired(false);

  // Q11: 担当者
  form.addMultipleChoiceItem()
    .setTitle('受付担当')
    .setChoiceValues(['土屋', 'その他'])
    .setRequired(false);

  // Q12: 備考
  form.addParagraphTextItem()
    .setTitle('備考・申し送り')
    .setHelpText('料金の質問・過去の対応履歴・その他メモ')
    .setRequired(false);

  // ─── 回答用スプレッドシートを作成して紐付け ───
  const ss = SpreadsheetApp.create('匠 - 受付フォーム回答');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  // 回答シートにトリガーを設定（回答送信時に顧客管理シートへ自動転記したい場合）
  // ※ 必要に応じて後から onFormSubmit 関数を有効化可能

  // ─── URL をログに出力 ───
  const formUrl = form.getPublishedUrl();
  const editUrl = form.getEditUrl();
  const ssUrl = ss.getUrl();

  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('✅ フォーム生成完了');
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('📱 受付用URL（これをスマホにブックマーク）:');
  Logger.log(formUrl);
  Logger.log('');
  Logger.log('✏️ フォーム編集URL:');
  Logger.log(editUrl);
  Logger.log('');
  Logger.log('📊 回答が溜まるスプレッドシート:');
  Logger.log(ssUrl);
  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

  return {
    formUrl: formUrl,
    editUrl: editUrl,
    spreadsheetUrl: ssUrl
  };
}


/**
 * 【応用】フォーム回答を「顧客管理.xlsx」シートへ自動転記するトリガー
 *
 * 使い方:
 * 1. 上のcreateIntakeForm()を実行後、生成されたスプレッドシートを開く
 * 2. 拡張機能 → Apps Script でエディタを開く
 * 3. この関数をコピペして保存
 * 4. トリガー → トリガーを追加:
 *    - 関数: onFormSubmit
 *    - イベントのソース: スプレッドシートから
 *    - イベントの種類: フォーム送信時
 * 5. 顧客管理スプレッドシートのIDを CUSTOMER_SHEET_ID に入れる
 */
function onFormSubmit(e) {
  // 顧客管理スプレッドシートのID
  const CUSTOMER_SHEET_ID = '1PDj77WN9jKzzC8aHpfo_FcX71-rz6sYz';
  const SHEET_NAME = '案件管理';

  const values = e.values; // [タイムスタンプ, 問合せ経路, 広告経由, KW, お客様名, 電話, 住所, 建物, 築年, 症状, 訪問希望日, 担当者, 備考]
  const timestamp = new Date(values[0]);

  const ss = SpreadsheetApp.openById(CUSTOMER_SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);

  // 最終行＋1に追記
  const lastRow = sheet.getLastRow();
  const newRow = lastRow + 1;

  // 案件No自動採番（3桁ゼロパディング）
  const caseNo = String(lastRow).padStart(3, '0');

  // 顧客管理シートの列構成に合わせて書き込み
  // A:案件No B:受付日 C:受付時刻 D:経路 E:広告 F:KW G:氏名 H:電話 I:住所
  // J:建物種別 K:築年数 L:症状 M:訪問希望日 ... AF:備考
  const adFlag = values[2] && values[2].indexOf('はい') >= 0 ? 'YES' : (values[2] === 'いいえ' ? 'NO' : '');
  const buildingType = values[7] || '';

  sheet.getRange(newRow, 1).setValue(caseNo);                              // 案件No
  sheet.getRange(newRow, 2).setValue(timestamp);                           // 受付日
  sheet.getRange(newRow, 3).setValue(Utilities.formatDate(timestamp, 'JST', 'HH:mm')); // 受付時刻
  sheet.getRange(newRow, 4).setValue(values[1] || '');                     // 問合せ経路
  sheet.getRange(newRow, 5).setValue(adFlag);                              // 広告経由
  sheet.getRange(newRow, 6).setValue(values[3] || '');                     // KW
  sheet.getRange(newRow, 7).setValue(values[4] || '');                     // お客様名
  sheet.getRange(newRow, 8).setValue(values[5] || '');                     // 電話
  sheet.getRange(newRow, 9).setValue(values[6] || '');                     // 住所
  sheet.getRange(newRow, 10).setValue(buildingType);                       // 建物種別
  sheet.getRange(newRow, 11).setValue(values[8] || '');                    // 築年数
  sheet.getRange(newRow, 12).setValue(values[9] || '');                    // 症状
  sheet.getRange(newRow, 13).setValue(values[10] || '');                   // 訪問希望日
  sheet.getRange(newRow, 16).setValue(values[11] || '');                   // 担当者
  sheet.getRange(newRow, 32).setValue(values[12] || '');                   // 備考

  Logger.log('✅ 顧客管理シートに転記完了: ' + values[4]);
}
