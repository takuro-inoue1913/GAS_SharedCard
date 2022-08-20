const TAKU_FUMI_SPREAD_SHEET = SpreadsheetApp.openById('1EmOKt3h89vG1ahKSliNoKEGKmgax0VNnmVRK-pa4DmQ');
const SHARED_CARD_MANAGEMENT_SHEET = TAKU_FUMI_SPREAD_SHEET.getSheetByName('共有カード運用管理');

function addCardUseDetail() {
  /** メール検索クエリを作成 */
  const SUBJECT = 'カード利用のお知らせ(本人ご利用分)'; // 利用お知らせメールの件名
  // const ADDRESS = 'rila0327@gmail.com'; // テスト用
  const ADDRESS = 'info@mail.rakuten-card.co.jp'; // お知らせメールの送信元

  /** 検索期間の初めと終わりを昨日と明日にする事で今日のみのMailを検索できる */
  let afterDate = new Date();
  afterDate.setDate(afterDate.getDate() - 1);
  let beforeDate = new Date();
  beforeDate.setDate(beforeDate.getDate() + 1);
  const DATE_AFTER = Utilities.formatDate(afterDate, 'JST', 'yyyy/M/d');
  const DATE_BEFORE = Utilities.formatDate(beforeDate, 'JST', 'yyyy/M/d');

  // const QUERY = 'subject:' + SUBJECT + ' from:' + ADDRESS; // テスト用
  const QUERY = 'subject:' + SUBJECT + ' from:' + ADDRESS + ' after:' + DATE_AFTER + ' before:' + DATE_BEFORE;

  /** メールを検索 */
  threads = GmailApp.search(QUERY);

  /** 該当メールがあった場合 */
  if(threads.length > 0) {
    const alertData = [];
    const msgs = GmailApp.getMessagesForThreads(threads);

    /** テーブルの左端 */
    const TABLE_LEFT_MOST = 1;
    /** テーブルの右端 */
    const TABLE_RIGHT_MOST = 7;

    /**
     * 検索ヒットしたMailを一つずつ処理する
     */
    for(let i=0; i < msgs.length; i++) {
      /** 最終行番号取得 */
      let lastRow = SHARED_CARD_MANAGEMENT_SHEET.getLastRow();
      /** 新規で追加する行番号 */
      let newRow = lastRow + 1

      /** 元となるデータがある範囲 */
      const sourceRange = SHARED_CARD_MANAGEMENT_SHEET.getRange(
        `${getColName(TABLE_LEFT_MOST)}${lastRow}:${getColName(TABLE_RIGHT_MOST)}${lastRow}`
      );

      /** 本文を取得 */
      const plainBody = msgs[i][0].getPlainBody();
      console.log(`メール本文: \n${plainBody}`);

      /** 受信日を取得 */
      const mailDate = msgs[i][0].getDate();

      /** テーブルデータ取得 */
      const tableData = SHARED_CARD_MANAGEMENT_SHEET.getRange(
        `${getColName(TABLE_LEFT_MOST)}6:${getColName(TABLE_RIGHT_MOST)}${lastRow}`
      ).getValues();

      /** 利用先の配列を取得 */
      const useTargets = plainBody.match(/■利用先:.*/g);
      if (useTargets && useTargets.length) {
        useTargets.forEach((val, index) => {
          useTargets[index] = val.replace(/■利用先:|\s/g, '');
        })
      }
      
      /** 明細日付の配列を取得 */
      const histories = plainBody.match(/[0-9]{4}\/[0-9]{2}\/[0-9]{2}/g);
      if (histories && histories.length){
        histories.forEach((history, index) => {
          const [year, mouth, day] = history.split('/')
          histories[index] = new Date(Number(year), Number(mouth - 1), Number(day))
        })
      }

      /** 金額の配列を取得 */
      const prices = plainBody.match(/■利用金額:.*円/g)
      if (prices && prices.length){
        prices.forEach((price, index) => {
          prices[index] = price.replace(/■利用金額:|円|,|\s/g, '')
        })
      }

      /**
       * データ登録処理
       * indexで取るとreturnで弾かれた分ずれるのでデータ挿入成功した分のみcurrentNumでカウントする
       */
      let currentNum = 0;
      if (useTargets && useTargets.length && useTargets[0]) {
        useTargets.forEach((useTarget) => {
          /** 比較用データ生成 */
          const compareData = [
            mailDate ?? new Date(),
            histories[currentNum] ?? new Date(),
            useTargets[currentNum] ?? '', 
            '共有', 
            -Number(prices[currentNum]) ?? 0, 
            '未支払'
          ];

          /** 受信日時、購入品名もしくは金額が一緒の場合は処理をスキップ (重複を防ぐため) */
          if (tableData.find((val => {
            // 受信日時
            return val[0] && 
              formatDate(val[0]) === formatDate(compareData[0]) &&
              // 購入品名
              (val[2] === compareData[2] ||
              // 金額
              val[4] === compareData[4])
            })
          ) !== undefined){
            return;
          }

          /** Slackアラート用のデータ作成 */
          alertData.push(compareData);

          /** オートフィルを反映させたい範囲 */
          const destination = SHARED_CARD_MANAGEMENT_SHEET.getRange(
            `${getColName(TABLE_LEFT_MOST)}${newRow + currentNum}:${getColName(TABLE_RIGHT_MOST)}${newRow + currentNum}`
          );

          /** 元のデータを新規で追加する行にコピーする */
          sourceRange.copyTo(destination);

          console.log(`
            受信日時: ${mailDate}, 
            履歴: ${histories[currentNum]}, 
            購入品名: ${useTargets[currentNum]}, 
            金額: ${-Number(prices[currentNum])}
          `)

          /** 受信日時: メール受信時間を設定 */
          const dateSell = SHARED_CARD_MANAGEMENT_SHEET.getRange(`A${newRow + currentNum}`);
          dateSell.setValue(mailDate ?? new Date());

          /** 履歴: 明細日付を設定 */
          const historySell = SHARED_CARD_MANAGEMENT_SHEET.getRange(`B${newRow + currentNum}`);
          historySell.setValue(histories[currentNum] ?? new Date());

          /** 購入品名: 利用先を設定 */
          const purchaseProductNameSell = SHARED_CARD_MANAGEMENT_SHEET.getRange(`C${newRow + currentNum}`);
          purchaseProductNameSell.setValue(useTargets[currentNum] ?? '');

          /** 支払者: デフォルトは「共有」に設定 */
          const payerSell = SHARED_CARD_MANAGEMENT_SHEET.getRange(`D${newRow + currentNum}`);
          payerSell.setValue('共有');

          /** 金額: 利用金額を負の数で設定 */
          const priceSell = SHARED_CARD_MANAGEMENT_SHEET.getRange(`E${newRow + currentNum}`);
          // 固定費の場合金額は0円にする
          isFixedCost(useTarget) ? priceSell.setValue(0) : priceSell.setValue(-Number(prices[currentNum]) ?? 0);

          /** 支払状況フラグ設定 */
          const paymentStatusSell = SHARED_CARD_MANAGEMENT_SHEET.getRange(`F${newRow + currentNum}`);
          // 固定費の場合支払済にする
          isFixedCost(useTarget) ? paymentStatusSell.setValue('支払済') : paymentStatusSell.setValue('未入金');

          /** 固定費支払金額設定 */
          const fixedCostSell = SHARED_CARD_MANAGEMENT_SHEET.getRange(`G${newRow + currentNum}`);
          // 固定費の場合支払済にする
          isFixedCost(useTarget) ? fixedCostSell.setValue(-Number(prices[currentNum]) ?? '') : fixedCostSell.setValue('');

          currentNum++
        })
      }
    };
    /** Slackへデータ送信 */
    if (alertData.length) {
      slackAlert(alertData);
    }
  }
}

/** スラックへの通知 */
function slackAlert(data) {
  const slackMessage = data.map((val) => `
  ======================================
  利用日: ${Utilities.formatDate(val[1], 'JST', 'yyyy/M/d')}
  購入品名: ${val[2]}
  金額: ${Math.abs(val[4])}円
  ======================================
  `)

  const totalPrice = SHARED_CARD_MANAGEMENT_SHEET.getRange(`H3`).getValue();
  const postUrl = 'https://hooks.slack.com/services/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';
  const username = 'たくふみシート Bot';
  const jsonData ={
    "username" : username,
    "text" : `<@U01AP8MAZNX> <@U01AP8QRE2X>\nスプレッドシートに記入完了しました！📝 支払い状況を更新してください💁‍♀️ \n
  https://docs.google.com/spreadsheets/d/1EmOKt3h89vG1ahKSliNoKEGKmgax0VNnmVRK-pa4DmQ/edit#gid=31098273 \n
  ちなみに今の残り金額は ${totalPrice.toLocaleString()}円です。\n
  ${slackMessage}`
  };
  const payload = JSON.stringify(jsonData);

  const options =
  {
    "method" : "post",
    "contentType" : "application/json",
    "payload" : payload
  };

  UrlFetchApp.fetch(postUrl, options);
}

/** 時間のフォーマット */
function formatDate(date) {
  const yyyy = date.getFullYear(),
  mm = toDoubleDigits(date.getMonth() + 1)
  dd = toDoubleDigits(date.getDate())
  hh = toDoubleDigits(date.getHours())
  mi = toDoubleDigits(date.getMinutes())
  se = toDoubleDigits(date.getSeconds());

  return yyyy + '/' + mm + '/' + dd + ' ' + hh + ':' + mi + ':' + se;
}

/** 日付の0埋め */
function toDoubleDigits(num){
  num += "";
  if (num.length === 1) {
    num = "0" + num;
  }
 return num;     
};

/** 固定費かどうかの判定 (金額に入れたくないものを随時追加する) */
function isFixedCost(useTarget) {
  if (
    /ﾄｳｷﾖｳﾃﾞﾝﾘﾖｸ|ＰｉｎＴ|ﾃﾞｲﾃｲｱｲﾄｰﾝ|東京都水道局|東京ガス/.test(useTarget)
  ){
    return true;
  }

  return false;
};

/** セルの列名取得 */
function getColName(num) {
  let result = SHARED_CARD_MANAGEMENT_SHEET.getRange(1, num);
  result = result.getA1Notation();
  result = result.replace(/\d/,'');
 
  return result;
}