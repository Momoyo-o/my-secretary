//メイン処理
function main() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('スケジュール');
  const logSheet = ss.getSheetByName('ログ');
  const data = sheet.getDataRange().getValues();

  const now = new Date();
  const dateStr = Utilities.formatDate(now, 'JST', 'M月d日');
  const days = ['日曜日', '月曜日', '火曜日', '水曜日', '木曜日', '金曜日', '土曜日'];
  const currentDay = days[now.getDay()];
  const currentMonth = now.getMonth() + 1; // 🆕 花粉情報用

  const scriptProperties = PropertiesService.getScriptProperties();
  const LINE_TOKEN = scriptProperties.getProperty('LINE_ACCESS_TOKEN');
  const USER_ID = scriptProperties.getProperty('USER_ID');

  //2行目（i=1）からスタート
  for (let i = 1; i < data.length; i++) {
    // [0]有効, [1]地域コード, [2]ニュースカテゴリ, [3]路線名, [4]備考
    const [active, cityCode, newsCategory, routeName, memo] = data[i];

    if (active === true) {
      try {
        // 1. ユーザー名取得
        const userName = getUserDisplayName(LINE_TOKEN, USER_ID);

        // 2. 天気情報取得（構造化データ）
        const weatherData = getWeatherData(cityCode || "130000");

        // 3. 天気アラート生成
        const weatherAlert = generateWeatherAlert(weatherData);

        // 4. ニュース取得（カテゴリ指定）
        const newsList = getNews(newsCategory || "一般");

        // 5. 鉄道運行情報取得
        const trainInfo = getTrainInfo(routeName);

        // 6. 花粉情報取得（2-5月のみ）
        const pollenInfo = getPollenInfo(cityCode || "130000", currentMonth);

        // 名言
        const meigen = getRandomMeigen();

        // 7. メッセージ組み立て
        const finalMessage = buildMessage({
          userName,
          date: dateStr,
          day: currentDay,
          weatherData,
          weatherAlert,
          newsList,
          newsCategory: newsCategory || "一般",
          trainInfo,
          routeName,
          pollenInfo,
          meigen: meigen
        });

        // 8. 送信
        const result = sendLineMessage(LINE_TOKEN, USER_ID, finalMessage);
        

        // ログに表示する
        console.log("--- 送信内容の確認 ---");
        console.log(finalMessage);
        console.log("----------------------");

        // 9. （デバック中のみ）ログ記録用のダミー結果
        //const result = { status: "デバッグ中" };

        // 9. ログ記録
        logSheet.appendRow([
          new Date(),
          memo,
          newsCategory || "一般",
          routeName || "",
          result.status,
          weatherAlert ? "有" : "無",
          trainInfo ? "有" : "無"
        ]);

      } catch (error) {
        Logger.log("スケジュール" + i + "でエラー: " + error.message);
        logSheet.appendRow([new Date(), memo, "", "", "エラー", "", ""]);
      }
    }
  }
}

//天気データ取得
function getWeatherData(cityCode) {
  try {
    const code = String(cityCode).padStart(6, '0');
    const response = UrlFetchApp.fetch(
      "https://www.jma.go.jp/bosai/forecast/data/forecast/" + code + ".json"
    );
    const json = JSON.parse(response.getContentText());

    // 天気概況
    const weather = json[0].timeSeries[0].areas[0].weathers[0];

    // 気温
    let maxTemp = null, minTemp = null;
    try {
      const temps = json[0].timeSeries[2].areas[0].temps;
      // 数字に変換して配列にする
      const tempArray = temps.map(t => parseInt(t));
      
      // 配列の中で一番大きいのが最高、一番小さいのが最低
      maxTemp = Math.max(...tempArray);
      minTemp = Math.min(...tempArray);

      // 万が一、データが1つしかなかった時のための安全策
      if (tempArray.length === 1) {
        maxTemp = tempArray[0];
        minTemp = "---"; // または前日のデータを引き継ぐ処理
      }
      } catch(e) {
      Logger.log("気温データなし");
    }

    // 降水確率
    let precipitation = 0;
    try {
      const pops = json[0].timeSeries[1].areas[0].pops;
      // 午前中の降水確率を取得（pops[0]または[1]）
      precipitation = parseInt(pops[1]) || parseInt(pops[0]) || 0;
    } catch(e) {
      Logger.log("降水確率データなし");
    }

    return {
      weather: weather.replace(/\s+/g, " "),
      maxTemp,
      minTemp,
      precipitation,
      rawText: weather
    };

  } catch (e) {
    Logger.log("天気取得エラー: " + e.message);
    return null;
  }
}

//天気アラート生成
function generateWeatherAlert(weatherData) {
  if (!weatherData) return "";

  const alerts = [];

  // 降水確率チェック
  if (weatherData.precipitation >= 50) {
    alerts.push("☂️ 傘を忘れずに！（降水確率" + weatherData.precipitation + "%）");
  } else if (weatherData.precipitation >= 30) {
    alerts.push("☁️ 傘があると安心です（降水確率" + weatherData.precipitation + "%）");
  }

  // 高温チェック
  if (weatherData.maxTemp >= 30) {
    alerts.push("🌡️ 熱中症に注意！こまめに水分補給を");
  } else if (weatherData.maxTemp >= 25) {
    alerts.push("🌞 暑くなりそうです");
  }

  // 低温チェック
  if (weatherData.minTemp <= 5) {
    alerts.push("🧥 しっかり防寒してください（最低気温" + weatherData.minTemp + "度）");
  } else if (weatherData.minTemp <= 10) {
    alerts.push("🍃 朝晩は冷えます。上着があると安心");
  }

  // 特殊キーワードチェック
  const keywords = ["大雨", "暴風", "雪", "警報", "注意報", "雷"];
  for (const keyword of keywords) {
    if (weatherData.rawText.includes(keyword)) {
      alerts.push("⚠️ " + keyword + "に注意してください");
      break;
    }
  }

  if (alerts.length > 0) {
    return "\n⚡アラート⚡\n" + alerts.join("\n");
  }
  return "";
}

// ニュース取得（カテゴリ対応）
function getNews(category) {
  const NEWS_URLS = {
    "一般": "https://news.google.com/rss?hl=ja&gl=JP&ceid=JP:ja",
    "テクノロジー": "https://news.google.com/rss/topics/CAAqJggKIiBDQkFTRWdvSUwyMHZNRGRqTVhZU0FtcGhHZ0pLVUNnQVAB?hl=ja&gl=JP&ceid=JP:ja",
    "ビジネス": "https://news.google.com/rss/topics/CAAqJggKIiBDQkFTRWdvSUwyMHZNRGx6TVdZU0FtcGhHZ0pLVUNnQVAB?hl=ja&gl=JP&ceid=JP:ja",
    "スポーツ": "https://news.google.com/rss/topics/CAAqJggKIiBDQkFTRWdvSUwyMHZNRFp1ZEdvU0FtcGhHZ0pLVUNnQVAB?hl=ja&gl=JP&ceid=JP:ja",
    "エンタメ": "https://news.google.com/rss/topics/CAAqJggKIiBDQkFTRWdvSUwyMHZNREpxYW5RU0FtcGhHZ0pLVUNnQVAB?hl=ja&gl=JP&ceid=JP:ja"
  };

  const url = NEWS_URLS[category] || NEWS_URLS["一般"];

  try {
    const response = UrlFetchApp.fetch(url);
    const xml = response.getContentText();
    const items = xml.split('<item>');
    const newsList = [];

    for (let i = 1; i <= 3; i++) {
      if (items[i]) {
        let title = items[i].split('<title>')[1].split('</title>')[0];
        title = title.split(' - ')[0]; // 配信元を除去
        newsList.push(i + ". " + title);
      }
    }

    return newsList.join("\n");
  } catch (e) {
    Logger.log("ニュース取得エラー: " + e.message);
    return "ニュースの取得に失敗しました";
  }
}

//鉄道運行情報取得
function getTrainInfo(routeName) {
  if (!routeName || routeName.trim() === "") {
    return null;
  }

  // マスタシートからリストを読み込む処理
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName("路線マスタ");
  const masterData = masterSheet.getDataRange().getValues();
  
  // シートのデータを変換
  const routeMap = {};
  for (let i = 1; i < masterData.length; i++) { // 1行目は見出しなので飛ばす
    const name = masterData[i][0]; // A列: 路線名
    const code = masterData[i][1]; // B列: コード
    routeMap[name] = code.toString();
  }

  const routeCode = routeMap[routeName];
  if (!routeCode) {
    Logger.log("路線マスタに未登録: " + routeName);
    return {
      status: "未対応",
      detail: "この路線はまだ対応していません"
    };
  }

  try {
    const url = "https://transit.yahoo.co.jp/diainfo/" + routeCode + "/0";
    const response = UrlFetchApp.fetch(url);
    const html = response.getContentText();

    // 1. 平常運転の場合
    if (html.includes("icnNormalLarge") || html.includes("平常運転")) {
      return {
        status: "平常運転",
        detail: "現在、事故・遅延に関する情報はありません。"
      };
    }

    // 2. 異常がある場合、その理由（テキスト）を抜き出す
    // <dd class="trouble"> または <dd class="normal"> の中の <p>タグの中身を取得
    const statusMatch = html.match(/<dd class="(?:trouble|normal)">\s*<p>(.*?)<\/p>/);
    
    if (statusMatch) {
      const statusText = statusMatch[1].replace(/<[^>]+>/g, "").trim();
      return {
        status: "⚠️ 運行情報あり",
        detail: statusText
      };
    }

    return null;

  } catch (e) {
    Logger.log("解析エラー: " + e.message);
    return { status: "情報取得エラー", detail: "路線の状態を確認できませんでした" };
  }
}

//花粉情報取得
function getPollenInfo(cityCode, currentMonth) {
  // 花粉シーズン（2-5月）以外はスキップ
  if (currentMonth < 2 || currentMonth > 5) {
    return null;
  }

  try {
    // 都道府県コード（地域コードの最初の2桁）
    const prefCode = String(cityCode).substring(0, 2);
    const url = "https://tenki.jp/pollen/" + prefCode + "/";

    const response = UrlFetchApp.fetch(url);
    const html = response.getContentText();

    // 今日の花粉レベルを抽出
    const levelMatch = html.match(/今日の花粉.+?level-(\d)/s);
    if (!levelMatch) {
      return null;
    }

    const level = parseInt(levelMatch[1]);
    const levelTexts = ["", "少ない", "やや多い", "多い", "非常に多い"];
    const levelEmojis = ["", "😊", "😐", "😷", "🤧"];

    let message = levelEmojis[level] + " 花粉：" + levelTexts[level];

    if (level >= 3) {
      message += "（マスク・メガネの着用をおすすめします）";
    }

    return message;

  } catch (e) {
    Logger.log("花粉情報取得エラー: " + e.message);
    return null;
  }
}

// 名言取得
function getRandomMeigen() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("名言");
    const data = sheet.getDataRange().getValues();
    
    // データが1行（見出しのみ）しかない場合は終了
    if (data.length <= 1) return null;

    // 2行目以降からランダムに1行選ぶ
    const randomIndex = Math.floor(Math.random() * (data.length - 1)) + 1;
    return data[randomIndex][0]; // A列の言葉を返す
  } catch (e) {
    console.log("名言取得エラー: " + e.message);
    return null;
  }
}

// LINEの登録名を取得する関数
function getUserDisplayName(token, userId) {
  try {
    const url = 'https://api.line.me/v2/bot/profile/' + userId;
    const options = {
      'method': 'get',
      'headers': {
        'Authorization': 'Bearer ' + token
      }
    };
    const response = UrlFetchApp.fetch(url, options);
    const resJson = JSON.parse(response.getContentText());
    return resJson.displayName; // これがLINEの登録名です
  } catch (e) {
    console.log("名前取得エラー: " + e.message);
    return "ユーザー"; // 失敗したときのバックアップ
  }
}

//メッセージ
function sendLineMessage(token, userId, message) {
  const url = 'https://api.line.me/v2/bot/message/push';
  const payload = {
    'to': userId,
    'messages': [{ 'type': 'text', 'text': message }]
  };
  
  const options = {
    'method': 'post',
    'headers': {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + token
    },
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  const resCode = response.getResponseCode();
  
  return {
    status: resCode === 200 ? '成功' : '失敗',
    message: response.getContentText()
  };
}

//各情報を結合してLINE用のメッセージを作成する
function buildMessage(data) {
  let message = data.userName + "さん、おはようございます！\n";
  message += "今日は" + data.date + "(" + data.day + ")です。\n\n";

  // 天気情報
  message += "【今日の天気】\n";
  if (data.weatherData) {
    message += data.weatherData.weather;
    if (data.weatherData.maxTemp && data.weatherData.minTemp) {
      message += "\n（気温：最高" + data.weatherData.maxTemp + "度";
      message += " / 最低" + data.weatherData.minTemp + "度）";
    }
    message += "\n降水確率：" + data.weatherData.precipitation + "%";
  } else {
    message += "天気情報を取得できませんでした";
  }

  // 天気アラート
  if (data.weatherAlert) {
    message += "\n" + data.weatherAlert;
  }
  message += "\n\n";

  // 花粉情報（該当月のみ）
  if (data.pollenInfo) {
    message += "【花粉情報】\n";
    message += data.pollenInfo + "\n\n";
  }

  // 鉄道運行情報
  if (data.trainInfo) {
    message += "【運行情報】\n";
    message += "🚃 " + data.routeName + "\n";
    message += data.trainInfo.status + "\n";
    if (data.trainInfo.detail) {
      message += data.trainInfo.detail + "\n";
    }
    message += "\n";
  }

  // ニュース情報
  message += "【最新ニュース";
  if (data.newsCategory !== "一般") {
    message += "（" + data.newsCategory + "）";
  }
  message += "】\n";
  message += data.newsList + "\n\n";

  //　名言
  if (data.meigen) {
    message += "📜 今日の言葉\n「" + data.meigen + "」\n\n";
  }

  message += "今日も一日頑張りましょう！";

  return message;
}
