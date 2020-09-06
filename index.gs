const _columnInfo = {
  // シート内の各要素が何列目に位置しているか
  num: 1,
  assignee: 2,
  content: 3,
  done: 4,
  deadline: 5,
};

function doPost(e) {
  const verificationToken = e.parameter.token;
  if (verificationToken != "XXXXXXXXXXXXXXXXXXX") { // Slack slach command の Token
    throw new Error("Invalid token");
  }

  // whitespace でパースして処理を分ける
  const text = e.parameter.text;
  const parsedText = text.split(" ");
  const command = e.parameter.command;
  const channelId = e.parameter.channel_id;

  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();

  if (command === "/task_show") {
    const acceptableArgumentLength = Object.keys(_columnInfo).length - 1;
    if (parsedText.length > acceptableArgumentLength) {
      return ContentService.createTextOutput(
        JSON.stringify({
          response_type: "in_channel",
          text: `${command} の引数が多すぎます（許容数：${acceptableArgumentLength}）`,
        })
      ).setMimeType(ContentService.MimeType.JSON);
    }
    const showUndoneRes = showUndone(spreadSheet, channelId);
    return ContentService.createTextOutput(
      JSON.stringify({ response_type: "in_channel", text: showUndoneRes })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  if (command === "/task_show_all") {
    const acceptableArgumentLength = Object.keys(_columnInfo).length - 1;
    if (parsedText.length > acceptableArgumentLength) {
      return ContentService.createTextOutput(
        JSON.stringify({
          response_type: "in_channel",
          text: `${command} の引数が多すぎます（許容数：${acceptableArgumentLength}）`,
        })
      ).setMimeType(ContentService.MimeType.JSON);
    }
    const showAllRes = showAll(spreadSheet, channelId);
    return ContentService.createTextOutput(
      JSON.stringify({ response_type: "in_channel", text: showAllRes })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  if (command === "/task_show_stealth") {
    const acceptableArgumentLength = Object.keys(_columnInfo).length - 1;
    if (parsedText.length > acceptableArgumentLength) {
      return ContentService.createTextOutput(
        JSON.stringify({
          response_type: "in_channel",
          text: `${command} の引数が多すぎます（許容数：${acceptableArgumentLength}）`,
        })
      ).setMimeType(ContentService.MimeType.JSON);
    }
    const showAllRes = showAll(spreadSheet, channelId);
    return ContentService.createTextOutput(
      JSON.stringify({ response_type: "ephemeral", text: showAllRes })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  if (command === "/task_add") {
    const addRes = add(spreadSheet, channelId, parsedText);
    return ContentService.createTextOutput(
      JSON.stringify({ response_type: "in_channel", text: addRes })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  if (command === "/task_done") {
    const acceptableArgumentLength = 1;
    if (parsedText.length > acceptableArgumentLength) {
      return ContentService.createTextOutput(
        JSON.stringify({
          response_type: "in_channel",
          text: `${command} の引数が多すぎます（許容数：${acceptableArgumentLength}）`,
        })
      ).setMimeType(ContentService.MimeType.JSON);
    }
    const doneRes = done(spreadSheet, channelId, parsedText[0]);
    return ContentService.createTextOutput(
      JSON.stringify({ response_type: "in_channel", text: doneRes })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  if (command === "/task_assign") {
    const acceptableArgumentLength = 2;
    if (parsedText.length > acceptableArgumentLength) {
      return ContentService.createTextOutput(
        JSON.stringify({
          response_type: "in_channel",
          text: `${command} の引数が多すぎます（許容数：${acceptableArgumentLength}）`,
        })
      ).setMimeType(ContentService.MimeType.JSON);
    }
    const assignRes = assign(
      spreadSheet,
      channelId,
      parsedText[0],
      parsedText[1]
    );
    return ContentService.createTextOutput(
      JSON.stringify({ response_type: "in_channel", text: assignRes })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  if (command === "/task_set_deadline") {
    const acceptableArgumentLength = 2;
    if (parsedText.length > acceptableArgumentLength) {
      return ContentService.createTextOutput(
        JSON.stringify({
          response_type: "in_channel",
          text: `${command} の引数が多すぎます（許容数：${acceptableArgumentLength}）`,
        })
      ).setMimeType(ContentService.MimeType.JSON);
    }
    const setDeadlineRes = setDeadline(
      spreadSheet,
      channelId,
      parsedText[0],
      parsedText[1]
    );
    return ContentService.createTextOutput(
      JSON.stringify({ response_type: "in_channel", text: setDeadlineRes })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(
    JSON.stringify({ response_type: "in_channel", text: command })
  ).setMimeType(ContentService.MimeType.JSON);
}

function showAll(spreadSheet, channelId, assignee = null) {
  const sheet = spreadSheet.getSheetByName(channelId);
  if (!sheet) {
    return "タスクはありません";
  }
  if (assignee === null || assignee === "") {
    const values = getAllValues(sheet);
    if (values === []) {
      return "タスクはありません";
    }
    let message = "現在のタスクはこちらです :male-technologist::skin-tone-2:\n";
    for (row of values) {
      const [num, assignee, content, done, deadline] = [
        row[_columnInfo["num"] - 1],
        row[_columnInfo["assignee"] - 1],
        row[_columnInfo["content"] - 1],
        row[_columnInfo["done"] - 1],
        row[_columnInfo["deadline"] - 1] !== ""
          ? ` ${row[_columnInfo["deadline"] - 1]}`
          : "期限なし",
      ];
      const statusIcon = getStatusIcon(done, deadline);
      const strikeoutLine = done === true ? "~" : "";
      message += `${strikeoutLine}${statusIcon}#${num} ${assignee} ${content} ${deadline}${strikeoutLine}\n`;
    }
    return message;
  }
}

function showUndone(spreadSheet, channelId, assignee = null) {
  const sheet = spreadSheet.getSheetByName(channelId);
  if (!sheet) {
    return "タスクはありません";
  }
  if (assignee === null || assignee === "") {
    const values = getAllValues(sheet);
    if (values === []) {
      return "タスクはありません";
    }
    let message =
      "現在の未完タスクはこちらです :male-technologist::skin-tone-2:\n";
    for (row of values) {
      const [num, assignee, content, done, deadline] = [
        row[_columnInfo["num"] - 1],
        row[_columnInfo["assignee"] - 1],
        row[_columnInfo["content"] - 1],
        row[_columnInfo["done"] - 1],
        row[_columnInfo["deadline"] - 1] !== ""
          ? ` ${row[_columnInfo["deadline"] - 1]}`
          : "期限なし",
      ];
      if (done === true) {
        continue;
      }
      const statusIcon = getStatusIcon(done, deadline);
      const strikeoutLine = done === true ? "~" : "";
      message += `${strikeoutLine}${statusIcon}#${num} ${assignee} ${content} ${deadline}${strikeoutLine}\n`;
    }
    return message;
  }
}

function add(spreadSheet, channelId, parsedText) {
  const sheet = setSheet(spreadSheet, channelId);
  const lastRow = sheet.getLastRow();
  parsedText.unshift(lastRow + 1); // タスク番号を追加
  const lastText = parsedText[parsedText.length - 1];
  const deadline = isValidDate(lastText) ? parsedText.pop() : "期限なし";
  const [num, assignee, ...content] = parsedText;
  const length = Object.keys(_columnInfo).length;
  const range = sheet.getRange(lastRow + 1, 1, 1, length);
  range.setNumberFormat("@");
  range.setValues([[num, assignee, content.join(" "), "", deadline]]);
  const statusIcon = getStatusIcon(false, deadline);
  const message = `次のタスクを追加しました：${statusIcon}#${num} ${assignee} ${content} ${deadline}`;
  return message;
}

function done(spreadSheet, channelId, taskNum) {
  const sheet = spreadSheet.getSheetByName(channelId);
  if (!sheet) {
    return "タスクはありません";
  }
  
  sheet.getRange(taskNum, _columnInfo["done"]).setValue("true");
  const [num, assignee, content, done, deadline] = getTask(sheet, taskNum)
  const statusIcon = getStatusIcon(done, deadline)
  const message = `次のタスクを完了しました： ${statusIcon}#${num} ${assignee} ${content}`;
  return message;
}

function assign(spreadSheet, channelId, taskNum, assignee) {
  const sheet = spreadSheet.getSheetByName(channelId);
  if (!sheet) {
    return "タスクはありません";
  }
  sheet.getRange(taskNum, _columnInfo["assignee"]).setValue(assignee);
  const [num, _, content, done, deadline] = getTask(sheet, taskNum);
  const statusIcon = getStatusIcon(done, deadline);
  const message = `タスクに担当を割り当てました： ${statusIcon}#${num} ${assignee} ${content} ${deadline}`;
  return message;
}

function setDeadline(spreadSheet, channelId, taskNum, deadline) {
    if (!isValidDate(deadline)) {
        return "期限設定に失敗しました。期限を指定するには有効な日付を yyyy/mm/dd 形式で指定してください。"
    }
    const sheet = spreadSheet.getSheetByName(channelId);
    if (!sheet) {
      return "タスクはありません";
    }
    const [num, assignee, content, done, _] = getTask(sheet, taskNum);
    if (done) {
        return `#${num} のタスクはすでに完了しています。`
    }
    const range = sheet.getRange(taskNum, _columnInfo["deadline"])
    range.setNumberFormat("@")
    range.setValue(deadline);
    const statusIcon = getStatusIcon(done, deadline);
    const message = `タスクに期限を設定しました： ${statusIcon}#${num} ${assignee} ${content} ${deadline}`;
    return message;
  }

function getTask(sheet, taskNum) {
  const length = Object.keys(_columnInfo).length;
  const row = sheet.getRange(taskNum, 1, 1, length).getValues();
  const [num, assignee, content, done, deadline] = [
    row[0][_columnInfo["num"] - 1],
    row[0][_columnInfo["assignee"] - 1],
    row[0][_columnInfo["content"] - 1],
    row[0][_columnInfo["done"] - 1],
    row[0][_columnInfo["deadline"] - 1],
  ];
  return [num, assignee, content, done, deadline];
}

/**
 * Gets the all values in the sheet.
 *
 * @param sheet — the Sheet object
 * @return Object[][] — a two-dimensional array of values
 */
function getAllValues(sheet) {
  // get all values
  if (isEmptySheet(sheet)) {
    return [];
  }
  return sheet.getDataRange().getValues();
}

/**
 * The sheet is empty?
 *
 * @param sheet — the Sheet object
 * @return Boolean — true if the sheet is empty; false otherwise
 */
function isEmptySheet(sheet) {
  // the range is totally blank?
  return sheet.getDataRange().isBlank();
}

function setSheet(spreadSheet, name) {
  //同じ名前のシートがなければ作成
  let sheet = spreadSheet.getSheetByName(name);
  if (!sheet) {
    sheet = spreadSheet.insertSheet();
    sheet.setName(name);
  }
  return sheet;
}

function isValidDate(string) {
  if (typeof string == "string") {
    const dateStrings = string.match(/^(\d+)\/(\d+)\/(\d+)$/);
    if (dateStrings) {
      const year = parseInt(dateStrings[1]);
      const month = parseInt(dateStrings[2]) - 1;
      const day = parseInt(dateStrings[3]);
      const date = new Date(year, month, day);
      return (
        year == date.getFullYear() &&
        month == date.getMonth() &&
        day == date.getDate()
      );
    }
  }
  return false;
}

function getStatusIcon(done, deadline = undefined) {
  return done === true
    ? "✅ "
    : new Date(deadline).setDate(new Date(deadline).getDate() + 1) < new Date()
    ? ":warning: "
    : "⬛️️ ";
}
