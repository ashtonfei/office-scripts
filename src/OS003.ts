const SHEET_NAME = {
  TWEETS: "Tweets",
  POSTED: "Posted",
};

function main(wb: ExcelScript.Workbook): Tweet {
  const wsTweets = wb.getWorksheet(SHEET_NAME.TWEETS);
  const { tweet, index } = selectTweet(wsTweets);
  const wsPosted = wb.getWorksheet(SHEET_NAME.POSTED);
  const lastRow = wsPosted.getUsedRange().getRowCount();
  wsPosted
    .getRangeByIndexes(lastRow, 0, 1, 2)
    .setValues([[tweet.text, new Date().toISOString()]]);
  wsTweets
    .getRange(`${index + 2}:${index + 2}`)
    .delete(ExcelScript.DeleteShiftDirection.up);
  return tweet;
}

function selectTweet(ws: ExcelScript.Worksheet) {
  const [headers, ...values] = ws.getUsedRange().getTexts();

  let tweet: Tweet = {
    text: "",
    media: {},
  };

  const { index, item } = randomSelect(values);
  const [text, topics, links] = item;
  const hashtagses = topics
    .replace(/\s/g, "")
    .split(",")
    .filter((v) => v !== "")
    .map((v) => `#${v}`)
    .join(" ");
  const tweetLinks = links
    .replace(/\s/g, "")
    .split(",")
    .filter((v) => v !== "")
    .join("\n");
  tweet.text = `${text} ${hashtagses} ${tweetLinks}`;

  return { tweet, index };
}

function randomSelect(items: string[][]) {
  const index = Math.floor(Math.random() * items.length);
  return {
    item: items[index],
    index,
  };
}

interface Tweet {
  text: string;
  media?: Object;
}
