function updateWarframeMarketData() {
  const sheetName = "Item Tracker";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  sheet.getRange("M6").setValue("");
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Sheet Not Found: "${sheetName}"`);
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const itemNames = sheet.getRange("A2:A" + lastRow).getValues().flat();
  const amounts = sheet.getRange("B2:B" + lastRow).getValues().flat();

  const resultRange = sheet.getRange("C2:J" + lastRow);
  resultRange.clearContent();
  resultRange.setBackground("white");

  const errorRange = sheet.getRange("K13:K1000");
  errorRange.clearContent();
  errorRange.setBackground("white");

  sheet.getRange("K10:M10").setValue("\uD83D\uDD04 Starting…");
  sheet.getRange("M6").setValue("");
  SpreadsheetApp.flush();

  let errorRow = 13;
  let errorcnt = 0;
  let warncnt = 0;

  for (let i = 0; i < itemNames.length; i++) {
    if (sheet.getRange("M6").getValue().toString().toUpperCase() === "STOP") {
      sheet.getRange("K10:M10").setValue("\u26D4\uFE0F Update aborted.");
      console.log("⛔️ Update aborted.");
      return;
    }

    const itemName = itemNames[i];
    const row = i + 2;

    if (!itemName || itemName.trim() === "") {
      console.log(`⚠️ Blank item at row ${row}, stopping search.`);
      break;
    }

    const amount = Number(amounts[i]) || 0;
    const urlName = convertToUrlName(itemName);

    let sellerAvg = "N/A";
    let buyerAvg = "N/A";
    let ducats = "No Data";
    let hadError = false;
    let errorReason = "";
    let sellerCount = 0;
    let buyerCount = 0;

    sheet.getRange("K10:M10").setValue(`\uD83D\uDD0D Searching: ${itemName}`);
    SpreadsheetApp.flush();
    console.log(`🔍 Fetching: ${itemName}`);

    for (let attempt = 1; attempt <= 3; attempt++) {
      try {
        const itemData = fetchWithRetry(`https://api.warframe.market/v1/items/${urlName}`);
        const itemSet = itemData.payload?.item?.items_in_set?.find(x => x.url_name === urlName);
        if (itemSet && itemSet.ducats !== undefined) {
          ducats = itemSet.ducats;
        } else {
          hadError = true;
          errorReason = "Item not found or mis‑spelled";
          break;
        }

        const orderData = fetchWithRetry(`https://api.warframe.market/v1/items/${urlName}/orders`);
        const orders = orderData.payload?.orders || [];

        const sellStatuses = ["ingame", "online", "offline"];
        for (const status of sellStatuses) {
          const sellers = orders
            .filter(o => o.visible && o.order_type === "sell" && o.user.status === status)
            .sort((a, b) => a.platinum - b.platinum);
          if (sellers.length > 0) {
            sellerCount = sellers.length;
            sellerAvg = sellers[0].platinum;
            break;
          }
        }

        const buyStatuses = ["ingame", "online", "offline"];
        for (const status of buyStatuses) {
          const buyers = orders
            .filter(o => o.visible && o.order_type === "buy" && o.user.status === status)
            .sort((a, b) => b.platinum - a.platinum);
          if (buyers.length > 0) {
            buyerCount = buyers.length;
            buyerAvg = buyers[0].platinum;
            break;
          }
        }

        if (sellerAvg === "N/A" && buyerAvg === "N/A") {
          hadError = true;
          errorReason = "No market data found";
        } else {
          if (buyerAvg === "N/A") errorReason = "No buyers found";
          if (sellerAvg === "N/A") errorReason = "No sellers found";
        }

        break;
      } catch (err) {
        console.warn(`Attempt ${attempt} failed for ${itemName}: ${err}`);
        if (attempt === 3) {
          hadError = true;
          errorReason = "Exception thrown after retries";
        }
        Utilities.sleep(200);
      }
    }

    const sellerToDucat = (typeof sellerAvg === "number" && typeof ducats === "number")
      ? (ducats / sellerAvg).toFixed(2)
      : "NO VALUE";

    const buyerToDucat = (typeof buyerAvg === "number" && typeof ducats === "number")
      ? (ducats / buyerAvg).toFixed(2)
      : "NO VALUE";

    const totalPlatSeller = (typeof sellerAvg === "number" && amount > 0)
      ? sellerAvg * amount
      : "NO VALUE";

    const totalPlatBuyer = (typeof buyerAvg === "number" && amount > 0)
      ? buyerAvg * amount
      : "NO VALUE";

    const totalDucats = (typeof ducats === "number" && amount > 0)
      ? ducats * amount
      : "NO VALUE";

    const sellerCell = sheet.getRange(row, 3);
    const buyerCell = sheet.getRange(row, 4);
    sellerCell.setValue(sellerAvg);
    buyerCell.setValue(buyerAvg);

    if (typeof sellerAvg === "number" && sellerCount < 10) sellerCell.setBackground("#fff3cd");
    if (typeof buyerAvg === "number" && buyerCount < 10) buyerCell.setBackground("#fff3cd");

    sheet.getRange(row, 5).setValue(ducats);
    sheet.getRange(row, 6).setValue(sellerToDucat);
    sheet.getRange(row, 7).setValue(buyerToDucat);
    sheet.getRange(row, 8).setValue(totalPlatSeller);
    sheet.getRange(row, 9).setValue(totalPlatBuyer);
    sheet.getRange(row, 10).setValue(totalDucats);

    if (hadError) {
      const errorMsg = `❌ ${itemName} — ${errorReason}`;
      sheet.getRange("K" + errorRow).setValue(errorMsg);
      sheet.getRange("K" + errorRow).setBackground("#ffcccc");
      console.log(errorMsg);
      errorcnt++;
      errorRow++;
    } else if (
      (typeof sellerAvg === "number" && sellerCount < 10) ||
      (typeof buyerAvg === "number" && buyerCount < 10) ||
      buyerAvg === "N/A" || sellerAvg === "N/A"
    ) {
      const warnMsgParts = [];
      if (sellerAvg === "N/A") warnMsgParts.push("No sellers found");
      else if (sellerCount < 10) warnMsgParts.push(`${sellerCount} seller/s`);

      if (buyerAvg === "N/A") warnMsgParts.push("No buyers found");
      else if (buyerCount < 10) warnMsgParts.push(`${buyerCount} buyer/s`);

      const warnMsg = `⚠️ ${itemName} — Limited data: ${warnMsgParts.join(" and ")}`;
      sheet.getRange("K" + errorRow).setValue(warnMsg);
      sheet.getRange("K" + errorRow).setBackground("#fff3cd");
      console.log(warnMsg);
      warncnt++;
      errorRow++;
    }

    Utilities.sleep(350);
  }

  let doneMsg = "";
  if (errorcnt === 0 && warncnt === 0) {
    doneMsg = "✅ All data fetched!";
  } else if (errorcnt === 0 && warncnt > 0) {
    doneMsg = "⚠️ All data fetched, but market data is limited.";
  } else {
    const parts = [];
    if (errorcnt > 0) parts.push(`${errorcnt} error(s)`);
    if (warncnt > 0) parts.push(`${warncnt} warning(s)`);
    doneMsg = `⚠️ ${parts.join(" and ")} found.`;
  }

  sheet.getRange("K10:M10").setValue(doneMsg);
  console.log(doneMsg);
}

function fetchWithRetry(url) {
  let attempt = 0;
  while (attempt < 3) {
    try {
      const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (res.getResponseCode() === 200) {
        return JSON.parse(res.getContentText());
      } else {
        throw new Error(`Status code ${res.getResponseCode()}`);
      }
    } catch (e) {
      attempt++;
      if (attempt >= 3) throw e;
      Utilities.sleep(200);
    }
  }
}

function convertToUrlName(name) {
  return name.toLowerCase().replace(/[\s-]/g, "_");
}
