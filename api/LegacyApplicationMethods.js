import {
  APPS_SCRIPT_SPREADSHEET_ID,
  CREDIT_ACTIVITY_SPREADSHEET_ID,
  USER_MASTER_SPREADSHEET_ID,
  getGoogleAccessToken,
} from "../utils/googleConfig.js";

const MAIN_SHEET_ID = APPS_SCRIPT_SPREADSHEET_ID;
const USER_ORDER_SHEET = "User Order Master";

async function sheetsRequest(env, spreadsheetId, path, options = {}) {
  const token = await getGoogleAccessToken(env);
  const response = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}${path}`,
    {
      ...options,
      headers: {
        Authorization: `Bearer ${token}`,
        ...(options.body ? { "Content-Type": "application/json" } : {}),
        ...options.headers,
      },
    },
  );
  const result = await response.json();

  if (!response.ok) {
    throw new Error(
      result.error?.message || `Google Sheets request failed (${response.status})`,
    );
  }

  return result;
}

async function readValues(env, spreadsheetId, range) {
  const path = `/values/${encodeURIComponent(range)}?valueRenderOption=FORMATTED_VALUE`;
  const result = await sheetsRequest(env, spreadsheetId, path);
  return result.values || [];
}

async function appendValues(env, spreadsheetId, range, values) {
  return sheetsRequest(
    env,
    spreadsheetId,
    `/values/${encodeURIComponent(range)}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`,
    {
      method: "POST",
      body: JSON.stringify({ values }),
    },
  );
}

async function updateValues(env, spreadsheetId, range, values) {
  return sheetsRequest(
    env,
    spreadsheetId,
    `/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`,
    {
      method: "PUT",
      body: JSON.stringify({ values }),
    },
  );
}

async function batchUpdate(env, spreadsheetId, requests) {
  return sheetsRequest(env, spreadsheetId, ":batchUpdate", {
    method: "POST",
    body: JSON.stringify({ requests }),
  });
}

async function getSheetMetadata(env, spreadsheetId, range) {
  const query = new URLSearchParams({
    fields: "sheets(properties(sheetId,title),data(rowData(values(effectiveFormat(textFormat(foregroundColor))))))",
    includeGridData: "true",
    ranges: range,
  });
  return sheetsRequest(env, spreadsheetId, `?${query}`);
}

function sheetIdFor(metadata, title) {
  const sheet = metadata.sheets?.find(
    (candidate) => candidate.properties?.title === title,
  );
  if (!sheet) {
    throw new Error(`Sheet '${title}' not found`);
  }
  return sheet.properties.sheetId;
}

function fontColorAt(metadata, rowIndex, columnIndex = 0) {
  const color =
    metadata.sheets?.[0]?.data?.[0]?.rowData?.[rowIndex]?.values?.[columnIndex]
      ?.effectiveFormat?.textFormat?.foregroundColor;
  if (!color) return "";
  if (
    Math.abs((color.red || 0)) < 0.001 &&
    Math.abs((color.green || 0)) < 0.001 &&
    Math.abs((color.blue || 0) - 1) < 0.001
  ) {
    return "#0000ff";
  }
  if (
    Math.abs((color.red || 0) - 1) < 0.001 &&
    Math.abs((color.green || 0)) < 0.001 &&
    Math.abs((color.blue || 0)) < 0.001
  ) {
    return "#ff0000";
  }
  return "";
}

async function setFontColors(env, spreadsheetId, title, updates, metadata) {
  if (!updates.length) return;
  const sheetMetadata =
    metadata || (await getSheetMetadata(env, spreadsheetId, `'${title}'!A:Z`));
  const sheetId = sheetIdFor(sheetMetadata, title);
  const requests = updates.map(({ rowNumber, color }) => ({
    repeatCell: {
      range: {
        sheetId,
        startRowIndex: rowNumber - 1,
        endRowIndex: rowNumber,
        startColumnIndex: 0,
        endColumnIndex: 1,
      },
      cell: {
        userEnteredFormat: {
          textFormat: {
            foregroundColor:
              color === "blue"
                ? { red: 0, green: 0, blue: 1 }
                : { red: 1, green: 0, blue: 0 },
          },
        },
      },
      fields: "userEnteredFormat.textFormat.foregroundColor",
    },
  }));
  await batchUpdate(env, spreadsheetId, requests);
}

function columnsToObjects(headers, rows) {
  return rows.map((row) =>
    headers.reduce((record, header, index) => {
      record[header] = row[index] ?? "";
      return record;
    }, {}),
  );
}

function parseJsonValue(value, fieldName) {
  if (typeof value !== "string") return value;
  try {
    return JSON.parse(value);
  } catch {
    throw new Error(`${fieldName} must contain valid JSON`);
  }
}

function indiaDateParts(date = new Date()) {
  const parts = new Intl.DateTimeFormat("en-GB", {
    timeZone: "Asia/Kolkata",
    day: "2-digit",
    month: "short",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
    hour12: true,
  }).formatToParts(date);
  return Object.fromEntries(parts.map(({ type, value }) => [type, value]));
}

function timestamp(date = new Date()) {
  const parts = indiaDateParts(date);
  return `${parts.day}-${parts.month}-${parts.year} ${parts.hour}:${parts.minute} ${parts.dayPeriod.toUpperCase()}`;
}

function dateOnly(date = new Date()) {
  const parts = indiaDateParts(date);
  return `${parts.day}-${parts.month}-${parts.year}`;
}

function formattedDate(value) {
  const dateParts = String(value ?? "").match(
    /^(\d{1,2})[-\s]([A-Za-z]{3})[-\s](\d{4})$/,
  );
  if (dateParts) {
    const month =
      dateParts[2][0].toUpperCase() + dateParts[2].slice(1).toLowerCase();
    return `${dateParts[1].padStart(2, "0")}-${month}-${dateParts[3]}`;
  }
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    throw new Error(`Invalid date value: ${value}`);
  }
  const parts = indiaDateParts(date);
  return `${parts.day}-${parts.month}-${parts.year}`;
}

function responseSuccess(requestData, data) {
  return {
    status: data === false || data?.status === false ? false : true,
    data,
    request: redactPasswords(requestData),
  };
}

function redactPasswords(value) {
  if (Array.isArray(value)) {
    return value.map(redactPasswords);
  }
  if (!value || typeof value !== "object") {
    return value;
  }

  return Object.fromEntries(
    Object.entries(value).map(([key, entry]) => [
      key,
      key.toLowerCase() === "password" ? "[REDACTED]" : redactPasswords(entry),
    ]),
  );
}

function unsupported(apiType, reason) {
  return {
    status: false,
    apiType,
    message: `${apiType} is not available in the Worker: ${reason}`,
    blocked: true,
  };
}

async function listUsers(password, includeName = false, env) {
  const [masterRows, otherRows] = await Promise.all([
    readValues(env, USER_MASTER_SPREADSHEET_ID, "'NKD Master'!A2:M"),
    readValues(env, USER_MASTER_SPREADSHEET_ID, "'Other User Master'!A2:N"),
  ]);
  const normalizedPassword = String(password ?? "").trim().toLowerCase();
  const response = {
    data: [],
    isAdminAccess: false,
    role: "",
  };
  if (includeName) response.name = "";

  for (const row of masterRows) {
    if (!String(row[0] ?? "").trim()) break;
    if (row[3] !== "Active") continue;

    if (
      normalizedPassword &&
      String(row[7] ?? "").trim().toLowerCase() === normalizedPassword
    ) {
      response.role = row[10] ?? "";
      if (includeName) response.name = row[6] ?? "";
      response.isAdminAccess = ["Admin", "Super Admin"].includes(response.role);
    }

    response.data.push({
      name: row[6] ?? "",
      mobile: row[9] ?? "",
      schemeDiscount: row[12] ?? "",
      devType: "NKDDevotee",
    });
  }

  for (const row of otherRows) {
    if (!String(row[1] ?? "").trim()) break;
    response.data.push({
      name: row[1] ?? "",
      mobile: row[10] ?? "",
      devType: "Non-NKDDevotee",
      schemeDiscount: row[13] ?? "",
    });
  }

  response.data.sort((a, b) => String(a.name).localeCompare(String(b.name)));
  return response;
}

async function getStock(env) {
  const [stockTable, priceTable, bomRows] = await Promise.all([
    readValues(env, MAIN_SHEET_ID, "'Stock Master'!A:Z"),
    readValues(env, MAIN_SHEET_ID, "'Product List & Price Master'!A:P"),
    readValues(env, MAIN_SHEET_ID, "'BOM'!A:J"),
  ]);
  const headers = stockTable[0] || [];
  const stockRows = stockTable.slice(1);
  const extraHeaders = [
    "Price",
    "ExpiryDays",
    "Category",
    "ProductType",
    "Type",
    "HSN",
    "GST %",
  ];
  const priceMap = new Map();

  for (const row of priceTable.slice(1)) {
    const name = String(row[1] ?? "").trim();
    if (!name) break;
    priceMap.set(name, {
      price: row[2] ?? "",
      expiryDays: row[3] ?? "",
      category: row[6] ?? "",
      productType: row[7] ?? "",
      type: row[9] ?? "",
      hsn: row[14] ?? "",
      gstPercentage: row[15] ?? "",
    });
  }

  const itemConsumptionMap = new Map();
  for (const row of bomRows.slice(1)) {
    const dependentItem = row[0];
    const baseItem = row[1];
    if (!dependentItem || !baseItem) continue;
    const items = itemConsumptionMap.get(baseItem) || [];
    items.push({
      base: dependentItem,
      quantity: Number(row[7]) || 0,
      price: row[8] ?? "",
      expiryDays: row[9] ?? "",
    });
    itemConsumptionMap.set(baseItem, items);
  }

  const rows = [];
  for (const row of stockRows) {
    const itemName = String(row[1] ?? "").trim();
    if (!itemName) continue;
    const details = priceMap.get(itemName) || {
      price: "N/A",
      expiryDays: "",
      category: "",
      productType: "",
      type: "",
      hsn: "",
      gstPercentage: "",
    };
    const quantity = row[3] ?? "";
    if (
      details.productType !== "Root-dependent-item" &&
      (details.price === "N/A" ||
        details.price === "" ||
        quantity === 0 ||
        quantity === "0")
    ) {
      continue;
    }
    rows.push(
      Object.fromEntries(
        [...headers, ...extraHeaders].map((header, index) => [
          header,
          index < headers.length
            ? row[index] ?? ""
            : [
                details.price,
                details.expiryDays,
                details.category,
                details.productType,
                details.type,
                details.hsn,
                details.gstPercentage,
              ][index - headers.length],
        ]),
      ),
    );
  }

  const currentDate = new Date();
  const numericDate = Object.fromEntries(
    new Intl.DateTimeFormat("en-GB", {
      timeZone: "Asia/Kolkata",
      day: "2-digit",
      month: "2-digit",
      year: "numeric",
    })
      .formatToParts(currentDate)
      .map(({ type, value }) => [type, value]),
  );
  const dateString = `${numericDate.day}-${numericDate.month}-${numericDate.year}`;
  const timeString = currentDate.toLocaleTimeString("en-US", {
    timeZone: "Asia/Kolkata",
  });
  for (const rootItem of rows.filter(
    (row) => row.ProductType === "Root-dependent-item",
  )) {
    for (const { base, quantity, price, expiryDays } of
      itemConsumptionMap.get(rootItem.Item) || []) {
      if (quantity <= 0) {
        throw new Error(`BOM quantity for '${rootItem.Item}' must be greater than zero`);
      }
      rows.push({
        DateTime: `${dateString} ${timeString}`,
        Item: base,
        "Batch/Date of Manufacture": rootItem["Batch/Date of Manufacture"],
        Natures: (Number(rootItem.Natures) || 0) / quantity,
        Price: price,
        ExpiryDays: expiryDays,
        Category: rootItem.Category,
        ProductType: "Dependent-item",
        Type: "Qty",
        HSN: "0000",
        "GST %": 0,
      });
    }
  }

  return rows.filter((row) => row.ProductType !== "Root-dependent-item");
}

async function getProductList(env) {
  const table = await readValues(
    env,
    MAIN_SHEET_ID,
    "'Product List & Price Master'!A:I",
  );
  const products = [];
  for (const row of table.slice(1)) {
    if (!String(row[1] ?? "")) break;
    if (row[7] === "Dependent-item" || row[8] !== "Active") continue;
    products.push({ product: row[1] ?? "", category: row[7] ?? "" });
  }
  products.sort((a, b) => String(a.product).localeCompare(String(b.product)));
  return products;
}

async function saveProductionData(requestData, env) {
  const productionRows = parseJsonValue(
    requestData.productionData,
    "productionData",
  );
  if (!Array.isArray(productionRows) || productionRows.length === 0) {
    throw new Error("productionData must be a non-empty JSON array");
  }
  const now = new Date();
  const rows = productionRows.map((item) => [
    timestamp(now),
    item.selectedItem ?? "",
    item.batchDate ?? "",
    item.productionCount ?? "",
    item.comments ?? "",
    item.transactionType ?? "",
    item.location ?? "",
    item.category ?? "",
    dateOnly(now),
  ]);
  await appendValues(env, MAIN_SHEET_ID, "'Activity Master'!A:I", rows);
  return true;
}

function buildSaleMessage(saleData) {
  const grouped = new Map();
  for (const item of saleData) {
    const current = grouped.get(item.name);
    if (!current) {
      grouped.set(item.name, { ...item });
    } else {
      current.quantity += item.quantity;
      current.baseTotal += item.baseTotal;
      current.total += item.total;
      if (new Date(item.stockDateTime) > new Date(current.stockDateTime)) {
        current.stockDateTime = item.stockDateTime;
      }
    }
  }

  const selectedItems = [...grouped.values()];
  const first = selectedItems[0] || {};
  const rawStatus = String(first.paymentStatus || "n/a").toLowerCase();
  const formattedStatus =
    rawStatus === "paid" ? "✅ Paid" : rawStatus === "pending" ? "❌ Pending" : "N/A";
  let grandTotal = 0;
  let message =
    `🛒 Nature's Bill 🛒\n\n` +
    `Customer: ${first.customerName || ""}\n` +
    `Payment Status: ${formattedStatus}\n\n`;
  const saleItems = selectedItems.filter(
    (item) => Number(item.quantity) > 0 && Number(item.total) > 0,
  );
  const returnItems = selectedItems.filter(
    (item) => Number(item.quantity) < 0 && Number(item.total) < 0,
  );

  if (saleItems.length) {
    message += "🛍️ *Purchased Items*\n\n";
    saleItems.forEach((item, index) => {
      const baseTotal = Number(item.baseTotal) || 0;
      const discount = Number(item.discountAmount) || 0;
      const schemeDiscount = Number(item.schemeDiscountAmount) || 0;
      const netTotal = baseTotal - discount - schemeDiscount;
      grandTotal += netTotal;
      message += `${index + 1}. *${item.name}*\n`;
      message += `   Qty: ${item.quantity} × ₹${item.price} = ₹${baseTotal.toFixed(2)}\n`;
      if (discount > 0) message += `   🔹 Discount: -₹${discount.toFixed(2)}\n`;
      if (schemeDiscount > 0) {
        message += `   🔸 Scheme Discount: -₹${schemeDiscount.toFixed(2)}\n`;
      }
      if (discount > 0 || schemeDiscount > 0) {
        message += `   ➡ Final: ₹${netTotal.toFixed(2)}\n`;
      }
      message += "\n";
    });
    message += "\n";
  }

  if (returnItems.length) {
    message += "🔁 *Returned Items*\n";
    returnItems.forEach((item, index) => {
      const quantity = Math.abs(Number(item.quantity));
      const total = Math.abs(Number(item.total));
      grandTotal -= total;
      message += `${index + 1}. *${item.name}* - ${quantity} x ₹${item.price} = ₹${total.toFixed(2)}\n`;
    });
    message += "\n";
  }

  message += `*Grand Total: ₹${grandTotal.toFixed(2)}*\n\n`;
  message += "Chant and be Happy!\n🛍️ Hare Krishna 🛍️";
  return message;
}

async function markOrderDelivered(env, billNumber) {
  const values = await readValues(env, MAIN_SHEET_ID, `'${USER_ORDER_SHEET}'!A:P`);
  const matchingRows = [];
  for (let index = 1; index < values.length; index++) {
    if (values[index][8] === billNumber) matchingRows.push(index + 1);
  }
  await setFontColors(
    env,
    MAIN_SHEET_ID,
    USER_ORDER_SHEET,
    matchingRows.map((rowNumber) => ({ rowNumber, color: "blue" })),
  );
  if (matchingRows.length) {
    await sheetsRequest(
      env,
      MAIN_SHEET_ID,
      "/values:batchUpdate?valueInputOption=USER_ENTERED",
      {
        method: "POST",
        body: JSON.stringify({
          data: matchingRows.map((rowNumber) => ({
            range: `'${USER_ORDER_SHEET}'!P${rowNumber}`,
            values: [["Delivered"]],
          })),
        }),
      },
    );
  }
}

async function createSale(requestData, env) {
  const saleData = parseJsonValue(requestData.inputData, "inputData");
  if (!Array.isArray(saleData) || saleData.length === 0) {
    throw new Error("inputData must contain a non-empty sale array");
  }
  const paymentStatus = saleData[0].paymentStatus;
  if (paymentStatus === "pending") {
    throw new Error(
      "Pending sales require the Apps Script Drive-based customer credit book integration, which is not available in the Worker",
    );
  }
  const billNumber = String(saleData[0].billNo ?? "").trim();
  if (!billNumber) throw new Error("billNo is required");
  if (billNumber.startsWith("UO/")) await markOrderDelivered(env, billNumber);

  const now = new Date();
  const rows = saleData.map((item) => [
    timestamp(now),
    item.customerName ?? "",
    item.name ?? "",
    item.batch ?? "",
    item.quantity ?? "",
    item.price ?? "",
    item.discount ?? "",
    item.total ?? "",
    item.paymentStatus ?? "",
    item.type ?? "",
    item.location ?? "",
    item.saleType ?? "",
    item.billNo ?? "",
    item.category ?? "",
    item.devType ?? "",
    item.comment ?? "",
    dateOnly(now),
    item.totalBillCost ?? "",
    item.discountAmount ?? "",
    item.schemeDiscountAmount ?? "",
    item.gstStatus ?? "No",
    item.hsn ?? "",
    item.gstPercent ?? "",
  ]);
  await appendValues(env, MAIN_SHEET_ID, "'Sale Master'!A:W", rows);

  const first = saleData[0];
  await appendValues(env, MAIN_SHEET_ID, "'DailySaleDetails'!A:G", [
    [
      timestamp(now),
      dateOnly(now),
      first.customerName || "Unknown",
      first.totalBillCost || 0,
      first.billNo ?? "",
      first.paymentStatus ?? "",
      buildSaleMessage(saleData),
    ],
  ]);
  return true;
}

async function addNewUser(requestData, env) {
  const input = parseJsonValue(requestData.inputData, "inputData");
  const users = Array.isArray(input) ? input : [input];
  if (users.length === 0 || users.some((item) => !item || typeof item !== "object")) {
    throw new Error("inputData must contain a user object or array of users");
  }
  const now = new Date();
  const rows = users.map((user) => [
    timestamp(now),
    user.name ?? "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    user.email ?? "",
    user.mobileNumber ?? "",
    "Non-NKD",
    dateOnly(now),
  ]);
  await appendValues(
    env,
    USER_MASTER_SPREADSHEET_ID,
    "'Other User Master'!A:M",
    rows,
  );
  return true;
}

async function addLibraryUser(requestData, env) {
  const input = parseJsonValue(requestData.inputData, "inputData");
  const users = Array.isArray(input) ? input : [input];
  if (users.length === 0 || users.some((item) => !item || typeof item !== "object")) {
    throw new Error("inputData must contain a user object or array of users");
  }

  const now = new Date();
  const rows = users.map((user) => [
    user.name ?? "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    user.email ?? "",
    user.mobileNumber ?? "",
    "Non-NKD",
    timestamp(now),
    dateOnly(now),
  ]);
  await appendValues(env, MAIN_SHEET_ID, "'LibUserMaster'!A:M", rows);
  return true;
}

async function getSheetObjects(env, spreadsheetId, sheetName) {
  const values = await readValues(env, spreadsheetId, `'${sheetName}'!A:Z`);
  return columnsToObjects(values[0] || [], values.slice(1));
}

async function getLibraryBookList(env) {
  const [bookMaster, activityMaster] = await Promise.all([
    getSheetObjects(env, MAIN_SHEET_ID, "Book List & Price Master"),
    getSheetObjects(env, MAIN_SHEET_ID, "LibraryActivityMaster"),
  ]);
  return {
    bookMasterResponse: bookMaster,
    libraryActivityMasterResponse: activityMaster,
  };
}

async function getLibraryUserList(env) {
  return getSheetObjects(env, MAIN_SHEET_ID, "LibUserMaster");
}

async function issueLibraryBook(requestData, env) {
  const input = parseJsonValue(requestData.inputData, "inputData");
  const entries = Array.isArray(input) ? input : [input];
  if (entries.length === 0 || entries.some((item) => !item || typeof item !== "object")) {
    throw new Error("inputData must contain one or more issue-book objects");
  }

  const now = timestamp();
  const rows = entries.map((item) => [
    now,
    ...Object.keys(item).map((key) => item[key] ?? ""),
  ]);
  await appendValues(env, MAIN_SHEET_ID, "'LibraryActivityMaster'!A:Z", rows);
  return true;
}

async function saveUserOrder(requestData, env) {
  const input = parseJsonValue(requestData.inputData, "inputData");
  const values = input?.data;
  const items = Array.isArray(values) ? values : values ? [values] : [];
  if (items.length === 0) {
    throw new Error("inputData.data must contain one or more order objects");
  }
  const now = timestamp();
  const rows = items.map((item) => [
    now,
    ...Object.keys(item).map((key) => item[key] ?? ""),
  ]);
  await appendValues(env, MAIN_SHEET_ID, `'${USER_ORDER_SHEET}'!A:Z`, rows);
  return true;
}

async function getUserOrderList(env) {
  const [values, metadata] = await Promise.all([
    readValues(env, MAIN_SHEET_ID, `'${USER_ORDER_SHEET}'!A:Z`),
    getSheetMetadata(env, MAIN_SHEET_ID, `'${USER_ORDER_SHEET}'!A:Z`),
  ]);
  const headers = values[0] || [];
  const colors =
    metadata.sheets?.[0]?.data?.[0]?.rowData?.map((row) =>
      fontColorAt({ sheets: [{ data: [{ rowData: [row] }] }] }, 0),
    ) || [];
  const result = {};
  for (let index = 1; index < values.length; index++) {
    if (colors[index] === "#0000ff") continue;
    const record = Object.fromEntries(
      headers.map((header, column) => [header, values[index][column] ?? ""]),
    );
    const key = `${record.Name} - ${record.BillNo}`;
    (result[key] ||= []).push(record);
  }
  return result;
}

async function readyUserOrder(requestData, env) {
  const newData = parseJsonValue(requestData.inputData, "inputData");
  if (!Array.isArray(newData) || newData.length === 0) {
    throw new Error("inputData must be a non-empty order array");
  }
  const sheetRows = await readValues(env, MAIN_SHEET_ID, `'${USER_ORDER_SHEET}'!A:Z`);
  const headers = sheetRows[0] || [];
  const billNoIndex = headers.indexOf("BillNo");
  const statusIndex = headers.indexOf("OrderStatus");
  const keyIndex = headers.indexOf("Key");
  if (billNoIndex < 0 || statusIndex < 0 || keyIndex < 0) {
    throw new Error("User Order Master must contain BillNo, OrderStatus, and Key headers");
  }
  const targetBillNo = newData[0].BillNo;
  const newKeys = new Set(newData.map((item) => item.Key));
  const updates = [];
  for (let index = 1; index < sheetRows.length; index++) {
    const current = sheetRows[index];
    if (current[billNoIndex] !== targetBillNo) continue;
    const rowNumber = index + 1;
    const replacement = newData.find((item) => item.Key === current[keyIndex]);
    if (!newKeys.has(current[keyIndex])) {
      const updated = [...current];
      updated[statusIndex] = "Deleted";
      updates.push({ range: `'${USER_ORDER_SHEET}'!A${rowNumber}:Z${rowNumber}`, values: [updated] });
      continue;
    }
    if (replacement) {
      const updated = headers.map((header, column) =>
        column < 2
          ? current[column] ?? ""
          : replacement[header] !== undefined
            ? replacement[header]
            : current[column] ?? "",
      );
      updates.push({ range: `'${USER_ORDER_SHEET}'!A${rowNumber}:Z${rowNumber}`, values: [updated] });
    }
  }
  if (updates.length) {
    await sheetsRequest(env, MAIN_SHEET_ID, "/values:batchUpdate?valueInputOption=USER_ENTERED", {
      method: "POST",
      body: JSON.stringify({ data: updates }),
    });
  }
  return true;
}

async function getKhataBookUserList(env) {
  const values = await readValues(
    env,
    CREDIT_ACTIVITY_SPREADSHEET_ID,
    "'CreditNameMaster'!A:Z",
  );
  const records = columnsToObjects(values[0] || [], values.slice(1));
  records.sort((a, b) => String(a.Name).localeCompare(String(b.Name)));
  return records;
}

async function getCreditSheetMap(env) {
  const values = await readValues(
    env,
    CREDIT_ACTIVITY_SPREADSHEET_ID,
    "'CreditNameMaster'!A:D",
  );
  const mapping = new Map();
  for (let index = 1; index < values.length; index++) {
    const name = String(values[index][1] ?? "").trim();
    const spreadsheetId = String(values[index][2] ?? "").trim();
    if (name && spreadsheetId) {
      mapping.set(name, { spreadsheetId, rowNumber: index + 1 });
    }
  }
  return mapping;
}

async function getKhataBookByUserId(requestData, env) {
  const input = parseJsonValue(requestData.inputData, "inputData") || {};
  const requestedId = String(input.sheetId ?? "").trim();
  if (!requestedId) throw new Error("inputData.sheetId is required");
  const knownSheets = await getCreditSheetMap(env);
  if (![...knownSheets.values()].some((entry) => entry.spreadsheetId === requestedId)) {
    throw new Error("The requested sheetId is not registered in CreditNameMaster");
  }
  const values = await readValues(env, requestedId, "Sheet1!A:Z");
  const headers = values[0] || [];
  const rows = [];
  for (const row of values.slice(1)) {
    rows.push(row);
    if (Number.parseInt(row[4], 10) < 10) break;
  }
  return columnsToObjects(headers, rows);
}

async function updateCreditBalance(requestData, env) {
  const input = parseJsonValue(requestData.inputData, "inputData") || {};
  const name = String(input.name ?? "").trim();
  const amount = Number(input.amount);
  const selectedType = String(input.selectedType ?? "").trim();
  if (!name || !Number.isFinite(amount) || !selectedType) {
    throw new Error("inputData.name, amount, and selectedType are required");
  }
  const customer = (await getCreditSheetMap(env)).get(name);
  if (!customer) throw new Error(`Customer sheet not found: ${name}`);

  const customerRows = await readValues(env, customer.spreadsheetId, "Sheet1!A:E");
  const row2 = customerRows[1] || [];
  const lastBalance = Number(row2[4]) || 0;
  const newBalance = lastBalance - amount;
  const customerMeta = await getSheetMetadata(env, customer.spreadsheetId, "Sheet1!A:E");
  const customerSheetId = sheetIdFor(customerMeta, "Sheet1");
  const row2IsEmpty = row2.every((value) => value === "");
  const customerRequests = [];
  if (!row2IsEmpty) {
    customerRequests.push({
      insertDimension: {
        range: {
          sheetId: customerSheetId,
          dimension: "ROWS",
          startIndex: 1,
          endIndex: 2,
        },
        inheritFromBefore: true,
      },
    });
  }
  if (customerRequests.length) {
    await batchUpdate(env, customer.spreadsheetId, customerRequests);
  }
  await updateValues(env, customer.spreadsheetId, "Sheet1!A2:E2", [
    [timestamp(), "", "", amount, newBalance],
  ]);

  const creditMeta = await getSheetMetadata(
    env,
    CREDIT_ACTIVITY_SPREADSHEET_ID,
    "'CreditRecdMaster'!A:D",
  );
  const creditSheetId = sheetIdFor(creditMeta, "CreditRecdMaster");
  await batchUpdate(env, CREDIT_ACTIVITY_SPREADSHEET_ID, [
    {
      insertDimension: {
        range: {
          sheetId: creditSheetId,
          dimension: "ROWS",
          startIndex: 1,
          endIndex: 2,
        },
        inheritFromBefore: true,
      },
    },
  ]);
  await updateValues(env, CREDIT_ACTIVITY_SPREADSHEET_ID, "'CreditRecdMaster'!A2:D2", [
    [timestamp(), name, amount, selectedType],
  ]);
  await updateValues(
    env,
    CREDIT_ACTIVITY_SPREADSHEET_ID,
    `'CreditNameMaster'!D${customer.rowNumber}`,
    [[newBalance]],
  );
  return true;
}

async function getUserInfoByPassword(requestData, env) {
  const input = parseJsonValue(requestData.inputData, "inputData") || {};
  const password = String(input.password ?? "").trim().toLowerCase();
  if (!password) throw new Error("inputData.password is required");
  const [masterRows, otherRows] = await Promise.all([
    readValues(env, USER_MASTER_SPREADSHEET_ID, "'NKD Master'!A2:M"),
    readValues(env, USER_MASTER_SPREADSHEET_ID, "'Other User Master'!A2:N"),
  ]);
  const response = {
    isUserFound: false,
    isAdminAccess: false,
    role: "",
    userDetails: null,
  };
  for (const row of masterRows) {
    if (!String(row[0] ?? "")) break;
    if (
      row[3] === "Active" &&
      String(row[7] ?? "").trim().toLowerCase() === password
    ) {
      response.isUserFound = true;
      response.role = row[10] ?? "";
      response.isAdminAccess = ["Admin", "Super Admin"].includes(response.role);
      response.userDetails = {
        name: row[6] ?? "",
        mobile: row[9] ?? "",
        role: row[10] ?? "",
        email: row[8] ?? "",
        schDiscount: row[12] ?? "",
        devType: "NKDDevotee",
      };
      return response;
    }
  }
  for (const row of otherRows) {
    if (!String(row[1] ?? "")) break;
    if (String(row[8] ?? "").trim().toLowerCase() === password) {
      response.isUserFound = true;
      response.userDetails = {
        name: row[1] ?? "",
        mobile: row[10] ?? "",
        role: row[9] ?? "",
        email: row[7] ?? "",
        schDiscount: row[11] ?? "",
        devType: "Non-NKDDevotee",
      };
      response.role = response.userDetails.role;
      response.isAdminAccess = ["Admin", "Super Admin"].includes(response.role);
      return response;
    }
  }
  return response;
}

async function updateActivityMaster(env) {
  const [activity, stock, activityMetadata] = await Promise.all([
    readValues(env, MAIN_SHEET_ID, "'Activity Master'!A:Z"),
    readValues(env, MAIN_SHEET_ID, "'Stock Master'!A:Z"),
    getSheetMetadata(env, MAIN_SHEET_ID, "'Activity Master'!A:Z"),
  ]);
  const stockHeaders = stock[0] || [];
  const stockRows = stock.slice(1).map((row) => [...row]);
  const colorUpdates = [];

  for (let index = 1; index < activity.length; index++) {
    if (["#0000ff", "#ff0000"].includes(fontColorAt(activityMetadata, index))) {
      continue;
    }
    const row = activity[index];
    const [inputDate, item, batch, rawQuantity, , transactionType, location, category] = row;
    if (transactionType !== "Production") continue;
    try {
      const locationIndex = stockHeaders.indexOf(location);
      const batchIndex = stockHeaders.indexOf("Batch/Date of Manufacture");
      const itemIndex = stockHeaders.indexOf("Item");
      if (locationIndex < 0 || batchIndex < 0 || itemIndex < 0) {
        throw new Error("Stock Master is missing the item, batch, or location column");
      }
      const batchDate = formattedDate(batch);
      const quantity = Number(rawQuantity) || 0;
      let found = false;
      for (let stockIndex = 0; stockIndex < stockRows.length; stockIndex++) {
        const stockRow = stockRows[stockIndex];
        if (
          stockRow[itemIndex] === item &&
          formattedDate(stockRow[batchIndex]) === batchDate
        ) {
          stockRow[locationIndex] = (Number(stockRow[locationIndex]) || 0) + quantity;
          stockRow[0] = timestamp();
          stockRow[4] = category ?? "";
          found = true;
          break;
        }
      }
      if (!found) {
        const newRow = Array(stockHeaders.length).fill("");
        newRow[0] = timestamp();
        newRow[itemIndex] = item ?? "";
        newRow[batchIndex] = batchDate;
        newRow[locationIndex] = quantity;
        newRow[4] = category ?? "";
        stockRows.push(newRow);
      }
      colorUpdates.push({ rowNumber: index + 1, color: "blue" });
    } catch (error) {
      console.error(`Failed to process Activity Master row ${index + 1}:`, error);
      colorUpdates.push({ rowNumber: index + 1, color: "red" });
    }
  }

  if (stockRows.length) {
    await updateValues(env, MAIN_SHEET_ID, "'Stock Master'!A2", stockRows);
  }
  await setFontColors(
    env,
    MAIN_SHEET_ID,
    "Activity Master",
    colorUpdates,
    activityMetadata,
  );
  return null;
}

export const LEGACY_IMPLEMENTATIONS = {
  GET_STOCK: (_requestData, env) => getStock(env),
  GET_ALL_USER_LIST: (requestData, env) =>
    listUsers(requestData.password, false, env),
  GET_PRODUCT_LIST: (_requestData, env) => getProductList(env),
  SAVE_PRODUCTION_DATA: (requestData, env) =>
    saveProductionData(requestData, env),
    CREATE_SALE: (requestData, env) => createSale(requestData, env),
    UPDATE_ACTIVITY_MASTER: (_requestData, env) => updateActivityMaster(env),
  ADD_NEW_USER: (requestData, env) => addNewUser(requestData, env),
  ADD_LIB_USER: (requestData, env) => addLibraryUser(requestData, env),
  LIB_BOOK_LIST: (_requestData, env) => getLibraryBookList(env),
  LIB_USER_LIST: (_requestData, env) => getLibraryUserList(env),
  LIB_ISSUE_BOOK: (requestData, env) => issueLibraryBook(requestData, env),
  GET_KHATA_BOOK_USER_LIST: (_requestData, env) => getKhataBookUserList(env),
  UPDATE_CUST_CREDIT_BALANCE: (requestData, env) =>
    updateCreditBalance(requestData, env),
  GET_KHATA_BOOK_BY_USER_ID: (requestData, env) =>
    getKhataBookByUserId(requestData, env),
  GET_USER_INFO_BY_PASSWORD: (requestData, env) =>
    getUserInfoByPassword(requestData, env),
  SAVE_USER_ORDER_DATA: (requestData, env) => saveUserOrder(requestData, env),
  GET_USER_ORDER_LIST: (_requestData, env) => getUserOrderList(env),
  READY_USER_ORDER: (requestData, env) => readyUserOrder(requestData, env),
};

export const BLOCKED_LEGACY_APIS = {
  UPDATE_STOCK_VIA_SALE:
    "the API handler calls updateStockViaSale(), which is absent from the supplied implementation; the available newer helper is called only inside the LockService-protected update flow",
  INSERT_DAILY_INPUT:
    "the Apps Script route sends an email reconciliation report, and no email integration is configured for the Worker",
  UPDATE_STOCK:
    "the Apps Script operation depends on a script lock and stock-dependency processing that are not available as a safe Worker operation",
  CREATE_SALE_NKD:
    "the Apps Script route invokes Drive-based invoice rendering and downstream integrations that are not available in the Worker",
  GENERATE_NATURES_GST_INVOICE:
    "the Apps Script invoice generator depends on Google Drive PDF rendering and related integrations that are not available in the Worker",
};

export async function runLegacyApi(apiType, requestData, env) {
  const implementation = LEGACY_IMPLEMENTATIONS[apiType];
  if (!implementation) {
    return unsupported(apiType, BLOCKED_LEGACY_APIS[apiType]);
  }

  try {
    const result = await implementation(requestData, env);
    return responseSuccess(requestData, result);
  } catch (error) {
    return {
      status: false,
      data: error?.message || "Apps Script API failed",
      request: requestData,
    };
  }
}
