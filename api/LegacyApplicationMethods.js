import {
  APPS_SCRIPT_SPREADSHEET_ID,
  CREDIT_ACTIVITY_MASTER_SHEET_ID,
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
      result.error?.message ||
        `Google Sheets request failed (${response.status})`,
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
    fields:
      "sheets(properties(sheetId,title),data(rowData(values(effectiveFormat(textFormat(foregroundColor))))))",
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
    Math.abs(color.red || 0) < 0.001 &&
    Math.abs(color.green || 0) < 0.001 &&
    Math.abs((color.blue || 0) - 1) < 0.001
  ) {
    return "#0000ff";
  }
  if (
    Math.abs((color.red || 0) - 1) < 0.001 &&
    Math.abs(color.green || 0) < 0.001 &&
    Math.abs(color.blue || 0) < 0.001
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
  const normalizedPassword = String(password ?? "")
    .trim()
    .toLowerCase();
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
      String(row[7] ?? "")
        .trim()
        .toLowerCase() === normalizedPassword
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
            ? (row[index] ?? "")
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
    for (const { base, quantity, price, expiryDays } of itemConsumptionMap.get(
      rootItem.Item,
    ) || []) {
      if (quantity <= 0) {
        throw new Error(
          `BOM quantity for '${rootItem.Item}' must be greater than zero`,
        );
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

  const activitySheetName = "Activity Master";

  const activityDataBefore = await getSheetData(env, activitySheetName);

  const firstNewRow = activityDataBefore.length + 1;

  await appendValues(env, MAIN_SHEET_ID, "'Activity Master'!A:I", rows);

  const activitySheetId = await getSheetId(env, activitySheetName);

  const newRowsColors = rows.map(() => "#000000");

  await updateFontColors(env, activitySheetId, firstNewRow, newRowsColors);

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
    rawStatus === "paid"
      ? "✅ Paid"
      : rawStatus === "pending"
        ? "❌ Pending"
        : "N/A";
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
  const values = await readValues(
    env,
    MAIN_SHEET_ID,
    `'${USER_ORDER_SHEET}'!A:P`,
  );
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
    await appendCreditActivity(env, saleData[0]);

    try {
      await updateNKDCreditBook(env);
    } catch (error) {
      console.error("Credit Book Error:", error.stack || error.toString());
    }
  }

  const billNumber = String(saleData[0].billNo ?? "").trim();

  if (!billNumber) {
    throw new Error("billNo is required");
  }

  if (billNumber.startsWith("UO/")) {
    await markOrderDelivered(env, billNumber);
  }

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

  const salesSheetName = "Sale Master";

  const salesDataBefore = await getSheetData(env, salesSheetName);

  const firstNewRow = salesDataBefore.length + 1;

  await appendValues(env, MAIN_SHEET_ID, "'Sale Master'!A:W", rows);

  const salesSheetId = await getSheetId(env, salesSheetName);

  const newRowsColors = rows.map(() => "#000000");

  await updateFontColors(env, salesSheetId, firstNewRow, newRowsColors);

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
  if (
    users.length === 0 ||
    users.some((item) => !item || typeof item !== "object")
  ) {
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
  if (
    users.length === 0 ||
    users.some((item) => !item || typeof item !== "object")
  ) {
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
  if (
    entries.length === 0 ||
    entries.some((item) => !item || typeof item !== "object")
  ) {
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
  const sheetRows = await readValues(
    env,
    MAIN_SHEET_ID,
    `'${USER_ORDER_SHEET}'!A:Z`,
  );
  const headers = sheetRows[0] || [];
  const billNoIndex = headers.indexOf("BillNo");
  const statusIndex = headers.indexOf("OrderStatus");
  const keyIndex = headers.indexOf("Key");
  if (billNoIndex < 0 || statusIndex < 0 || keyIndex < 0) {
    throw new Error(
      "User Order Master must contain BillNo, OrderStatus, and Key headers",
    );
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
      updates.push({
        range: `'${USER_ORDER_SHEET}'!A${rowNumber}:Z${rowNumber}`,
        values: [updated],
      });
      continue;
    }
    if (replacement) {
      const updated = headers.map((header, column) =>
        column < 2
          ? (current[column] ?? "")
          : replacement[header] !== undefined
            ? replacement[header]
            : (current[column] ?? ""),
      );
      updates.push({
        range: `'${USER_ORDER_SHEET}'!A${rowNumber}:Z${rowNumber}`,
        values: [updated],
      });
    }
  }
  if (updates.length) {
    await sheetsRequest(
      env,
      MAIN_SHEET_ID,
      "/values:batchUpdate?valueInputOption=USER_ENTERED",
      {
        method: "POST",
        body: JSON.stringify({ data: updates }),
      },
    );
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
  if (
    ![...knownSheets.values()].some(
      (entry) => entry.spreadsheetId === requestedId,
    )
  ) {
    throw new Error(
      "The requested sheetId is not registered in CreditNameMaster",
    );
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

  const customerRows = await readValues(
    env,
    customer.spreadsheetId,
    "Sheet1!A:E",
  );
  const row2 = customerRows[1] || [];
  const lastBalance = Number(row2[4]) || 0;
  const newBalance = lastBalance - amount;
  const customerMeta = await getSheetMetadata(
    env,
    customer.spreadsheetId,
    "Sheet1!A:E",
  );
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
  await updateValues(
    env,
    CREDIT_ACTIVITY_SPREADSHEET_ID,
    "'CreditRecdMaster'!A2:D2",
    [[timestamp(), name, amount, selectedType]],
  );
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
  const password = String(input.password ?? "")
    .trim()
    .toLowerCase();
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
      String(row[7] ?? "")
        .trim()
        .toLowerCase() === password
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
    if (
      String(row[8] ?? "")
        .trim()
        .toLowerCase() === password
    ) {
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

async function updateActivityMaster_github(env) {
  const [activity, stock, activityMetadata] = await Promise.all([
    readValues(env, MAIN_SHEET_ID, "'Activity Master'!A:Z"),
    readValues(env, MAIN_SHEET_ID, "'Stock Master'!A:Z"),
    getSheetMetadata(env, MAIN_SHEET_ID, "'Activity Master'!A:Z"),
  ]);
  const stockHeaders = stock[0] || [];
  const stockRows = stock.slice(1).map((row) => [...row]);
  const colorUpdates = [];
  const errors = [];
  let processed = 0;
  let skippedBlue = 0;
  let skippedRed = 0;
  let skippedNonProduction = 0;
  const transactionTypeCounts = {};

  for (let index = 1; index < activity.length; index++) {
    const rowNumber = index + 1;
    const row = activity[index];
    const transactionType = String(row[5] ?? "").trim();
    transactionTypeCounts[transactionType || "[blank]"] =
      (transactionTypeCounts[transactionType || "[blank]"] || 0) + 1;

    const color = fontColorAt(activityMetadata, index);
    if (color === "#0000ff") {
      skippedBlue++;
      continue;
    }
    if (color === "#ff0000") {
      skippedRed++;
      errors.push({
        rowNumber,
        message:
          "This row was previously marked as failed; correct it and clear its red font to retry",
      });
      continue;
    }
    const item = String(row[1] ?? "").trim();
    const batch = row[2];
    const rawQuantity = row[3];
    const location = String(row[6] ?? "").trim();
    const category = row[7] ?? "";
    if (transactionType !== "Production") {
      skippedNonProduction++;
      continue;
    }
    try {
      const locationIndex = stockHeaders.indexOf(location);
      const batchIndex = stockHeaders.indexOf("Batch/Date of Manufacture");
      const itemIndex = stockHeaders.indexOf("Item");
      if (locationIndex < 0 || batchIndex < 0 || itemIndex < 0) {
        throw new Error(
          "Stock Master is missing the item, batch, or location column",
        );
      }
      const batchDate = formattedDate(batch);
      const quantity = Number(rawQuantity) || 0;
      let found = false;
      for (let stockIndex = 0; stockIndex < stockRows.length; stockIndex++) {
        const stockRow = stockRows[stockIndex];
        if (
          String(stockRow[itemIndex] ?? "").trim() === item &&
          formattedDate(stockRow[batchIndex]) === batchDate
        ) {
          stockRow[locationIndex] =
            (Number(stockRow[locationIndex]) || 0) + quantity;
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
      colorUpdates.push({ rowNumber, color: "blue" });
      processed++;
    } catch (error) {
      console.error(
        `Failed to process Activity Master row ${rowNumber}:`,
        error,
      );
      errors.push({
        rowNumber,
        message: error?.message || "Failed to update stock from production row",
      });
      colorUpdates.push({ rowNumber, color: "red" });
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
  return {
    processed,
    skippedBlue,
    skippedRed,
    skippedNonProduction,
    transactionTypeCounts,
    errors,
  };
}

// Fixed dependency map
function getFixedDependencyMap() {
  return {
    "Pizza Mini Pan": {
      items: { "Pizza Base Mini": 1.0 },
      count: 1.0,
      src: 1.0,
    },
    "Cream Roll Choco": {
      items: { "Cream roll waffer": 1.0 },
      count: 1.0,
      src: 1.0,
    },
    "Burger Veggie Cheese": {
      items: { "Aloo Tikki": 1.0, "Bun Plain": 1.0 },
      count: 2.0,
      src: 1.0,
    },
    "Cream Roll Plain": {
      items: { "Cream roll waffer": 1.0 },
      count: 1.0,
      src: 1.0,
    },
    "Grilled Sandwich": {
      items: { "Bread Whole Wheat Slices": 2.0 },
      count: 1.0,
      src: 1.0,
    },
    "Chaumeen Full Plate": {
      items: { "Nkd Noddles": 150.0 },
      count: 1.0,
      src: 1.0,
    },
    "Gol Gappe Pani": { items: { "Gol Gappe": 5.0 }, count: 1.0, src: 1.0 },
    "Icing cake": { items: { "Icing cake Base": 1.0 }, count: 1.0, src: 1.0 },
    "Bun Plain Row": { items: { "Bun Plain": 1.0 }, count: 1.0, src: 1.0 },
    "Gol Gappe Chaat": { items: { "Gol Gappe": 5.0 }, count: 1.0, src: 1.0 },
    "Chaumeen Half Plate": {
      items: { "Nkd Noddles": 80.0 },
      count: 1.0,
      src: 1.0,
    },
    Pastry: { items: { "Pastry Base": 1.0 }, count: 1.0, src: 1.0 },
    "Pizza Bun": { items: { "Bun Plain": 1.0 }, count: 1.0, src: 1.0 },
    Tart: { items: { "Tart biscuits": 1.0 }, count: 1.0, src: 1.0 },
    "Samosa Row": { items: { Samosa: 1.0 }, count: 1.0, src: 1.0 },
    "Burger Veggie": {
      items: { "Bun Plain": 1.0, "Aloo Tikki": 1.0 },
      count: 2.0,
      src: 1.0,
    },
    "Samosa Chaat": { items: { Samosa: 1.0 }, count: 1.0, src: 1.0 },
    "Bread Whole Wheat": {
      items: { "Bread Whole Wheat Slices": 10.0 },
      count: 1.0,
      src: 1.0,
    },
  };
}

async function updateStockViaSaleNew_github(env) {
  const [sales, stock, salesMetadata] = await Promise.all([
    readValues(env, MAIN_SHEET_ID, "'Sale Master'!A:W"),
    readValues(env, MAIN_SHEET_ID, "'Stock Master'!A:Z"),
    getSheetMetadata(env, MAIN_SHEET_ID, "'Sale Master'!A:W"),
  ]);
  if (!sales.length || !stock.length) {
    throw new Error("Sale Master or Stock Master is missing its header row");
  }

  const stockHeaders = stock[0];
  const stockRows = stock.slice(1).map((row) => [...row]);
  const dependencyMap = getFixedDependencyMap();
  const colorUpdates = [];
  const errors = [];
  let processed = 0;
  let skippedBlue = 0;
  let skippedRed = 0;

  for (let index = 1; index < sales.length; index++) {
    const rowNumber = index + 1;
    const fontColor = fontColorAt(salesMetadata, index);
    if (fontColor === "#0000ff") {
      skippedBlue++;
      continue;
    }
    if (fontColor === "#ff0000") {
      skippedRed++;
      errors.push({
        rowNumber,
        message:
          "This row was previously marked as failed; correct it and clear its red font to retry",
      });
      continue;
    }

    const sale = sales[index];
    const item = String(sale[2] ?? "").trim();
    const batch = sale[3];
    const quantity = Number(sale[4]) || 0;
    const location = String(sale[10] ?? "").trim();
    const category = sale[13] ?? "";

    try {
      if (!item || !batch || quantity === 0) {
        throw new Error("Sale row is missing item, batch, or quantity");
      }

      if (dependencyMap[item]) {
        for (const [dependentItem, multiplier] of Object.entries(
          dependencyMap[item],
        )) {
          const requiredQuantity = multiplier * quantity;
          const stockRow = stockRows.find(
            (row) => row[1] === dependentItem && Number(row[3]) > 0,
          );

          if (stockRow) {
            stockRow[3] = (Number(stockRow[3]) || 0) - requiredQuantity;
            stockRow[0] = timestamp();
            stockRow[4] = category;
          } else {
            const newRow = Array(stockHeaders.length).fill("");
            newRow[0] = timestamp();
            newRow[1] = dependentItem;
            newRow[2] = formattedDate(new Date());
            newRow[3] = -requiredQuantity;
            newRow[4] = category;
            stockRows.push(newRow);
          }
        }
      } else {
        const locationIndex = stockHeaders.indexOf(location);
        if (locationIndex < 0) {
          throw new Error(`Stock Master has no location column '${location}'`);
        }

        const batchDate = formattedDate(batch);
        const stockRow = stockRows.find(
          (row) =>
            String(row[1] ?? "").trim() === item &&
            formattedDate(row[2]) === batchDate,
        );
        if (!stockRow) {
          throw new Error(`No stock row found for '${item}' on ${batchDate}`);
        }

        stockRow[locationIndex] =
          (Number(stockRow[locationIndex]) || 0) - quantity;
        stockRow[0] = timestamp();
        stockRow[4] = category;
      }

      colorUpdates.push({ rowNumber, color: "blue" });
      processed++;
    } catch (error) {
      console.error(`Failed to process Sale Master row ${rowNumber}:`, error);
      errors.push({
        rowNumber,
        message: error?.message || "Failed to update stock from sale row",
      });
      colorUpdates.push({ rowNumber, color: "red" });
    }
  }

  if (stockRows.length) {
    await updateValues(env, MAIN_SHEET_ID, "'Stock Master'!A2", stockRows);
  }
  await setFontColors(
    env,
    MAIN_SHEET_ID,
    "Sale Master",
    colorUpdates,
    salesMetadata,
  );

  return { processed, skippedBlue, skippedRed, errors };
}

async function updateStock(env) {
  const activityResult = await updateActivityMaster_github(env);
  const saleResult = await updateStockViaSaleNew(env);
  const errors = [
    ...activityResult.errors.map((error) => ({
      source: "Activity Master",
      ...error,
    })),
    ...saleResult.errors.map((error) => ({
      source: "Sale Master",
      ...error,
    })),
  ];

  return {
    status: errors.length === 0,
    activityRowsProcessed: activityResult.processed,
    activityRowsSkippedBlue: activityResult.skippedBlue,
    activityRowsSkippedRed: activityResult.skippedRed,
    activityRowsSkippedNonProduction: activityResult.skippedNonProduction,
    activityTransactionTypeCounts: activityResult.transactionTypeCounts,
    saleRowsProcessed: saleResult.processed,
    saleRowsSkippedBlue: saleResult.skippedBlue,
    saleRowsSkippedRed: saleResult.skippedRed,
    errors,
  };
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
  UPDATE_STOCK: (_requestData, env) => FUN_UPDATE_STOCK(env),
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

// Google Sheets API request
async function sheetsRequestGPT(env, url, options = {}) {
  const accessToken = await getGoogleAccessToken(env);

  const response = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
  });

  const text = await response.text();

  let data;

  try {
    data = text ? JSON.parse(text) : {};
  } catch {
    data = text;
  }

  if (!response.ok) {
    throw new Error(
      `Google Sheets API Error ${response.status}: ${JSON.stringify(data)}`,
    );
  }

  return data;
}

// Get complete sheet data
async function getSheetData(env, sheetName) {
  const range = `'${sheetName}'`;
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${MAIN_SHEET_ID}/values/${encodeURIComponent(range)}?valueRenderOption=UNFORMATTED_VALUE`;

  const result = await sheetsRequestGPT(env, url);

  return result.values || [];
}

// Update sheet values
async function updateSheetValues(
  env,
  sheetName,
  startRow,
  startColumn,
  values,
) {
  if (!values || values.length === 0) {
    return;
  }

  const startColumnLetter = columnToLetter(startColumn);
  const endColumnLetter = columnToLetter(startColumn + values[0].length - 1);
  const endRow = startRow + values.length - 1;

  const range = `'${sheetName}'!${startColumnLetter}${startRow}:${endColumnLetter}${endRow}`;
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${MAIN_SHEET_ID}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`;

  return sheetsRequestGPT(env, url, {
    method: "PUT",
    body: JSON.stringify({
      range,
      majorDimension: "ROWS",
      values,
    }),
  });
}

// Convert column number to Excel column letter
function columnToLetter(column) {
  let result = "";

  while (column > 0) {
    const remainder = (column - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    column = Math.floor((column - 1) / 26);
  }

  return result;
}

// Format date as dd-MMM-yyyy
function formatDate(value) {
  if (value === null || value === undefined || value === "") {
    return "";
  }

  let date;

  if (typeof value === "number") {
    date = new Date((value - 25569) * 86400000);
  } else {
    date = new Date(value);
  }

  if (isNaN(date.getTime())) {
    return "";
  }

  const months = [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
  ];

  const day = String(date.getDate()).padStart(2, "0");
  const month = months[date.getMonth()];
  const year = date.getFullYear();

  return `${day}-${month}-${year}`;
}

// Get current timestamp
function getTimestamp() {
  const now = new Date();

  const day = String(now.getDate()).padStart(2, "0");

  const months = [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
  ];

  const month = months[now.getMonth()];
  const year = now.getFullYear();

  let hours = now.getHours();
  const minutes = String(now.getMinutes()).padStart(2, "0");

  const amPm = hours >= 12 ? "PM" : "AM";

  hours = hours % 12;
  if (hours === 0) {
    hours = 12;
  }

  return `${day}-${month}-${year} ${hours}:${minutes} ${amPm}`;
}

// Get Google Sheet ID by sheet name
async function getSheetId(env, sheetName) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${MAIN_SHEET_ID}?fields=sheets.properties`;

  const result = await sheetsRequestGPT(env, url);

  const sheet = (result.sheets || []).find(
    (x) => x.properties?.title === sheetName,
  );

  if (!sheet) {
    throw new Error(`Sheet not found: ${sheetName}`);
  }

  return sheet.properties.sheetId;
}

// Get font colors from column A
async function getColumnAColors(env, sheetName, rowCount) {
  if (rowCount <= 1) {
    return [];
  }

  const range = `'${sheetName}'!A2:A${rowCount}`;
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${MAIN_SHEET_ID}?includeGridData=true&ranges=${encodeURIComponent(range)}`;

  const result = await sheetsRequestGPT(env, url);

  const sheetData = result.sheets?.[0]?.data?.[0]?.rowData || [];

  return sheetData.map((row) => {
    const color = row.values?.[0]?.effectiveFormat?.textFormat?.foregroundColor;

    if (!color) {
      return "";
    }

    const r = color.red || 0;
    const g = color.green || 0;
    const b = color.blue || 0;

    if (b > 0.8 && r < 0.2 && g < 0.2) {
      return "#0000ff";
    }

    if (r > 0.8 && g < 0.2 && b < 0.2) {
      return "#ff0000";
    }

    return "";
  });
}

// Get font colors from column A
async function getColumnAColorsActivity(
  env,
  spreadsheetId,
  sheetName,
  rowCount,
) {
  if (rowCount <= 1) {
    return [];
  }

  const range = `'${sheetName}'!A2:A${rowCount}`;

  const url =
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}` +
    `?includeGridData=true` +
    `&ranges=${encodeURIComponent(range)}`;

  const result = await sheetsRequestGPT(env, url, {
    method: "GET",
  });

  const sheetData = result.sheets?.[0]?.data?.[0]?.rowData || [];

  return sheetData.map((row) => {
    const color = row.values?.[0]?.effectiveFormat?.textFormat?.foregroundColor;

    if (!color) {
      return "";
    }

    const r = color.red || 0;
    const g = color.green || 0;
    const b = color.blue || 0;

    if (b > 0.8 && r < 0.2 && g < 0.2) {
      return "#0000ff";
    }

    if (r > 0.8 && g < 0.2 && b < 0.2) {
      return "#ff0000";
    }

    return "#000000";
  });
}

// Update font colors in column A
async function updateFontColors(env, sheetId, startRow, colors) {
  if (!colors || colors.length === 0) {
    return;
  }

  const requests = colors.map((color, index) => {
    const normalizedColor = String(color || "").toLowerCase();

    let foregroundColor;

    if (normalizedColor === "#0000ff") {
      foregroundColor = {
        red: 0,
        green: 0,
        blue: 1,
      };
    } else if (normalizedColor === "#ff0000") {
      foregroundColor = {
        red: 1,
        green: 0,
        blue: 0,
      };
    } else {
      foregroundColor = {
        red: 0,
        green: 0,
        blue: 0,
      };
    }

    return {
      repeatCell: {
        range: {
          sheetId,
          startRowIndex: startRow - 1 + index,
          endRowIndex: startRow + index,
          startColumnIndex: 0,
          endColumnIndex: 1,
        },
        cell: {
          userEnteredFormat: {
            textFormat: {
              foregroundColor,
            },
          },
        },
        fields: "userEnteredFormat.textFormat.foregroundColor",
      },
    };
  });

  const url = `https://sheets.googleapis.com/v4/spreadsheets/${MAIN_SHEET_ID}:batchUpdate`;

  return sheetsRequestGPT(env, url, {
    method: "POST",
    body: JSON.stringify({ requests }),
  });
}

// Main stock update method
export async function FUN_UPDATE_STOCK(env) {
  console.log("FUN_UPDATE_STOCK started");

  try {
    console.log("Calling updateActivityMaster");
    await updateActivityMaster(env);

    console.log("Calling updateStockViaSaleNew");
    const result = await updateStockViaSaleNew(env);

    console.log("FUN_UPDATE_STOCK completed");

    return result;
  } catch (error) {
    console.error("FUN_UPDATE_STOCK Error:", error.stack || error.toString());

    // Add your common email function here if required
    // await SendMailToAdmin(env, "NKD Sales Error", "FUN_UPDATE_STOCK", error.stack || error.toString());

    throw error;
  }
}

// Update stock based on sales
async function updateStockViaSaleNew(env) {
  console.log("updateStockViaSaleNew started");

  console.log(getTimestamp(), "Starting stock update via sales");

  const response = {
    status: false,
    output: [],
    errors: [],
  };

  try {
    const dependencyData = getFixedDependencyMap();
    const salesSheetName = "Sale Master";
    const stockSheetName = "Stock Master";

    const salesData = await getSheetData(env, salesSheetName);
    debugger;
    const stockData = await getSheetData(env, stockSheetName);

    if (!salesData.length) {
      throw new Error("Sale Master is empty");
    }

    if (!stockData.length) {
      throw new Error("Stock Master is empty");
    }

    const stockHeaders = stockData[0];

    console.log("Stock Headers:", stockHeaders);

    const salesFontColors = await getColumnAColors(
      env,
      salesSheetName,
      salesData.length,
    );

    const updatedFontColors = [];

    for (let i = 1; i < salesData.length; i++) {
      const salesRow = salesData[i] || [];
      const fontColor = salesFontColors[i - 1] || "";

      console.log("Processing Sale Row:", i + 1);

      if (fontColor === "#0000ff" || fontColor === "#ff0000") {
        console.log("Sale row already processed:", i + 1, fontColor);
        updatedFontColors.push([fontColor]);
        continue;
      }

      const saleItem = String(salesRow[2] || "").trim();
      const saleBatch = salesRow[3];
      const saleQty = Number(salesRow[4]) || 0;
      const saleLocation = String(salesRow[10] || "").trim();
      const category = salesRow[13];

      let updatedColor = "#ff0000";
      let isStockUpdated = false;

      console.log("Sale Details:", {
        row: i + 1,
        saleItem,
        saleBatch,
        saleQty,
        saleLocation,
        category,
      });

      if (!saleItem || !saleBatch || saleQty === 0) {
        response.errors.push({
          row: i + 1,
          item: saleItem,
          batch: saleBatch,
          qty: saleQty,
          location: saleLocation,
          message: "Invalid sale data",
        });

        updatedFontColors.push([updatedColor]);
        continue;
      }

      const formattedSaleBatch = formatDate(saleBatch);

      console.log("Formatted Sale Batch:", {
        original: saleBatch,
        formatted: formattedSaleBatch,
      });

      if (dependencyData[saleItem]) {
        const dependentItems = dependencyData[saleItem].items;

        console.log("Dependency item found:", {
          saleItem,
          dependentItems,
        });

        for (const dependentItem in dependentItems) {
          const requiredQty = dependentItems[dependentItem] * saleQty;
          let dependentFound = false;

          for (let j = 1; j < stockData.length; j++) {
            const row = stockData[j];

            const stockItem = String(row[1] || "").trim();
            const stockQty = Number(row[3]) || 0;

            if (stockItem === dependentItem && stockQty > 0) {
              const oldQty = stockQty;
              const newQty = oldQty - requiredQty;

              row[3] = newQty;
              row[0] = getTimestamp();
              row[4] = category;

              dependentFound = true;
              isStockUpdated = true;

              response.output.push({
                row: j + 1,
                item: stockItem,
                batch: formattedSaleBatch,
                location: saleLocation,
                saleQty,
                requiredQty,
                oldQty,
                newQty,
                category,
              });

              console.log("Dependency Stock Updated:", {
                row: j + 1,
                item: stockItem,
                oldQty,
                requiredQty,
                newQty,
              });

              break;
            }
          }

          if (!dependentFound) {
            const newRow = Array(stockHeaders.length).fill("");

            newRow[0] = getTimestamp();
            newRow[1] = dependentItem;
            newRow[2] = formattedSaleBatch;
            newRow[3] = -requiredQty;
            newRow[4] = category;

            stockData.push(newRow);

            isStockUpdated = true;

            response.output.push({
              row: stockData.length,
              item: dependentItem,
              batch: formattedSaleBatch,
              location: saleLocation,
              saleQty,
              requiredQty,
              oldQty: 0,
              newQty: -requiredQty,
              category,
              message: "New negative stock row created",
            });

            console.log("New Dependency Stock Row Created:", {
              item: dependentItem,
              requiredQty,
              newQty: -requiredQty,
            });
          }
        }
      } else {
        const locationIndex = stockHeaders.findIndex(
          (header) =>
            String(header || "")
              .trim()
              .toLowerCase() === saleLocation.toLowerCase(),
        );

        console.log("Normal Item Processing:", {
          saleItem,
          saleLocation,
          locationIndex,
        });

        if (locationIndex === -1) {
          response.errors.push({
            row: i + 1,
            item: saleItem,
            batch: formattedSaleBatch,
            location: saleLocation,
            message: `Location column not found in Stock Master: ${saleLocation}`,
          });

          updatedFontColors.push([updatedColor]);
          continue;
        }

        let itemFound = false;
        let batchFound = false;

        for (let j = 1; j < stockData.length; j++) {
          const row = stockData[j];

          const stockItem = String(row[1] || "").trim();
          const stockBatch = row[2];
          const formattedStockBatch = formatDate(stockBatch);

          if (stockItem === saleItem) {
            itemFound = true;

            console.log("Item Matched:", {
              stockRow: j + 1,
              stockItem,
              saleItem,
              stockBatch,
              formattedStockBatch,
              formattedSaleBatch,
            });
          }

          if (
            stockItem === saleItem &&
            formattedStockBatch === formattedSaleBatch
          ) {
            batchFound = true;

            const oldQty = Number(row[locationIndex]) || 0;
            const newQty = oldQty - saleQty;

            console.log("FINAL STOCK MATCH:", {
              stockRow: j + 1,
              item: stockItem,
              batch: formattedStockBatch,
              location: saleLocation,
              locationIndex,
              oldQty,
              saleQty,
              newQty,
            });

            row[locationIndex] = newQty;
            row[0] = getTimestamp();
            row[4] = category;

            isStockUpdated = true;

            response.output.push({
              row: j + 1,
              item: stockItem,
              batch: formattedStockBatch,
              location: saleLocation,
              oldQty,
              saleQty,
              newQty,
              category,
            });

            break;
          }
        }

        if (!itemFound) {
          console.log("ITEM NOT FOUND:", {
            saleItem,
            formattedSaleBatch,
          });

          response.errors.push({
            row: i + 1,
            item: saleItem,
            batch: formattedSaleBatch,
            location: saleLocation,
            message: "Item not found in Stock Master",
          });
        } else if (!batchFound) {
          console.log("BATCH NOT MATCHED:", {
            saleItem,
            formattedSaleBatch,
          });

          response.errors.push({
            row: i + 1,
            item: saleItem,
            batch: formattedSaleBatch,
            location: saleLocation,
            message: "Item found but batch did not match",
          });
        }
      }

      if (isStockUpdated) {
        updatedColor = "#0000ff";
      }

      updatedFontColors.push([updatedColor]);
    }

    if (stockData.length > 1) {
      console.log("Writing Stock Master:", {
        rows: stockData.length - 1,
        columns: stockHeaders.length,
      });

      await updateSheetValues(env, stockSheetName, 2, 1, stockData.slice(1));

      console.log("Stock Master write completed");
    }

    if (updatedFontColors.length > 0) {
      const saleSheetId = await getSheetId(env, salesSheetName);

      await updateFontColors(
        env,
        saleSheetId,
        2,
        updatedFontColors.map((x) => x[0]),
      );

      console.log("Sale Master colors updated");
    }

    const successfulUpdates = response.output.length;
    const failedUpdates = response.errors.length;

    response.status = failedUpdates === 0;

    response.summary = {
      totalSaleRows: salesData.length - 1,
      successfulUpdates,
      failedUpdates,
    };

    console.log("FINAL STOCK RESPONSE:", response);

    return response;
  } catch (error) {
    console.error(
      "updateStockViaSaleNew Error:",
      error.stack || error.toString(),
    );

    response.status = false;
    response.data = error.stack || error.toString();

    throw error;
  }
}

// Update stock based on Activity Master
async function updateActivityMaster(env) {
  console.log("updateActivityMaster started");

  const activitySheetName = "Activity Master";
  const stockSheetName = "Stock Master";

  const activityData = await getSheetData(env, activitySheetName);
  let stockData = await getSheetData(env, stockSheetName);

  if (!activityData.length) {
    return;
  }

  if (!stockData.length) {
    throw new Error("Stock Master is empty");
  }

  const stockHeaders = stockData[0];

  const activityFontColors = await getColumnAColors(
    env,
    activitySheetName,
    activityData.length,
  );

  const activityColors = [];

  for (let rowIndex = 1; rowIndex < activityData.length; rowIndex++) {
    const row = activityData[rowIndex] || [];

    const item = row[1];
    const batch = row[2];
    let qty = row[3];
    const transactionType = row[5];
    const location = row[6];
    const category = row[7];

    const fontColor = String(
      activityFontColors[rowIndex - 1] || "",
    ).toLowerCase();

    if (fontColor === "#0000ff" || fontColor === "#ff0000") {
      activityColors.push(fontColor);
      continue;
    }

    try {
      if (transactionType === "Production") {
        const formattedBatch = formatDate(batch);

        const locationIndex = stockHeaders.indexOf(location);
        const itemIndex = stockHeaders.indexOf("Item");
        const batchIndex = stockHeaders.indexOf("Batch/Date of Manufacture");

        if (locationIndex === -1) {
          throw new Error(`Location column not found: ${location}`);
        }

        if (itemIndex === -1) {
          throw new Error("Item column not found in Stock Master");
        }

        if (batchIndex === -1) {
          throw new Error("Batch/Date of Manufacture column not found");
        }

        qty = Number(qty) || 0;

        let itemFound = false;

        for (let i = 1; i < stockData.length; i++) {
          const stockRow = stockData[i];

          const stockItem = String(stockRow[itemIndex] || "").trim();

          const stockBatch = stockRow[batchIndex];

          const formattedStockBatch = formatDate(stockBatch);

          if (stockItem === item && formattedStockBatch === formattedBatch) {
            stockRow[locationIndex] =
              (Number(stockRow[locationIndex]) || 0) + qty;

            stockRow[0] = getTimestamp();
            stockRow[4] = category;

            itemFound = true;
            break;
          }
        }

        if (!itemFound) {
          const newRow = Array(stockHeaders.length).fill("");

          newRow[0] = getTimestamp();
          newRow[itemIndex] = item;
          newRow[batchIndex] = formattedBatch;
          newRow[locationIndex] = qty;
          newRow[4] = category;

          stockData.push(newRow);
        }

        activityColors.push("#0000ff");
      } else {
        activityColors.push("#000000");
      }
    } catch (error) {
      console.error(`Activity row ${rowIndex + 1} error:`, error);

      activityColors.push("#ff0000");
    }
  }

  if (stockData.length > 1) {
    await updateSheetValues(env, stockSheetName, 2, 1, stockData.slice(1));
  }

  if (activityColors.length > 0) {
    const activitySheetId = await getSheetId(env, activitySheetName);

    await updateFontColors(env, activitySheetId, 2, activityColors);
  }

  console.log("updateActivityMaster completed");
}

async function updateNKDCreditBook(env) {
  console.log("updateNKDCreditBook started");

  await createCustomerFile(env);

  await updateCustomerSheets(env);

  console.log("updateNKDCreditBook completed");

  return true;
}

async function createCustomerFile(env) {
  const activityData = await getSheetDataFromSpreadsheet(
    env,
    CREDIT_ACTIVITY_MASTER_SHEET_ID,
    "CreditActivityMaster",
  );

  const masterData = await getSheetDataFromSpreadsheet(
    env,
    CREDIT_ACTIVITY_MASTER_SHEET_ID,
    "CreditNameMaster",
  );

  const existingFiles = {};

  for (let i = 1; i < masterData.length; i++) {
    const customerName = String(masterData[i][1] ?? "").trim();
    const fileId = String(masterData[i][2] ?? "").trim();

    if (customerName && fileId) {
      existingFiles[customerName] = fileId;
    }
  }

  const newEntries = [];

  for (let i = 1; i < activityData.length; i++) {
    const customerName = String(activityData[i][1] ?? "").trim();

    if (!customerName) {
      continue;
    }

    if (existingFiles[customerName]) {
      continue;
    }

    console.log("Creating customer file:", customerName);

    const fileId = await copyCustomerTemplate(env, customerName);

    existingFiles[customerName] = fileId;

    newEntries.push([timestamp(new Date()), customerName, fileId]);
  }

  if (newEntries.length > 0) {
    await appendValues(
      env,
      CREDIT_ACTIVITY_MASTER_SHEET_ID,
      "'CreditNameMaster'!A:C",
      newEntries,
    );
  }

  return true;
}

async function copyCustomerTemplate(env, customerName) {
  const accessToken = await getGoogleAccessToken(env);

  const templateFileId = "1aQz2odviNF-lSaoLPcw88SCM3fgUPgpQOSfLf3SM09w";

  const parentFolderId = "1z89ooOqEmCCJIMPHL3JoVZVe-MGF4o03";

  const url = `https://www.googleapis.com/drive/v3/files/${templateFileId}/copy`;

  const response = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      name: customerName,
      parents: [parentFolderId],
    }),
  });

  const data = await response.json();

  if (!response.ok) {
    console.error("Drive Copy Error:", data);

    throw new Error(data?.error?.message || "Unable to copy customer template");
  }

  if (!data.id) {
    throw new Error("Customer file ID was not returned");
  }

  console.log("Customer file created:", customerName, data.id);

  return data.id;
}

async function getSheetIdFromSpreadsheet(env, spreadsheetId, sheetName) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?fields=sheets.properties`;

  const response = await sheetsRequestGPT(env, url, {
    method: "GET",
  });

  const sheets = response.sheets || [];

  const sheet = sheets.find((item) => item.properties?.title === sheetName);

  if (!sheet) {
    throw new Error(`Sheet not found: ${sheetName}`);
  }

  return sheet.properties.sheetId;
}

async function insertRowAfterFirstRow(env, spreadsheetId, sheetId) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`;

  return sheetsRequestGPT(env, url, {
    method: "POST",
    body: JSON.stringify({
      requests: [
        {
          insertDimension: {
            range: {
              sheetId,
              dimension: "ROWS",
              startIndex: 1,
              endIndex: 2,
            },
            inheritFromBefore: false,
          },
        },
      ],
    }),
  });
}

async function getSheetDataFromSpreadsheet(env, spreadsheetId, sheetName) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(
    `'${sheetName}'!A:Z`,
  )}`;

  const response = await sheetsRequestGPT(env, url, {
    method: "GET",
  });

  return response.values || [];
}

async function updateSpreadsheetValues(env, spreadsheetId, range, values) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(
    range,
  )}?valueInputOption=USER_ENTERED`;

  return sheetsRequestGPT(env, url, {
    method: "PUT",
    body: JSON.stringify({
      range,
      majorDimension: "ROWS",
      values,
    }),
  });
}

async function updateCreditActivityFontColor(env, sheetId, startRow, color) {
  const normalizedColor = String(color || "").toLowerCase();

  let foregroundColor;

  if (normalizedColor === "#0000ff") {
    foregroundColor = {
      red: 0,
      green: 0,
      blue: 1,
    };
  } else if (normalizedColor === "#ff0000") {
    foregroundColor = {
      red: 1,
      green: 0,
      blue: 0,
    };
  } else {
    foregroundColor = {
      red: 0,
      green: 0,
      blue: 0,
    };
  }

  const url = `https://sheets.googleapis.com/v4/spreadsheets/${CREDIT_ACTIVITY_MASTER_SHEET_ID}:batchUpdate`;

  return sheetsRequestGPT(env, url, {
    method: "POST",
    body: JSON.stringify({
      requests: [
        {
          repeatCell: {
            range: {
              sheetId,
              startRowIndex: startRow - 1,
              endRowIndex: startRow,
              startColumnIndex: 0,
              endColumnIndex: 1,
            },
            cell: {
              userEnteredFormat: {
                textFormat: {
                  foregroundColor,
                },
              },
            },
            fields: "userEnteredFormat.textFormat.foregroundColor",
          },
        },
      ],
    }),
  });
}

async function appendCreditActivity(env, saleData) {
  const now = new Date();

  const row = [
    timestamp(now),
    saleData.customerName ?? "",
    saleData.billNo ?? "",
    saleData.totalBillCost ?? "",
    saleData.paymentStatus ?? "",
  ];

  const sheetName = "CreditActivityMaster";

  const dataBefore = await getSheetDataFromSpreadsheet(
    env,
    CREDIT_ACTIVITY_MASTER_SHEET_ID,
    sheetName,
  );

  const firstNewRow = dataBefore.length + 1;

  await appendValues(
    env,
    CREDIT_ACTIVITY_MASTER_SHEET_ID,
    "'CreditActivityMaster'!A:E",
    [row],
  );

  const sheetId = await getSheetIdFromSpreadsheet(
    env,
    CREDIT_ACTIVITY_MASTER_SHEET_ID,
    sheetName,
  );

  await updateCreditActivityFontColor(env, sheetId, firstNewRow, ["#000000"]);

  return true;
}

async function updateCustomerSheets(env) {
  const activityData = await getSheetDataFromSpreadsheet(
    env,
    CREDIT_ACTIVITY_MASTER_SHEET_ID,
    "CreditActivityMaster",
  );

  const nameData = await getSheetDataFromSpreadsheet(
    env,
    CREDIT_ACTIVITY_MASTER_SHEET_ID,
    "CreditNameMaster",
  );

  const activitySheetId = await getSheetIdFromSpreadsheet(
    env,
    CREDIT_ACTIVITY_MASTER_SHEET_ID,
    "CreditActivityMaster",
  );

  const activityColors = await getColumnAColorsActivity(
    env,
    CREDIT_ACTIVITY_MASTER_SHEET_ID,
    "CreditActivityMaster",
    activityData.length,
  );

  const nameMap = {};

  for (let i = 1; i < nameData.length; i++) {
    const customerName = String(nameData[i][1] ?? "").trim();
    const sheetId = String(nameData[i][2] ?? "").trim();

    if (customerName && sheetId) {
      nameMap[customerName] = {
        sheetId,
        rowIndex: i + 1,
      };
    }
  }

  for (let j = 1; j < activityData.length; j++) {
    const rowNumber = j + 1;

    const fontColor = String(activityColors[j] ?? "").toLowerCase();

    if (fontColor === "#0000ff" || fontColor === "#ff0000") {
      continue;
    }

    const customerName = String(activityData[j][1] ?? "").trim();

    const billNo = activityData[j][2] ?? "";
    const billAmount = Number(activityData[j][3] ?? 0);

    if (!nameMap[customerName]) {
      console.error("Customer sheet not found:", customerName);

      await updateCreditActivityFontColor(env, activitySheetId, rowNumber, [
        "#ff0000",
      ]);

      continue;
    }

    const customerSpreadsheetId = nameMap[customerName].sheetId;

    const creditNameMasterRow = nameMap[customerName].rowIndex;

    try {
      const sheetId = await getSheetIdFromSpreadsheet(
        env,
        customerSpreadsheetId,
        "Sheet1",
      );

      const currentData = await getSheetDataFromSpreadsheet(
        env,
        customerSpreadsheetId,
        "Sheet1",
      );

      const row2 = currentData[1] || [];

      const lastBalance = Number(row2[4] ?? 0);

      const newBalance = lastBalance + billAmount;

      if (
        row2.some((cell) => cell !== "" && cell !== null && cell !== undefined)
      ) {
        await insertRowAfterFirstRow(env, customerSpreadsheetId, sheetId);
      }

      const inputDate = timestamp(new Date());

      const newRow = [inputDate, billNo, billAmount, "", newBalance];

      await updateSpreadsheetValues(
        env,
        customerSpreadsheetId,
        "'Sheet1'!A2:E2",
        [newRow],
      );

      await updateSpreadsheetValues(
        env,
        CREDIT_ACTIVITY_MASTER_SHEET_ID,
        `'CreditNameMaster'!D${creditNameMasterRow}`,
        [[newBalance]],
      );

      await updateCreditActivityFontColor(env, activitySheetId, rowNumber, [
        "#0000ff",
      ]);

      console.log("Credit updated:", customerName, billNo, newBalance);
    } catch (error) {
      console.error("Credit customer update error:", customerName, error);

      await updateCreditActivityFontColor(env, activitySheetId, rowNumber, [
        "#ff0000",
      ]);
    }
  }

  return true;
}
