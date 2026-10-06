import {
  SPREADSHEET_ID,
  USER_MASTER_SPREADSHEET_ID,
  getGoogleAccessToken,
} from "../utils/googleConfig.js";
import { addNumbers, sayHello, multiplyNumbers } from "nkd-common-dev-lib";

// GET DATA

export async function GET_DATA(inputData, env) {
  const sheetName = inputData.sheetName;
  //testing honey again
  if (!sheetName) {
    return {
      status: false,
      message: "sheetName is required",
    };
  }

  const accessToken = await getGoogleAccessToken(env);
  const range = encodeURIComponent(`'${sheetName}'`);
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${range}`;

  const response = await fetch(url, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
  });

  const data = await response.json();

  if (!response.ok) {
    console.error("GET_DATA Google Error:", data);

    throw new Error(data.error?.message || "Unable to read Google Sheet");
  }

  const libraryTest = {
    addResult: addNumbers(10, 20),
    helloResult: sayHello("Honey"),
    multiplyResult: multiplyNumbers(5, 5),
  };

  return {
    status: true,
    message: "GET_DATA successful",

    // Existing Google Sheet data
    data: data.values || [],

    // Common Library Test
    libraryTest: libraryTest,
  };
}

// SAVE DATA

export async function SAVE_DATA(inputData, env) {
  const sheetName = inputData.sheetName;
  const rowData = inputData.rowData;

  if (!sheetName) {
    return {
      status: false,
      message: "sheetName is required",
    };
  }

  if (!Array.isArray(rowData)) {
    return {
      status: false,
      message: "rowData must be an array",
    };
  }

  const accessToken = await getGoogleAccessToken(env);
  const range = encodeURIComponent(`'${sheetName}'`);

  const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${range}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`;

  const response = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      values: [rowData],
    }),
  });

  const data = await response.json();

  if (!response.ok) {
    console.error("SAVE_DATA Google Error:", data);

    throw new Error(
      data.error?.message || "Unable to save data to Google Sheet",
    );
  }

  return {
    status: true,
    message: "SAVE_DATA successful",
    data: data,
  };
}

// DELETE DATA

export async function DELETE_DATA(inputData, env) {
  const sheetName = inputData.sheetName;
  const rowNumber = Number(inputData.rowNumber ?? inputData.rowIndex);

  if (!sheetName) {
    return {
      status: false,
      message: "sheetName is required",
    };
  }

  // Row numbers are 1-based.

  if (!Number.isInteger(rowNumber) || rowNumber < 1) {
    return {
      status: false,
      message: "Valid rowNumber is required",
    };
  }

  const accessToken = await getGoogleAccessToken(env);

  const metadataUrl = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}?fields=sheets.properties`;

  const metadataResponse = await fetch(metadataUrl, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
  });

  const metadata = await metadataResponse.json();

  if (!metadataResponse.ok) {
    console.error("Spreadsheet Metadata Error:", metadata);

    throw new Error(
      metadata.error?.message || "Unable to get spreadsheet metadata",
    );
  }

  const sheet = metadata.sheets?.find((s) => s.properties?.title === sheetName);

  if (!sheet) {
    return {
      status: false,
      message: `Sheet '${sheetName}' not found`,
    };
  }

  const sheetId = sheet.properties.sheetId;
  const startIndex = rowNumber - 1;
  const endIndex = rowNumber;

  const batchUrl = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}:batchUpdate`;

  const response = await fetch(batchUrl, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      requests: [
        {
          deleteDimension: {
            range: {
              sheetId: sheetId,
              dimension: "ROWS",
              startIndex: startIndex,
              endIndex: endIndex,
            },
          },
        },
      ],
    }),
  });

  const data = await response.json();

  if (!response.ok) {
    console.error("DELETE_DATA Google Error:", data);

    throw new Error(data.error?.message || "Unable to delete row");
  }

  return {
    status: true,
    message: "DELETE_DATA successful",
    data: {
      sheetName: sheetName,
      deletedRowNumber: rowNumber,
    },
  };
}

// SEARCH VOUCHER

export async function SEARCH_VOUCHER(inputData, env) {
  const sheetName = inputData?.sheetName || "Sheet1";
  const voucherId = String(inputData?.voucherId || "").trim();

  if (!voucherId) {
    return {
      status: false,
      message: "Voucher ID is required",
    };
  }

  try {
    // Get voucher IDs from the first column.

    const range = `${sheetName}!A:A`;
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encodeURIComponent(range)}`;

    const token = await getGoogleAccessToken(env);

    const googleResponse = await fetch(url, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${token}`,
      },
    });

    const googleData = await googleResponse.json();

    if (!googleResponse.ok) {
      return {
        status: false,
        message: "Google Sheets API error",
        googleStatus: googleResponse.status,
        error: googleData,
      };
    }

    const values = googleData.values || [];

    // Find voucher in the first column.

    const rowIndex = values.findIndex(
      (row) => String(row?.[0] || "").trim() === voucherId,
    );

    if (rowIndex === -1) {
      return {
        status: false,
        message: "Voucher ID not found",
        voucherId: voucherId,
      };
    }

    // Google Sheet row number is index + 1.

    const rowNumber = rowIndex + 1;

    // Get the complete voucher row.

    const rowRange = `${sheetName}!A${rowNumber}:Q${rowNumber}`;
    const rowUrl = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encodeURIComponent(rowRange)}`;

    const rowResponse = await fetch(rowUrl, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${token}`,
      },
    });

    const rowData = await rowResponse.json();

    if (!rowResponse.ok) {
      return {
        status: false,
        message: "Unable to get voucher details",
        googleStatus: rowResponse.status,
        error: rowData,
      };
    }

    const row = rowData.values?.[0] || [];

    return {
      status: true,
      message: "Voucher found",
      voucherId: voucherId,
      rowNumber: rowNumber,
      data: row,
    };
  } catch (error) {
    return {
      status: false,
      message: "Error searching voucher",
      error: error.message,
    };
  }
}

export async function GET_ALL_USER_LIST_NEW(inputData, env) {
  const password = String(inputData?.password ?? "")
    .trim()
    .toLowerCase();
  const accessToken = await getGoogleAccessToken(env);
  const ranges = ["'NKD Master'!A2:M", "'Other User Master'!A2:N"];

  const sheetValues = await Promise.all(
    ranges.map(async (range) => {
      const url = `https://sheets.googleapis.com/v4/spreadsheets/${USER_MASTER_SPREADSHEET_ID}/values/${encodeURIComponent(range)}?valueRenderOption=FORMATTED_VALUE`;
      const response = await fetch(url, {
        method: "GET",
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      });
      const data = await response.json();

      if (!response.ok) {
        console.error("GET_ALL_USER_LIST_NEW Google Error:", data);
        throw new Error(
          data.error?.message || "Unable to read user master sheets",
        );
      }

      return data.values || [];
    }),
  );

  const [masterInput, otherInput] = sheetValues;
  const response = {
    data: [],
    isAdminAccess: false,
    role: "",
    name: "",
  };

  for (const row of masterInput) {
    if (!String(row[0] ?? "").trim()) break;

    if (row[3] === "Active") {
      const rowPassword = String(row[7] ?? "").trim().toLowerCase();

      if (password && rowPassword === password) {
        response.role = row[10] ?? "";
        response.name = row[6] ?? "";

        if (response.role === "Admin" || response.role === "Super Admin") {
          response.isAdminAccess = true;
        }
      }

      response.data.push({
        name: row[6] ?? "",
        mobile: row[9] ?? "",
        schemeDiscount: row[12] ?? "",
        devType: "NKDDevotee",
      });
    }
  }

  for (const row of otherInput) {
    if (!String(row[1] ?? "").trim()) break;

    response.data.push({
      name: row[1] ?? "",
      mobile: row[10] ?? "",
      devType: "Non-NKDDevotee",
      schemeDiscount: row[13] ?? "",
    });
  }

  response.data.sort((a, b) =>
    String(a.name).localeCompare(String(b.name)),
  );

  return response;
}
