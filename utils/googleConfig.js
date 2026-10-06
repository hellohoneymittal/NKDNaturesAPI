// Google Sheet Configuration
export const SPREADSHEET_ID = "12RLRK6PjQVeysGskGu6Zanpx6AFU7QwMbU__Ec8JjWI";

// Google Authentication
export async function getGoogleAccessToken(env) {
  const clientEmail = env.NATURES_EMAIL;
  const privateKey = env.NATURES_PRIVATE_KEY;

  if (!clientEmail) {
    throw new Error("NATURES_EMAIL is missing");
  }

  if (!privateKey) {
    throw new Error("NATURES_PRIVATE_KEY is missing");
  }

  const formattedPrivateKey = privateKey
    .replace(/\\n/g, "\n")
    .replace(/\r/g, "")
    .trim();

  const header = {
    alg: "RS256",
    typ: "JWT",
  };

  const now = Math.floor(Date.now() / 1000);

  const payload = {
    iss: clientEmail,
    scope: "https://www.googleapis.com/auth/spreadsheets",
    aud: "https://oauth2.googleapis.com/token",
    iat: now,
    exp: now + 3600,
  };

  const encodedHeader = base64UrlEncode(JSON.stringify(header));
  const encodedPayload = base64UrlEncode(JSON.stringify(payload));

  const unsignedToken = encodedHeader + "." + encodedPayload;

  const privateKeyObject = await importPrivateKey(formattedPrivateKey);

  const signatureBuffer = await crypto.subtle.sign(
    {
      name: "RSASSA-PKCS1-v1_5",
    },
    privateKeyObject,
    new TextEncoder().encode(unsignedToken),
  );

  const signature = base64UrlEncodeBytes(new Uint8Array(signatureBuffer));

  const jwt = unsignedToken + "." + signature;

  const tokenResponse = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body:
      "grant_type=" +
      encodeURIComponent("urn:ietf:params:oauth:grant-type:jwt-bearer") +
      "&assertion=" +
      encodeURIComponent(jwt),
  });

  const tokenData = await tokenResponse.json();

  if (!tokenResponse.ok) {
    console.error("Google Token Error:", tokenData);

    throw new Error(
      tokenData.error_description ||
        tokenData.error ||
        "Unable to get Google access token",
    );
  }

  if (!tokenData.access_token) {
    throw new Error("Google access token was not returned");
  }

  return tokenData.access_token;
}

// Import Private Key

async function importPrivateKey(pem) {
  const pemContents = pem
    .replace(/-----BEGIN PRIVATE KEY-----/g, "")
    .replace(/-----END PRIVATE KEY-----/g, "")
    .replace(/\s/g, "");

  if (!pemContents) {
    throw new Error("NATURES_PRIVATE_KEY is empty after formatting");
  }

  let binaryDerString;

  try {
    binaryDerString = atob(pemContents);
  } catch (error) {
    throw new Error(
      "NATURES_PRIVATE_KEY is not a valid Base64 PKCS8 private key",
    );
  }

  const binaryDer = new Uint8Array(binaryDerString.length);

  for (let i = 0; i < binaryDerString.length; i++) {
    binaryDer[i] = binaryDerString.charCodeAt(i);
  }

  return crypto.subtle.importKey(
    "pkcs8",
    binaryDer.buffer,
    {
      name: "RSASSA-PKCS1-v1_5",
      hash: "SHA-256",
    },
    false,
    ["sign"],
  );
}

// Base64 URL Encoding

function base64UrlEncode(value) {
  const bytes = new TextEncoder().encode(value);
  return base64UrlEncodeBytes(bytes);
}

// Base64 URL Encode Bytes

function base64UrlEncodeBytes(bytes) {
  let binary = "";

  for (let i = 0; i < bytes.length; i++) {
    binary += String.fromCharCode(bytes[i]);
  }

  return btoa(binary)
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/, "");
}

// Test Google Private Key

export async function TEST_KEY(env) {
  const privateKey = env.NATURES_PRIVATE_KEY;

  if (!privateKey) {
    return {
      status: false,
      message: "NATURES_PRIVATE_KEY is missing",
    };
  }

  const key = String(privateKey);

  return {
    status: true,
    diagnostics: {
      length: key.length,
      startsWithBegin: key.includes("-----BEGIN PRIVATE KEY-----"),
      endsWithEnd: key.includes("-----END PRIVATE KEY-----"),
      containsLiteralSlashN: key.includes("\\n"),
      containsActualNewLine: key.includes("\n"),
      containsDots: key.includes("..."),
      startsWithQuote: key.startsWith('"'),
      endsWithQuote: key.endsWith('"'),
    },
  };
}
