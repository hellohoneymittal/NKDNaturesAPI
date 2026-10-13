const GOOGLE_SCOPES = [
  "https://www.googleapis.com/auth/spreadsheets",
  "https://www.googleapis.com/auth/drive",
  "https://www.googleapis.com/auth/calendar",
  "https://www.googleapis.com/auth/script.send_mail",
  "https://www.googleapis.com/auth/script.external_request",
];

export const GOOGLE_REDIRECT_URI =
  "https://natures-api.nkd-community-gzb.workers.dev/oauth/callback";

export async function READ_SECRET(env, name) {
  const value = env[name];

  if (typeof value === "string") {
    return value;
  }

  const binding = env[`${name}_STORE`] || value;

  if (binding && typeof binding.get === "function") {
    const value = await binding.get();

    if (typeof value === "string") {
      return value;
    }
  }

  if (binding == null) {
    return undefined;
  }

  throw new Error(
    `${name} must be a string or a readable Secrets Store binding`,
  );
}

// Google Authentication
export async function getGoogleAccessToken(env) {
  const [clientEmail, privateKey] = await Promise.all([
    READ_SECRET(env, "NATURES_EMAIL"),
    READ_SECRET(env, "NATURES_PRIVATE_KEY"),
  ]);

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
    scope:
      "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive",
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
  const binding = env.NATURES_PRIVATE_KEY;

  if (!binding) {
    return {
      status: false,
      message: "NATURES_PRIVATE_KEY is missing",
    };
  }

  const privateKey = await READ_SECRET(env, "NATURES_PRIVATE_KEY");

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

export async function getGoogleOAuthUrls(env, state) {
  const clientId = await READ_SECRET(env, "APP_CLIENT_ID");
  const clientSecret = await READ_SECRET(env, "APP_CLIENT_SECRET");

  const params = new URLSearchParams({
    client_id: clientId,
    redirect_uri: GOOGLE_REDIRECT_URI,
    response_type: "code",
    access_type: "offline",
    prompt: "select_account consent",
    scope: GOOGLE_SCOPES.join(" "),
    state,
  });

  return `https://accounts.google.com/o/oauth2/v2/auth?${params.toString()}`;
}

export function createOAuthState() {
  return crypto.randomUUID();
}

export async function getGoogleAccessTokenAppscript(env) {
  const clientId = await READ_SECRET(env, "APP_CLIENT_ID");
  const clientSecret = await READ_SECRET(env, "APP_CLIENT_SECRET");
  const refreshToken = await READ_SECRET(env, "APP_REFRESH_TOKEN");

  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    refresh_token: refreshToken,
    grant_type: "refresh_token",
  });

  const response = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body,
  });

  const data = await response.json();

  if (!response.ok || !data.access_token) {
    throw new Error(
      `Google access token refresh failed: ${JSON.stringify(data)}`,
    );
  }

  return data.access_token;
}

export async function exchangeGoogleCode(code, env) {
  const clientId = await READ_SECRET(env, "APP_CLIENT_ID");
  const clientSecret = await READ_SECRET(env, "APP_CLIENT_SECRET");

  const body = new URLSearchParams({
    code,
    client_id: clientId,
    client_secret: clientSecret,
    redirect_uri: GOOGLE_REDIRECT_URI,
    grant_type: "authorization_code",
  });

  const response = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body,
  });

  const data = await response.json();

  if (!response.ok) {
    throw new Error(`Google token exchange failed: ${JSON.stringify(data)}`);
  }

  return data;
}
