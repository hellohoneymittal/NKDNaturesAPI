import { DurableObject } from "cloudflare:workers";
import { TEST_KEY } from "./utils/googleConfig.js";
import {
  GET_ALL_USER_LIST_NEW,
  GET_DATA,
  SAVE_DATA,
  DELETE_DATA,
  SEARCH_VOUCHER,
} from "./api/ApplicationMethod.js";
import { runLegacyApi } from "./api/LegacyApplicationMethods.js";
const GOOGLE_SCRIPT_ID =
  "1JdxzMnE6B6gJCji6HF-NDjVqdteZOQoHHnNjNkVRg1BRfZ0f4m-_5zh3";

const GOOGLE_REDIRECT_URI =
  "https://natures-api.nkd-community-gzb.workers.dev/oauth/callback";

const GOOGLE_SCOPES = "https://www.googleapis.com/auth/spreadsheets";

async function getGoogleOAuthUrls(env, state) {
  const clientId = await getGoogleClientId(env);

  const params = new URLSearchParams({
    client_id: clientId,
    redirect_uri: GOOGLE_REDIRECT_URI,
    response_type: "code",
    access_type: "offline",
    prompt: "select_account consent",
    scope: GOOGLE_SCOPES,
    state,
  });

  return `https://accounts.google.com/o/oauth2/v2/auth?${params.toString()}`;
}

async function getGoogleClientId(env) {
  if (env.NATURES_CLIENT_ID_STORE) {
    return await env.NATURES_CLIENT_ID_STORE.get();
  }

  return env.NATURES_CLIENT_ID;
}

async function getGoogleClientSecret(env) {
  if (env.NATURES_CLIENT_SECRET_STORE) {
    return await env.NATURES_CLIENT_SECRET_STORE.get();
  }

  return env.NATURES_CLIENT_SECRET;
}

function createOAuthState() {
  return crypto.randomUUID();
}

async function exchangeGoogleCode(code, env) {
  const clientId = await getGoogleClientId(env);
  const clientSecret = await getGoogleClientSecret(env);

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

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);

    const corsHeaders = {
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type",
    };

    // CORS preflight
    if (request.method === "OPTIONS") {
      return new Response(null, {
        status: 204,
        headers: corsHeaders,
      });
    }

    // OAuth start
    if (request.method === "GET" && url.pathname === "/oauth/start") {
      const state = createOAuthState();
      const oauthUrl = await getGoogleOAuthUrls(env, state);

      return new Response(null, {
        status: 302,
        headers: {
          Location: oauthUrl,
          "Set-Cookie": `oauth_state=${state}; HttpOnly; Secure; SameSite=Lax; Path=/; Max-Age=600`,
        },
      });
    }

    // OAuth callback
    if (request.method === "GET" && url.pathname === "/oauth/callback") {
      try {
        const code = url.searchParams.get("code");
        const returnedState = url.searchParams.get("state");
        const cookie = request.headers.get("Cookie") || "";

        const stateMatch = cookie.match(/oauth_state=([^;]+)/);
        const storedState = stateMatch?.[1];

        if (!code) {
          throw new Error("Authorization code missing");
        }

        if (!returnedState || !storedState || returnedState !== storedState) {
          throw new Error("Invalid OAuth state");
        }

        const tokenData = await exchangeGoogleCode(code, env);

        return new Response(
          JSON.stringify({
            status: true,
            message: "Google OAuth successful",
            accessTokenReceived: !!tokenData.access_token,
            refreshTokenReceived: !!tokenData.refresh_token,
            expiresIn: tokenData.expires_in,
            scope: tokenData.scope,
            refreshToken: tokenData.refresh_token || null,
          }),
          {
            status: 200,
            headers: {
              "Content-Type": "application/json",
              "Set-Cookie":
                "oauth_state=; HttpOnly; Secure; SameSite=Lax; Path=/; Max-Age=0",
            },
          },
        );
      } catch (error) {
        return new Response(
          JSON.stringify({
            status: false,
            message: error?.message || "OAuth callback failed",
          }),
          {
            status: 400,
            headers: {
              "Content-Type": "application/json",
            },
          },
        );
      }
    }

    // Health check
    if (request.method === "GET") {
      return new Response(
        JSON.stringify({
          status: true,
          message: "Cloudflare Worker is running",
          service: "Google Sheets API",
        }),
        {
          status: 200,
          headers: {
            ...corsHeaders,
            "Content-Type": "application/json",
          },
        },
      );
    }

    // Only POST APIs

    if (request.method !== "POST") {
      return new Response(
        JSON.stringify({
          status: false,
          message: "Only GET, POST and OPTIONS methods are allowed",
        }),
        {
          status: 405,
          headers: {
            ...corsHeaders,
            "Content-Type": "application/json",
          },
        },
      );
    }

    try {
      const requestData = await request.json();

      const apiType = requestData.apiType;
      const inputData = requestData.inputData || {};

      let response;

      // API routing

      switch (apiType) {
        case "GET_ALL_USER_LIST_NEW":
          response = await GET_ALL_USER_LIST_NEW(inputData, env);
          break;

        case "TEST_KEY":
          response = await TEST_KEY(env);
          break;

        case "SEARCH_VOUCHER":
          response = await SEARCH_VOUCHER(inputData, env);
          break;

        case "GET_DATA":
          response = await GET_DATA(inputData, env);
          break;

        case "UPDATE_STOCK": {
          const lockId = env.STOCK_UPDATE_LOCK.idFromName(
            "global-stock-update",
          );
          const lock = env.STOCK_UPDATE_LOCK.get(lockId);
          const lockResponse = await lock.fetch(
            "https://stock-update/execute",
            {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify(requestData),
            },
          );
          response = await lockResponse.json();
          break;
        }

        case "SAVE_DATA":
          response = await SAVE_DATA(inputData, env);
          break;

        case "DELETE_DATA":
          response = await DELETE_DATA(inputData, env);
          break;

        case "GET_STOCK":
        case "GET_ALL_USER_LIST":
        case "GET_PRODUCT_LIST":
        case "SAVE_PRODUCTION_DATA":
        case "CREATE_SALE":
        case "UPDATE_STOCK_VIA_SALE":
        case "UPDATE_ACTIVITY_MASTER":
        case "INSERT_DAILY_INPUT":
        case "ADD_NEW_USER":
        case "ADD_LIB_USER":
        case "LIB_BOOK_LIST":
        case "LIB_USER_LIST":
        case "LIB_ISSUE_BOOK":
        case "GET_KHATA_BOOK_USER_LIST":
        case "UPDATE_CUST_CREDIT_BALANCE":
        case "GET_KHATA_BOOK_BY_USER_ID":
        case "GET_USER_INFO_BY_PASSWORD":
        case "SAVE_USER_ORDER_DATA":
        case "GET_USER_ORDER_LIST":
        case "READY_USER_ORDER":
        case "CREATE_SALE_NKD":
        case "GENERATE_NATURES_GST_INVOICE":
          response = await runLegacyApi(apiType, requestData, env);
          break;

        default:
          response = {
            status: false,
            message: "Invalid apiType",
            apiType: apiType,
          };
          break;
      }

      response = {
        ...response,
        status: response?.status === false ? false : true,
      };

      return new Response(JSON.stringify(response), {
        status: 200,
        headers: {
          ...corsHeaders,
          "Content-Type": "application/json",
        },
      });
    } catch (error) {
      console.error("Worker Error:", error);

      return new Response(
        JSON.stringify({
          status: false,
          message: error?.message || "Internal server error",
        }),
        {
          status: 200,
          headers: {
            ...corsHeaders,
            "Content-Type": "application/json",
          },
        },
      );
    }
  },
};

export class StockUpdateCoordinator extends DurableObject {
  constructor(ctx, env) {
    super(ctx, env);
    this.queue = Promise.resolve();
  }

  async fetch(request) {
    const requestData = await request.json();
    const run = this.queue.then(() =>
      runLegacyApi("UPDATE_STOCK", requestData, this.env),
    );
    this.queue = run.then(
      () => undefined,
      () => undefined,
    );

    const result = await run;
    return Response.json(result);
  }
}
