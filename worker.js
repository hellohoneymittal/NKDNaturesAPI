import {
  exchangeGoogleCode,
  getGoogleOAuthUrls,
} from "./utils/googleConfig.js";

import { callAppsScript } from "./api/CommonMethod.js";

const APP_S_ID_MAP = {
  NATURES: "1JdxzMnE6B6gJCji6HF-NDjVqdteZOQoHHnNjNkVRg1BRfZ0f4m-_5zh3",
  NKD: "1JdxzMnE6B6gJCji6HF-NDjVqdteZOQoHHnNjNkVRg1BRfZ0f4m-_5zG3",
};

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);

    const corsHeaders = {
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type",
    };

    if (request.method === "OPTIONS") {
      return new Response(null, {
        status: 204,
        headers: corsHeaders,
      });
    }

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

    if (request.method === "GET" && url.pathname === "/oauth/callback") {
      try {
        const cookie = request.headers.get("Cookie") || "";
        const code = url.searchParams.get("code");
        const returnedState = url.searchParams.get("state");

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
              ...corsHeaders,
              "Content-Type": "application/json",
            },
          },
        );
      }
    }

    if (request.method === "GET" && url.pathname === "/test-apps-script") {
      try {
        const result = await runAppsScriptFunction(
          env,
          "TEST_EMAIL_PERMISSION",
        );

        return new Response(
          JSON.stringify({
            status: true,
            message: "Apps Script executed successfully",
            result,
          }),
          {
            status: 200,
            headers: {
              ...corsHeaders,
              "Content-Type": "application/json",
            },
          },
        );
      } catch (error) {
        return new Response(
          JSON.stringify({
            status: false,
            message: error?.message || "Apps Script execution failed",
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
    }

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

      const apiTypeMethodName = requestData.apiType;
      const inputData = requestData.inputData || {};
      const appName = String(inputData.appName || requestData.appName || "")
        .trim()
        .toUpperCase();
      const appScriptId = APP_S_ID_MAP[appName];

      let response;

      switch (apiTypeMethodName) {
        default: {
          if (!appScriptId) {
            response = {
              status: false,
              message:
                "Invalid or missing appName. Please provide a valid appName.",
              appName: appName || null,
              apiType: apiTypeMethodName,
            };
            break;
          }

          if (appName === "NATURES") {
            switch (apiTypeMethodName) {
              case "GET_ALL_USER_LIST":
              case "CREATE_SALE":
              case "GET_STOCK":
              case "GET_PRODUCT_LIST":
              case "SAVE_PRODUCTION_DATA":
              case "UPDATE_STOCK":
              case "GENERATE_NATURES_GST_INVOICE":
              case "GET_KHATA_BOOK_USER_LIST":
              case "UPDATE_CUST_CREDIT_BALANCE":
              case "GET_KHATA_BOOK_BY_USER_ID":
              case "INSERT_DAILY_INPUT":
              case "ADD_NEW_USER":

              case "UPDATE_STOCK_VIA_SALE": // NOT IS USE AS OF NOW
              case "UPDATE_ACTIVITY_MASTER": // not in use as of now

              case "ADD_LIB_USER":
              case "LIB_BOOK_LIST":
              case "LIB_USER_LIST":
              case "LIB_ISSUE_BOOK":

              case "GET_USER_INFO_BY_PASSWORD":
              case "SAVE_USER_ORDER_DATA":
              case "GET_USER_ORDER_LIST":
              case "READY_USER_ORDER":
              case "CREATE_SALE_NKD":
                response = await callAppsScript(
                  env,
                  appScriptId,
                  apiTypeMethodName,
                  inputData,
                );
                break;

              default:
                response = {
                  status: false,
                  message: `Invalid apiType for ${appName}`,
                  apiType: apiTypeMethodName,
                  appName,
                };
                break;
            }
          }
          if (appName === "ANOTHER") {
            switch (apiTypeMethodName) {
              case "testing":
                response = await callAppsScript(
                  env,
                  appScriptId,
                  apiTypeMethodName,
                  inputData,
                );
                break;

              default:
                response = {
                  status: false,
                  message: `Invalid apiType for ${appName}`,
                  apiType: apiTypeMethodName,
                  appName,
                };
                break;
            }
          }

          break;
        }
      }

      const isSuccess = response?.status !== false;
      const responseData = response ?? {};

      if (responseData && typeof responseData === "object") {
        delete responseData.status;
      }

      return new Response(
        JSON.stringify({
          status: isSuccess,
          data: responseData,
        }),
        {
          status: 200,
          headers: {
            ...corsHeaders,
            "Content-Type": "application/json",
          },
        },
      );
    } catch (error) {
      console.error("Worker Error:", error?.stack || error);

      return new Response(
        JSON.stringify({
          status: false,
          data: {
            message: error?.message || "Internal server error",
            error: {
              code: error?.code || 500,
              type: error?.errorType || "INTERNAL_ERROR",
              function: error?.functionName || null,
              lineNumber: error?.lineNumber ?? null,
              stackTrace: error?.stackTrace || [],
            },
          },
        }),
        {
          status: 500,
          headers: {
            ...corsHeaders,
            "Content-Type": "application/json",
          },
        },
      );
    }
  },
};
