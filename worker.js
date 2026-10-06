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

export default {
  async fetch(request, env, ctx) {
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
          const lockId = env.STOCK_UPDATE_LOCK.idFromName("global-stock-update");
          const lock = env.STOCK_UPDATE_LOCK.get(lockId);
          const lockResponse = await lock.fetch("https://stock-update/execute", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(requestData),
          });
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
