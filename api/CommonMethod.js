import { getGoogleAccessTokenAppscript } from "../utils/googleConfig";

export function formatApiResponse(response) {
  const isSuccess = response?.status !== false;
  const responseData = response ?? {};

  if (responseData && typeof responseData === "object") {
    delete responseData.status;
  }

  return {
    status: isSuccess,
    data: responseData,
  };
}

export async function runAppsScriptFunction(
  env,
  scriptId,
  functionName,
  parameters = [],
) {
  try {
    if (!scriptId || typeof scriptId !== "string") {
      const error = new Error("Valid Google Apps Script ID is required");
      error.code = 400;
      error.errorType = "INVALID_SCRIPT_ID";
      throw error;
    }

    if (!functionName || typeof functionName !== "string") {
      const error = new Error("Valid Apps Script function name is required");
      error.code = 400;
      error.errorType = "INVALID_FUNCTION_NAME";
      throw error;
    }

    const accessToken = await getGoogleAccessTokenAppscript(env);

    const response = await fetch(
      `https://script.googleapis.com/v1/scripts/${scriptId}:run`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          function: functionName,
          parameters,
        }),
      },
    );

    const responseText = await response.text();
    let data;

    try {
      data = JSON.parse(responseText);
    } catch {
      const error = new Error(
        `Invalid JSON response from Apps Script API (HTTP ${response.status}): ${responseText.slice(0, 1000)}`,
      );

      error.code = response.status;
      error.errorType = "INVALID_JSON";
      throw error;
    }

    const apiError = data.error;

    const executionError = apiError?.details?.find((detail) =>
      detail["@type"]?.includes("ExecutionError"),
    );

    if (!response.ok || apiError || executionError) {
      const actualError = executionError || apiError;
      const stackElements = actualError?.scriptStackTraceElements || [];
      const stackElement = stackElements[0];

      const error = new Error(
        actualError?.errorMessage ||
          apiError?.message ||
          `Apps Script execution failed: ${functionName}`,
      );

      error.code = apiError?.code || response.status;
      error.errorType = actualError?.errorType || "API_ERROR";
      error.functionName = stackElement?.function || functionName;
      error.lineNumber = stackElement?.lineNumber ?? null;
      error.stackTrace = stackElements;

      throw error;
    }

    if (
      !data.response ||
      data.response.result === undefined ||
      data.response.result === null
    ) {
      const error = new Error(
        `Apps Script returned no result for function: ${functionName}`,
      );

      error.code = 502;
      error.errorType = "MISSING_RESULT";
      error.functionName = functionName;

      throw error;
    }

    return data.response.result;
  } catch (error) {
    console.error(
      `runAppsScriptFunction error (${functionName}):`,
      JSON.stringify({
        scriptId,
        message: error?.message,
        code: error?.code,
        errorType: error?.errorType,
        functionName: error?.functionName || functionName,
        lineNumber: error?.lineNumber,
        stackTrace: error?.stackTrace,
        stack: error?.stack,
      }),
    );

    throw error;
  }
}

export async function callAppsScript(
  env,
  scriptId,
  functionName,
  inputData = [],
) {
  try {
    if (typeof inputData === "string") {
      if (inputData.trim() === "") {
        inputData = [];
      } else {
        try {
          inputData = JSON.parse(inputData);
        } catch {
          throw new Error("Invalid JSON: inputData must be valid JSON");
        }
      }
    }

    if (inputData === null || inputData === undefined) {
      inputData = [];
    }
    inputData = [inputData]; // it is required always.

    return await runAppsScriptFunction(env, scriptId, functionName, inputData);
  } catch (error) {
    console.error(
      `callAppsScript error (${functionName}):`,
      JSON.stringify({
        message: error?.message,
        code: error?.code,
        errorType: error?.errorType,
        functionName: error?.functionName || functionName,
        lineNumber: error?.lineNumber,
        stackTrace: error?.stackTrace,
      }),
    );

    throw error;
  }
}
