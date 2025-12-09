import { findIndexByColumnsNameIn2DArray } from "../common.js";
import { appendError, appendLog, showStatus } from "../ui.js";

/**
 * @typedef {Object} ConfigFieldMap
 * @property {string} vertexAISearchAppId - The name of the config field for Vertex AI Search App ID.
 * @property {string} vertexAIProjectID - The name of the config field for Vertex AI Project ID.
 * @property {string} vertexAILocation - The name of the config field for Vertex AI Location.
 * @property {string} model - The name of the config field for Answer Model.
 * @property {string} preamble - The name of the config field for Preamble (Customized Summaries).
 * @property {string} summaryResultCount - The name of the config field for summaryResultCount (1-5).
 * @property {string} genereateGrounding - The name of the config field for Generate Grounding Score.
 * @property {string} useSemanticChunks - The name of the config field for useSemanticChunks (True or False).
 * @property {string} ignoreAdversarialQuery - The name of the config field for ignoreAdversarialQuery (True or False).
 * @property {string} ignoreNonSummarySeekingQuery - The name of the config field for ignoreNonSummarySeekingQuery (True or False).
 * @property {string} summaryMatchingAdditionalPrompt - The name of the config field for SummaryMatchingAdditionalPrompt.
 * @property {string} batchSize - The name of the config field for Batch Size (1-10).
 * @property {string} timeBetweenCallsInSec - The name of the config field for Time between Batches in Seconds (1-10).
 * @property {string} maxSnippetCount - The name of the config field for maxSnippetCount (1-5).
 * @property {Object} extractiveContentSpec - The name of the config field for extractiveContentSpec.
 * @property {string} extractiveContentSpec.maxExtractiveAnswerCount - The name of the config field for maxExtractiveAnswerCount (1-5).
 * @property {string} extractiveContentSpec.maxExtractiveSegmentCount - The name of the config field for maxExtractiveSegmentCount (1-5).
 */

/**
 * Map config fields to their corresponding row in vertex_ai_search_configValues
 * @type {ConfigFieldMap}
 */
const configFieldMap = {
  vertexAISearchAppId: "Vertex AI Search App ID",
  vertexAIProjectID: "Vertex AI Project ID",
  vertexAILocation: "Vertex AI Location",
  model: "Answer Model",
  preamble: "Preamble (Customized Summaries)",
  summaryResultCount: "summaryResultCount (1-5)",
  genereateGrounding: "Generate Grounding Score",
  useSemanticChunks: "useSemanticChunks (True or False)",
  ignoreAdversarialQuery: "ignoreAdversarialQuery (True or False)",
  ignoreNonSummarySeekingQuery: "ignoreNonSummarySeekingQuery (True or False)",
  summaryMatchingAdditionalPrompt: "SummaryMatchingAdditionalPrompt",
  batchSize: "Batch Size (1-10)",
  timeBetweenCallsInSec: "Time between Batches in Seconds (1-10)",
  maxSnippetCount: "maxSnippetCount (1-5)",
  extractiveContentSpec: {
    maxExtractiveAnswerCount: "maxExtractiveAnswerCount (1-5)",
    maxExtractiveSegmentCount: "maxExtractiveSegmentCount (1-5)",
  },
};

const answerConfigFieldMap = {
  vertexAISearchAppId: "Vertex AI Search App ID",
  vertexAIProjectID: "Vertex AI Project ID",
  vertexAILocation: "Vertex AI Location",
  model: "Answer Model",
  preamble: "Preamble (Customized Summaries)",
  ignoreAdversarialQuery: "ignoreAdversarialQuery (True or False)",
  ignoreNonAnswerSeekingQuery: "ignoreNonAnswerSeekingQuery (True or False)",
  ignoreLowRelevantContent: "ignoreLowRelevantContent (True or False)",
  includeGroundingSupports: "includeGroundingSupports (True or False)",
  includeCitations: "includeCitations (True or False)",
  summaryMatchingAdditionalPrompt: "SummaryMatchingAdditionalPrompt",
  batchSize: "Batch Size (1-10)",
  timeBetweenCallsInSec: "Time between Batches in Seconds (1-10)",
};

const synthQAFieldMap = {
  vertexAIProjectID: "Vertex AI Project ID",
  vertexAILocation: "Vertex AI Location",
  model: "Gemini Model ID",
  systemInstruction: "System Instructions",
  prompt: "Prompt",
  qaQualityFlag: "Generate Q & A Quality",
  qAQualityPrompt: "Q & A Quality Prompt",
  qAQualityModel: "Q & A Quality Model ID",
  batchSize: "Batch Size (1-10)",
  timeBetweenCallsInSec: "Time between Batches in Seconds (1-10)",
};

export function getColumn(table, columnName) {
  try {
    const column = table.columns.getItemOrNullObject(columnName);
    column.load();
    return column;
  } catch (error) {
    appendError("Error getColumn:", error);
    showStatus(`Exception when getting column: ${JSON.stringify(error)}`, true);
  }
}

export async function getSearchConfigFromActiveSheet(
  reportErrorTableNotFound = false,
  getAccessToken = false,
) {
  const config = await getConfigFromActiveSheet(
    buildSearchConfig,
    "TestCasesTable",
    reportErrorTableNotFound,
    getAccessToken,
  );

  if (config) {
    // Validate config
    const isValid =
      (config.extractiveContentSpec.maxExtractiveAnswerCount !== null) ^
      (config.extractiveContentSpec.maxExtractiveSegmentCount !== null) ^
      (config.maxSnippetCount !== null);

    if (!isValid) {
      // None, multiple, or all variables are non-null
      showStatus(
        `Error in executeSearchTests: Only one of the maxExtractiveAnswerCount, maxExtractiveSegmentCount, or maxSnippetCount should be set to a non-zero value`,
        true,
      );
      return null;
    }
  }
  return config;
}

export async function getAnswerConfigFromActiveSheet(
  reportErrorTableNotFound = false,
  getAccessToken = false,
) {
  return getConfigFromActiveSheet(
    buildAnswerConfig,
    "TestCasesTable",
    reportErrorTableNotFound,
    getAccessToken,
  );
}

export async function getSyntheticQAConfigFromActiveSheet(
  reportErrorTableNotFound = false,
  getAccessToken = false,
) {
  return getConfigFromActiveSheet(
    buildSynthQAConfig,
    "SyntheticQATable",
    reportErrorTableNotFound,
    getAccessToken,
  );
}



async function getConfigFromActiveSheet(
  fn_buildConfig,
  typeOfTable,
  reportErrorTableNotFound,
  getAccessToken,
) {
  var config = null;
  await Excel.run(async (context) => {
    try {
      const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      currentWorksheet.load("name");
      await context.sync();
      const worksheetName = currentWorksheet.name;

      // check if we config  table is there
      const configTable = currentWorksheet.tables.getItemOrNullObject(
        `${worksheetName}.ConfigTable`,
      );
      configTable.load();

      // check if we search data table is there
      const searchTable = currentWorksheet.tables.getItemOrNullObject(
        `${worksheetName}.${typeOfTable}`,
      );
      searchTable.load();
      await context.sync();

      // if both tables are found, then the current sheet is of same type as config being fetched
      // So return the current config
      if (configTable.isNullObject === false && searchTable.isNullObject === false) {
        const valueColumn = getColumn(configTable, "Value");
        const configColumn = getColumn(configTable, "Config");
        await context.sync();

        config = fn_buildConfig(config, valueColumn, configColumn);

        if (getAccessToken) {
          config.accessToken = $("#access-token").val();
        }
        config.originalWorksheetName = worksheetName;
      } else if (reportErrorTableNotFound) {
        var message = "Error in ${worksheetName}: ";
        if (searchTable.isNullObject) {
          message += `${typeOfTable} not found in current sheet. Make sure you are in the right sheet`;
        }
        if (configTable.isNullObject) {
          message += `ConfigTable not found in current sheet. Make sure you are in the right sheet`;
        }
        appendLog(message);
        showStatus(message, true);
      }
    } catch (error) {
      const message = `Caught Exception in getSearchConfigFromActiveSheet `;
      appendError(message, error);
      showStatus(
        `Caught Exception in getSearchConfigFromActiveSheet: ${error}. Trace: ${error.stack}`,
        true,
      );
    }
  });

  return config;
}

function buildSearchConfig(config, valueColumn, configColumn) {
  config = {
    vertexAISearchAppId:
      valueColumn.values[
        findIndexByColumnsNameIn2DArray(configColumn.values, configFieldMap.vertexAISearchAppId)
      ][0],
    vertexAIProjectID:
      valueColumn.values[
        findIndexByColumnsNameIn2DArray(configColumn.values, configFieldMap.vertexAIProjectID)
      ][0],
    vertexAILocation:
      valueColumn.values[
        findIndexByColumnsNameIn2DArray(configColumn.values, configFieldMap.vertexAILocation)
      ][0],
    extractiveContentSpec: {
      maxExtractiveAnswerCount:
        valueColumn.values[
          findIndexByColumnsNameIn2DArray(
            configColumn.values,
            configFieldMap.extractiveContentSpec.maxExtractiveAnswerCount,
          )
        ][0] === 0
          ? null
          : valueColumn.values[
              findIndexByColumnsNameIn2DArray(
                configColumn.values,
                configFieldMap.extractiveContentSpec.maxExtractiveAnswerCount,
              )
            ][0],
      maxExtractiveSegmentCount:
        valueColumn.values[
          findIndexByColumnsNameIn2DArray(
            configColumn.values,
            configFieldMap.extractiveContentSpec.maxExtractiveSegmentCount,
          )
        ][0] === 0
          ? null
          : valueColumn.values[
              findIndexByColumnsNameIn2DArray(
                configColumn.values,
                configFieldMap.extractiveContentSpec.maxExtractiveSegmentCount,
              )
            ][0],
    },
    maxSnippetCount:
      valueColumn.values[
        findIndexByColumnsNameIn2DArray(configColumn.values, configFieldMap.maxSnippetCount)
      ][0] === 0
        ? null
        : valueColumn.values[
            findIndexByColumnsNameIn2DArray(configColumn.values, configFieldMap.maxSnippetCount)
          ][0],
    preamble:
      valueColumn.values[
        findIndexByColumnsNameIn2DArray(configColumn.values, configFieldMap.preamble)
      ][0],
    model:
      valueColumn.values[
        findIndexByColumnsNameIn2DArray(configColumn.values, configFieldMap.model)
      ][0],
    summaryResultCount:
      valueColumn.values[
        findIndexByColumnsNameIn2DArray(configColumn.values, configFieldMap.summaryResultCount)
      ][0],
    genereateGrounding:
      valueColumn.values[
        findIndexByColumnsNameIn2DArray(configColumn.values, configFieldMap.genereateGrounding)
      ][0],
    useSemanticChunks:
      valueColumn.values[
        findIndexByColumnsNameIn2DArray(configColumn.values, configFieldMap.useSemanticChunks)
      ][0],
    ignoreAdversarialQuery:
      valueColumn.values[
        findIndexByColumnsNameIn2DArray(configColumn.values, configFieldMap.ignoreAdversarialQuery)
      ][0],
    ignoreNonSummarySeekingQuery:
      valueColumn.values[
        findIndexByColumnsNameIn2DArray(
          configColumn.values,
          configFieldMap.ignoreNonSummarySeekingQuery,
        )
      ][0],
    summaryMatchingAdditionalPrompt:
      valueColumn.values[
        findIndexByColumnsNameIn2DArray(
          configColumn.values,
          configFieldMap.summaryMatchingAdditionalPrompt,
        )
      ][0],

    batchSize: parseInt(
      valueColumn.values[
        findIndexByColumnsNameIn2DArray(configColumn.values, configFieldMap.batchSize)
      ][0],
    ),
    timeBetweenCallsInSec: parseInt(
      valueColumn.values[
        findIndexByColumnsNameIn2DArray(configColumn.values, configFieldMap.timeBetweenCallsInSec)
      ][0],
    ),
  };
  return config;
}

function buildAnswerConfig(config, valueColumn, configColumn) {
  config = {
    vertexAISearchAppId:
      valueColumn.values[
        findIndexByColumnsNameIn2DArray(
          configColumn.values,
          answerConfigFieldMap.vertexAISearchAppId,
        )
      ][0],
    vertexAIProjectID:
      valueColumn.values[
        findIndexByColumnsNameIn2DArray(configColumn.values, answerConfigFieldMap.vertexAIProjectID)
      ][0],
    vertexAILocation:
      valueColumn.values[
        findIndexByColumnsNameIn2DArray(configColumn.values, answerConfigFieldMap.vertexAILocation)
      ][0],
    preamble:
      valueColumn.values[
        findIndexByColumnsNameIn2DArray(configColumn.values, answerConfigFieldMap.preamble)
      ][0],
    model:
      valueColumn.values[
        findIndexByColumnsNameIn2DArray(configColumn.values, answerConfigFieldMap.model)
      ][0],

    ignoreAdversarialQuery:
      valueColumn.values[
        findIndexByColumnsNameIn2DArray(
          configColumn.values,
          answerConfigFieldMap.ignoreAdversarialQuery,
        )
      ][0],
    ignoreNonAnswerSeekingQuery:
      valueColumn.values[
        findIndexByColumnsNameIn2DArray(
          configColumn.values,
          answerConfigFieldMap.ignoreNonAnswerSeekingQuery,
        )
      ][0],
    ignoreLowRelevantContent:
      valueColumn.values[
        findIndexByColumnsNameIn2DArray(
          configColumn.values,
          answerConfigFieldMap.ignoreLowRelevantContent,
        )
      ][0],
    includeGroundingSupports:
      valueColumn.values[
        findIndexByColumnsNameIn2DArray(
          configColumn.values,
          answerConfigFieldMap.includeGroundingSupports,
        )
      ][0],
    includeCitations:
      valueColumn.values[
        findIndexByColumnsNameIn2DArray(configColumn.values, answerConfigFieldMap.includeCitations)
      ][0],
    summaryMatchingAdditionalPrompt:
      valueColumn.values[
        findIndexByColumnsNameIn2DArray(
          configColumn.values,
          answerConfigFieldMap.summaryMatchingAdditionalPrompt,
        )
      ][0],

    batchSize: parseInt(
      valueColumn.values[
        findIndexByColumnsNameIn2DArray(configColumn.values, answerConfigFieldMap.batchSize)
      ][0],
    ),
    timeBetweenCallsInSec: parseInt(
      valueColumn.values[
        findIndexByColumnsNameIn2DArray(
          configColumn.values,
          answerConfigFieldMap.timeBetweenCallsInSec,
        )
      ][0],
    ),
  };
  return config;
}

export async function buildSynthQAConfig(config, valueColumn, configColumn) {
  config = {
    vertexAIProjectID:
      valueColumn.values[
        findIndexByColumnsNameIn2DArray(configColumn.values, synthQAFieldMap.vertexAIProjectID)
      ][0],
    vertexAILocation:
      valueColumn.values[
        findIndexByColumnsNameIn2DArray(configColumn.values, synthQAFieldMap.vertexAILocation)
      ][0],
    model:
      valueColumn.values[
        findIndexByColumnsNameIn2DArray(configColumn.values, synthQAFieldMap.model)
      ][0],
    systemInstruction:
      valueColumn.values[
        findIndexByColumnsNameIn2DArray(configColumn.values, synthQAFieldMap.systemInstruction)
      ][0],
    prompt:
      valueColumn.values[
        findIndexByColumnsNameIn2DArray(configColumn.values, synthQAFieldMap.prompt)
      ][0],
    qaQualityFlag:
      valueColumn.values[
        findIndexByColumnsNameIn2DArray(configColumn.values, synthQAFieldMap.qaQualityFlag)
      ][0],
    qAQualityPrompt:
      valueColumn.values[
        findIndexByColumnsNameIn2DArray(configColumn.values, synthQAFieldMap.qAQualityPrompt)
      ][0],
    qAQualityModel:
      valueColumn.values[
        findIndexByColumnsNameIn2DArray(configColumn.values, synthQAFieldMap.qAQualityModel)
      ][0],
    batchSize: parseInt(
      valueColumn.values[
        findIndexByColumnsNameIn2DArray(configColumn.values, synthQAFieldMap.batchSize)
      ][0],
    ),
    timeBetweenCallsInSec: parseInt(
      valueColumn.values[
        findIndexByColumnsNameIn2DArray(configColumn.values, synthQAFieldMap.timeBetweenCallsInSec)
      ][0],
    ),
  };
  return config;
}


/**
 * Populates the vertex_ai_search_configValues array with values from the config object.
 *
 * @param {Array<Array<string>>} vertex_ai_search_configValues - The 2D array representing the config table.
 * @param {object} config - The config object containing the values to populate.
 */
export function populateConfigTableValues(vertex_ai_search_configValues, config) {
  if (!vertex_ai_search_configValues || !config) {
    console.error("Error: vertex_ai_search_configValues or config is null or undefined.");
    return;
  }

  // Iterate through the configFieldMap and update vertex_ai_search_configValues
  for (const configKey in configFieldMap) {
    const configValue = config[configKey];
    const configRowName = configFieldMap[configKey];

    if (configKey === "extractiveContentSpec") {
      for (const subConfigKey in configValue) {
        const subConfigValue = configValue[subConfigKey];
        const subConfigRowName = configFieldMap[configKey][subConfigKey];
        const rowIndex = findIndexByColumnsNameIn2DArray(
          vertex_ai_search_configValues,
          subConfigRowName,
        );
        if (rowIndex !== -1) {
          vertex_ai_search_configValues[rowIndex][1] = subConfigValue === null ? 0 : subConfigValue;
        }
      }
    } else {
      const rowIndex = findIndexByColumnsNameIn2DArray(
        vertex_ai_search_configValues,
        configRowName,
      );
      if (rowIndex !== -1) {
        vertex_ai_search_configValues[rowIndex][1] = configValue === null ? 0 : configValue;
      }
    }
  }
}
