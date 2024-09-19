import {
  NotAuthenticatedError,
  PermissionDeniedError,
  QuotaError,
  ResourceNotFoundError,
  summaryMatching_examples,
  summaryMatching_prompt,
  VertexAIError,
} from "./common.js";
import { appendLog } from "./ui.js";

export async function callVertexAISearch(id, query, config) {
  const token = config.accessToken;
  const preamble = config.preamble;
  const model = config.model === "" ? "gemini-1.0-pro-002/answer_gen/v1" : config.model;
  const summaryResultCount = config.summaryResultCount;
  const extractiveContentSpec =
    config.extractiveContentSpec === null ? {} : config.extractiveContentSpec;
  const snippetSpec = config.snippetSpec === null ? {} : config.snippetSpec;
  const useSemanticChunks = config.useSemanticChunks;
  const ignoreAdversarialQuery = config.ignoreAdversarialQuery;
  const ignoreNonSummarySeekingQuery = config.ignoreNonSummarySeekingQuery;
  const projectNumber = config.vertexAISearchProjectNumber;
  const searchAppId = config.vertexAISearchAppId;

  var data = {
    query: query,
    page_size: "5",
    offset: 0,
    contentSearchSpec: {
      extractiveContentSpec,
      snippetSpec,
      summarySpec: {
        useSemanticChunks: `${useSemanticChunks}`.toLowerCase(),
        summaryResultCount: `${summaryResultCount}`,
        ignoreAdversarialQuery: `${ignoreAdversarialQuery}`.toLowerCase(),
        ignoreNonSummarySeekingQuery: `${ignoreNonSummarySeekingQuery}`.toLowerCase(),
        modelPromptSpec: {
          preamble: `${preamble}`,
        },
        modelSpec: {
          version: `${model}`,
        },
      },
    },
  };

  const url = `https://discoveryengine.googleapis.com/v1alpha/projects/${projectNumber}/locations/global/collections/default_collection/engines/${searchAppId}/servingConfigs/default_search:search`;

  const { status, json_output } = await callVertexAI(url, token, data, id);

  appendLog(`testCaseID: ${id}: Search Query Finished Successfully`);

  return { id: id, status_code: status, output: json_output };
}

export async function calculateSimilarityUsingPalm2(id, sentence1, sentence2, config) {
  appendLog(`testCaseID: ${id}: SummaryMatch Started `);

  const token = config.accessToken;
  const projectId = config.vertexAIProjectID;
  const location = config.vertexAILocation;
  const summaryMatchingAdditionalPrompt =
    config.summaryMatchingAdditionalPrompt === null ? "" : config.summaryMatchingAdditionalPrompt;

  var prompt = summaryMatching_prompt + summaryMatchingAdditionalPrompt + summaryMatching_examples;

  var full_prompt = `${prompt} answer_1: ${sentence1} answer_2: ${sentence2} output:`;

  var data = {
    instances: [{ prompt: `${full_prompt}` }],
    parameters: {
      temperature: 0.2,
      maxOutputTokens: 256,
      topK: 40,
      topP: 0.95,
      logprobs: 2,
    },
  };

  const url = `https://${location}-aiplatform.googleapis.com/v1/projects/${projectId}/locations/${location}/publishers/google/models/text-bison:predict`;

  const { status, json_output } = await callVertexAI(url, token, data, id);
  const output = json_output.predictions[0].content;

  appendLog(`testCaseID: ${id}: SummaryMatch Finished Successfully `);

  return { id: id, status_code: status, output: `${output}` };
}

export async function callGeminiMultitModal(
  id,
  prompt,
  systemInstruction,
  fileUri,
  mimeType,
  model_id,
  responseMimeType,
  config,
) {
  const token = config.accessToken;
  const projectId = config.vertexAIProjectID;
  const location = config.vertexAILocation;
  const system_instruction = systemInstruction === null ? "" : systemInstruction;
  const local_responseMimeType = responseMimeType ? responseMimeType : "text/plain";
  var data = {
    contents: [
      {
        role: "user",
        parts: [
          {
            text: `${prompt}`,
          },
          {
            fileData: {
              mimeType: `${mimeType}`,
              fileUri: `${fileUri}`,
            },
          },
        ],
      },
    ],
    systemInstruction: {
      parts: [
        {
          text: `${system_instruction}`,
        },
      ],
    },
    generationConfig: {
      maxOutputTokens: 8192,
      temperature: 1,
      topP: 0.95,
      response_mime_type: `${local_responseMimeType}`,
    },
    safetySettings: [
      {
        category: "HARM_CATEGORY_HATE_SPEECH",
        threshold: "BLOCK_MEDIUM_AND_ABOVE",
      },
      {
        category: "HARM_CATEGORY_DANGEROUS_CONTENT",
        threshold: "BLOCK_MEDIUM_AND_ABOVE",
      },
      {
        category: "HARM_CATEGORY_SEXUALLY_EXPLICIT",
        threshold: "BLOCK_MEDIUM_AND_ABOVE",
      },
      {
        category: "HARM_CATEGORY_HARASSMENT",
        threshold: "BLOCK_MEDIUM_AND_ABOVE",
      },
    ],
  };
  if (fileUri === null) {
    data = {
      contents: [
        {
          role: "user",
          parts: [
            {
              text: `${prompt}`,
            },
          ],
        },
      ],
      systemInstruction: {
        parts: [
          {
            text: `${system_instruction}`,
          },
        ],
      },
      generationConfig: {
        maxOutputTokens: 8192,
        temperature: 1,
        topP: 0.95,
        response_mime_type: `${local_responseMimeType}`,
      },
      safetySettings: [
        {
          category: "HARM_CATEGORY_HATE_SPEECH",
          threshold: "BLOCK_MEDIUM_AND_ABOVE",
        },
        {
          category: "HARM_CATEGORY_DANGEROUS_CONTENT",
          threshold: "BLOCK_MEDIUM_AND_ABOVE",
        },
        {
          category: "HARM_CATEGORY_SEXUALLY_EXPLICIT",
          threshold: "BLOCK_MEDIUM_AND_ABOVE",
        },
        {
          category: "HARM_CATEGORY_HARASSMENT",
          threshold: "BLOCK_MEDIUM_AND_ABOVE",
        },
      ],
    };
  }

  const url = `https://${location}-aiplatform.googleapis.com/v1/projects/${projectId}/locations/${location}/publishers/google/models/${model_id}:generateContent`;

  const { status, json_output } = await callVertexAI(url, token, data, id);
  //const output = json_output.candidates[0].content.parts[0].text;

  appendLog(`callGeminiMultitModal: Finished Successfully.`);

  return { id: id, status_code: status, output: json_output };
}
export async function callCheckGrounding(config, answerCandidate, factsArray, id) {
  const token = config.accessToken;
  const projectId = config.vertexAIProjectID;

  var payload = {
    answerCandidate: `${answerCandidate}`,
    facts: factsArray,
    groundingSpec: {
      citationThreshold: "0.6",
    },
  };

  const url = `https://discoveryengine.googleapis.com/v1/projects/${projectId}/locations/global/groundingConfigs/default_grounding_config:check`;

  const { status, json_output } = await callVertexAI(url, token, payload, id);

  appendLog(`checkGrounding: Finished Successfully.`);

  return { id: id, status_code: status, output: json_output };
}

export async function createSearchEvalSampleQuerySetId(config, querySetId, querySetDisplayName) {
  const token = config.accessToken;
  const projectId = config.vertexAIProjectID;

  var payload = {
    displayName: `${querySetDisplayName}`,
  };

  const url = `https://discoveryengine.googleapis.com/v1beta/projects/${projectId}/locations/global/sampleQuerySets?sampleQuerySetId=${querySetId}`;

  const { status, json_output } = await callVertexAI(url, token, payload, id);

  appendLog(`createSearchEvalQuerySet: Finished Successfully.`);

  return { status_code: status, output: json_output };
}

export async function createSearchEvalImportSampleQueryDataset(config, querySetId, queryEntry) {
  const token = config.accessToken;
  const projectId = config.vertexAIProjectID;

  var payload = {
    inlineSource: {
      sampleQueries: [
        {
          queryEntry,
        },
      ],
    },
  };

  const url = `https://discoveryengine.googleapis.com/v1beta/projects/${projectId}/locations/global/sampleQuerySets/${querySetId}/sampleQueries:import`;

  const { status, json_output } = await callVertexAI(url, token, payload);

  appendLog(`createSearchEvalImportSampleQueryDataset: Finished Successfully.`);

  return { status_code: status, output: json_output };
}

export async function createSearchEvalSubmitEvalJob(config, querySetId, searchAppID) {
  const token = config.accessToken;
  const projectId = config.vertexAIProjectID;

  var payload = {
    evaluationSpec: {
      querySetSpec: {
        sampleQuerySet: `projects/${projectId}/locations/global/sampleQuerySets/${querySetId}`,
      },
      searchRequest: {
        servingConfig: `projects/${projectId}/locations/global/collections/default_collection/engines/${searchAppID}/servingConfigs/default_search`,
      },
    },
  };

  const url = `https://discoveryengine.googleapis.com/v1beta/projects/${projectId}/locations/global/evaluations`;

  const { status, json_output } = await callVertexAI(url, token, payload);

  appendLog(`createSearchEvalRunEval: Finished Successfully.`);

  return { status_code: status, output: json_output };
}

export async function getSearchEvalResults(config, evaluationId, checkStatus) {
  const token = config.accessToken;
  const projectId = config.vertexAIProjectID;

  var payload = {
    evaluationSpec: {
      querySetSpec: {
        sampleQuerySet: `projects/${projectId}/locations/global/sampleQuerySets/${querySetId}`,
      },
      searchRequest: {
        servingConfig: `projects/${projectId}/locations/global/collections/default_collection/engines/${searchAppID}/servingConfigs/default_search`,
      },
    },
  };

  var url = `https://discoveryengine.googleapis.com/v1beta/projects/${projectId}/locations/global/evaluations/${evaluationId}:listResults`;

  if (checkStatus) {
    url = `https://discoveryengine.googleapis.com/v1beta/projects/${projectId}/locations/global/evaluations/${evaluationId}`;
  }

  const { status, json_output } = await callVertexAI(url, token, payload);

  appendLog(`getSearchEvalResults: Finished Successfully.`);

  return { status_code: status, output: json_output };
}

export async function callVertexAI(url, token, data, id) {
  try {
    const response = await fetch(url, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(data),
    });

    const json = await response.json();

    if (!response.ok) {
      let errorMessage = "Unknown Error";
      if (json.hasOwnProperty("error")) {
        errorMessage = json.error.message;
      }

      switch (response.status) {
        case 401:
          throw new NotAuthenticatedError(id, errorMessage);
        case 429:
          throw new QuotaError(id, errorMessage);
        case 403:
          throw new PermissionDeniedError(id, errorMessage);
        case 404:
          throw new ResourceNotFoundError(id, errorMessage);
        default:
          throw new VertexAIError(
            id,
            `Error calling VertexAI. HTTP Code: ${response.status} reason: ${errorMessage}`,
          );
      }
    }

    return { status: response.status, json_output: json };
  } catch (error) {
    // Handle network errors or unexpected errors
    if (
      error instanceof NotAuthenticatedError ||
      error instanceof QuotaError ||
      error instanceof PermissionDeniedError ||
      error instanceof ResourceNotFoundError ||
      error instanceof VertexAIError
    ) {
      // Known errors, rethrow them
      throw error;
    } else {
      // Unknown errors, wrap them in a VertexAIError
      throw new VertexAIError(id, `Network or unexpected error: ${error.message}`);
    }
  }
}
