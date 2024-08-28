import {
  NotAuthenticatedError,
  PermissionDeniedError,
  QuotaError,
  ResourceNotFoundError,
  summaryMatching_examples,
  summaryMatching_prompt,
  VelvetError,
} from "./common.js";
import { appendLog } from "./ui.js";

export async function callVertexAISearch(id, query, config) {
  const token = config.accessToken;
  const preamble = config.preamble;
  const model =
    config.model === "" ? "gemini-1.0-pro-002/answer_gen/v1" : config.model;
  const summaryResultCount = config.summaryResultCount;
  const extractiveContentSpec =
    config.extractiveContentSpec === null ? {} : config.extractiveContentSpec;
  const snippetSpec = config.snippetSpec === null ? {} : config.snippetSpec;
  const useSemanticChunks = config.useSemanticChunks;
  const ignoreAdversarialQuery = config.ignoreAdversarialQuery;
  const ignoreNonSummarySeekingQuery = config.ignoreNonSummarySeekingQuery;
  const projectNumber = config.vertexAISearchProjectNumber;
  const datastoreName = config.vertexAISearchDataStoreName;

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
        ignoreNonSummarySeekingQuery:
          `${ignoreNonSummarySeekingQuery}`.toLowerCase(),
        modelPromptSpec: {
          preamble: `${preamble}`,
        },
        modelSpec: {
          version: `${model}`,
        },
      },
    },
  };

  const url = `https://discoveryengine.googleapis.com/v1alpha/projects/${projectNumber}/locations/global/collections/default_collection/dataStores/${datastoreName}/servingConfigs/default_search:search`;

  const { status, json_output } = await callVertexAI(url, token, data, id);

  appendLog(`testCaseID: ${id}: Search Query Finished Successfully`);

  return { id: id, status_code: status, output: json_output };
}

export async function calculateSimilarityUsingPalm2(
  id,
  sentence1,
  sentence2,
  config
) {
  const token = config.accessToken;
  const projectId = config.vertexAIProjectID;
  const location = config.vertexAILocation;
  const summaryMatchingAdditionalPrompt =
    config.summaryMatchingAdditionalPrompt === null
      ? ""
      : config.summaryMatchingAdditionalPrompt;

  var prompt =
    summaryMatching_prompt +
    summaryMatchingAdditionalPrompt +
    summaryMatching_examples;

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

export async function callVertexAI(url, token, data, id) {
  const response = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(data),
  });

  if (!response.ok) {
    if (response.status === 401) {
      const json = await response.json();
      throw new NotAuthenticatedError(id, json.error.message);
    } else if (response.status === 429) {
      const json = await response.json();
      throw new QuotaError(id, json.error.message);
    } else if (response.status === 403) {
      const json = await response.json();
      throw new PermissionDeniedError(id, json.error.message);
    } else if (response.status === 404) {
      const json = await response.json();
      throw new ResourceNotFoundError(id, json.error.message);
    } else {
      throw new VelvetError(
        id,
        `Error calling VertexAI for Summary, HTTP Code: ${
          response.status
        } Reason: ${JSON.stringify(response.body)}`
      );
    }
  } else {
    const json = await response.json();

    return { status: response.status, json_output: json };
  }
}

export async function callGeminiMultitModal(
  id,
  prompt,
  fileUri,
  mimeType,
  config
) {
  const token = config.accessToken;
  const projectId = config.vertexAIProjectID;
  const location = config.vertexAILocation;
  const model_id = config.model;
  const system_instruction =
    config.systemInstruction === null ? "" : config.systemInstruction;

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
      response_mime_type: `${config.responseMimeType}`,
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
        response_mime_type: `${config.responseMimeType}`,
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
  const output = json_output.candidates[0].content.parts[0].text;

  appendLog(`callGeminiMultitModal: Finished Successfully.`);

  return { id: id, status_code: status, output: `${output}` };
}
