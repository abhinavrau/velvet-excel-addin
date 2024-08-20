
import { NotAuthenticatedError, QuotaError, summaryMatching_examples, summaryMatching_prompt } from "./common.js";
import { appendLog } from "./ui.js";

export async function callVertexAISearch(testCaseRowNum, query, config) {

    const token = config.accessToken;
    const preamble = config.preamble;
    const model = config.model === "" ? 'gemini-1.0-pro-002/answer_gen/v1' : config.model;
    const summaryResultCount = config.summaryResultCount;
    const extractiveContentSpec = config.extractiveContentSpec === null ? {} : config.extractiveContentSpec;
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
                ignoreNonSummarySeekingQuery: `${ignoreNonSummarySeekingQuery}`.toLowerCase(),
                modelPromptSpec: {
                    preamble: `${preamble}`
                },
                modelSpec: {
                    version: `${model}`

                }
            },
        }
    };

    const url = `https://discoveryengine.googleapis.com/v1alpha/projects/${projectNumber}/locations/global/collections/default_collection/dataStores/${datastoreName}/servingConfigs/default_search:search`;

    const { status, json_output } = await callVettexAI(url, token, data);
    
    appendLog(`testCaseID: ${testCaseRowNum}: Search Query Finished Successfully`);
    
    return { testCaseRowNum: testCaseRowNum, status_code: status, output: json_output };

}

export async function calculateSimilarityUsingPalm2(testCaseNum, sentence1, sentence2, config) {

    const token = config.accessToken;
    const projectId = config.vertexAIProjectID;
    const location = config.vertexAILocation;
    const summaryMatchingAdditionalPrompt = config.summaryMatchingAdditionalPrompt === null ? "" : config.summaryMatchingAdditionalPrompt;

    var prompt = summaryMatching_prompt
        + summaryMatchingAdditionalPrompt
        + summaryMatching_examples;

    var full_prompt = `${prompt} answer_1: ${sentence1} answer_2: ${sentence2} output:`;

    var data = {
        instances: [
            { prompt: `${full_prompt}` }
        ],
        parameters: {
            temperature: 0.2,
            maxOutputTokens: 256,
            topK: 40,
            topP: 0.95,
            logprobs: 2
        }
    };

    const url = `https://${location}-aiplatform.googleapis.com/v1/projects/${projectId}/locations/${location}/publishers/google/models/text-bison:predict`;

    const { status, json_output } = await callVettexAI(url, token, data);
    const output = json_output.predictions[0].content;
    
    appendLog(`testCaseID: ${testCaseNum}: SummaryMatch Finished Successfully `);

    return { testCaseNum: testCaseNum, status_code: status, output: `${output}` };
}

async function callVettexAI(url, token, data) {
    const response = await fetch(url, {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
        },
        body: JSON.stringify(data),
    });

    if (!response.ok) {

        if (response.status === 401) {
            const json = await response.json();
            throw new NotAuthenticatedError(json.error.message);
        }
        else if (response.status === 429) {
            const json = await response.json();
            throw new QuotaError(json.error.message);
        }
        else {

            throw new Error("Error calling VertexAI for Summary, HTTP Status: " + response.status);
        }
    } 
    
    const json = await response.json();

    return { status: response.status, json_output: json };
}

export async function callGeminiMultitModal(prompt,fileUri, mimeType, config) {

    const token = config.accessToken;
    const projectId = config.vertexAIProjectID;
    const location = config.vertexAILocation;
    const model_id = config.function_model_id;
   
    var data = {
        contents: [
            {
                role: "user",
                parts: [
                    {
                        fileData: {
                            mimeType: `${mimeType}`,
                            fileUri: `${fileUri}`
                        }
                    },
                    {
                        text: `${prompt}`
                    },
                ]
            }
        ],
        generationConfig: {
            maxOutputTokens: 8192,
            temperature: 1,
            topP: 0.95,
            response_mime_type: "application/json"
        },
        safetySettings: [
            {
                category: "HARM_CATEGORY_HATE_SPEECH",
                threshold: "BLOCK_MEDIUM_AND_ABOVE"
            },
            {
                category: "HARM_CATEGORY_DANGEROUS_CONTENT",
                threshold: "BLOCK_MEDIUM_AND_ABOVE"
            },
            {
                category: "HARM_CATEGORY_SEXUALLY_EXPLICIT",
                threshold: "BLOCK_MEDIUM_AND_ABOVE"
            },
            {
                category: "HARM_CATEGORY_HARASSMENT",
                threshold: "BLOCK_MEDIUM_AND_ABOVE"
            }
        ],
    };
   
    const url = `https://${location}-aiplatform.googleapis.com/v1/projects/${projectId}/locations/${location}/publishers/google/models/${model_id}:generateContent`;

    const { status, json_output } = await callVettexAI(url, token, data);
    const output = json_output.candidates[0].candidates.parts[0].text;
   
    appendLog(`GeminiFn: Generate Finished Successfully.`);
    
    return {status_code: status, output: `${output}` };
}
