
import { NotAuthenticatedError, QuotaError, summaryMatching_examples, summaryMatching_prompt } from "./common.js";
import { appendLog } from "./ui.js";

export async function callVertexAISearch(testCaseRowNum, query, config) {
    var status;
    var output;


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

    const response = await fetch(url, {
        method: 'POST',
        headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json',
        },
        body: JSON.stringify(data),
    });
   
    if (!response.ok) {
        
        const log = `callVertexAISearch: Request failed for testCase#: ${testCaseRowNum} error: ${response.status}`;
        console.error(log);
        if (response.status === 401) {
            const json = await response.json();
            throw new NotAuthenticatedError(json.error.message);
        }
        else if (response.status === 429) {
            const json = await response.json();
            throw new QuotaError(json.error.message);
        }
        else {
          
            throw new Error("Error calling VertexAISearch, HTTP Status: " + response.status);
        }
        
    } else {
        output = await response.json();

        appendLog(`testCaseID: ${testCaseRowNum}: Search Query Finished Successfully`);
    }
    status = response.status;

    return { testCaseRowNum: testCaseRowNum, status_code: status, output: output };

}

export async function calculateSimilarityUsingVertexAI(testCaseNum, sentence1, sentence2, config) {

    var status;
    var output;



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
    } else {
        const json = await response.json();
        output = json.predictions[0].content;
    }
    status = response.status;
    appendLog(`testCaseID: ${testCaseNum}: SummaryMatch Finished Successfully `);

    return { testCaseNum: testCaseNum, status_code: status, output: `${output}` };
}
