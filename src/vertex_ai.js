
export async function callVertexAISearch(testCaseNum, query, config) {
    var status;
    var output;

    try {
        const token = config.accessToken;
        const preamble = config.preamble;
        const model = config.model === "" ? 'gemini-1.0-pro-001/answer_gen/v1' : config.model;
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
                    useSemanticChunks: `${useSemanticChunks}`,
                    summaryResultCount: `${summaryResultCount}`,
                    ignoreAdversarialQuery: `${ignoreAdversarialQuery}`,
                    ignoreNonSummarySeekingQuery: `${ignoreNonSummarySeekingQuery}`,
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
            const json = await response.json();
            output = `callVertexAISearch: Request failed for testCase#: ${testCaseNum} error: ${json.error.message}`;
            console.error(output);
        } else {
            output = await response.json();
            console.log(`callVertexAISearch: Finished Successfully row: ${testCaseNum}`);
        }
        status = response.status;
    
    } catch (error) {
        output = `callVertexAISearch: Caught Exception for testCase#: ${testCaseNum} error: ${error} with stack: ${error.stack}`;
        status = 0;
        console.error(output);
    }
    return { testCaseNum: testCaseNum, status_code: status, output: output};

}

export async function calculateSimilarityUsingVertexAI(testCaseNum, sentence1, sentence2, config) {

    var status;
    var output;

    try {
       
        const token = config.accessToken;
        const projectId = config.vertexAIProjectID;
        const location = config.vertexAILocation;
        const summaryMatchingAdditionalPrompt = config.summaryMatchingAdditionalPrompt === null ? "" : config.summaryMatchingAdditionalPrompt;

        var prompt = "You will get two answers to a question, you should determine if they are semantically similar or not. "
            + summaryMatchingAdditionalPrompt +
            " examples - answer_1: I was created by X. answer_2: X created me. output:same "
            + "answer_1:There are 52 days in a year. answer_2: A year is fairly long. output:different ";
        /* "answer_1:The revenue was $10 milllion in 2022. answer_2: In 2022 the revenue was $10 million output:same " +
        "answer_1:The revenue was $12 milllion in 2022. answer_2: In 2022 the revenue was $10 million output:different " +
        "answer_1:The revenue was $12 milllion in 2022. answer_2: In 2022 the revenue was $1200  output:different " +
        "answer_1:Alphabet total net income was $59.9 billion in 2022. answer_2: Alphabet's net income in 2022 was $59972. output:different "; */
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
            const json = await response.json();
            output = `calculateSimilarityUsingVertexAI: Request failed with for testCase#: ${testCaseNum} error: ${json.error.message}`;
            console.error(output);
        } else {
            const json = await response.json();
            output =json.predictions[0].content;
        }
        status = response.status;

    } catch (error) {
        output = `calculateSimilarityUsingVertexAI: Caught Exception for testCase#: ${testCaseNum} error: ${error} with stack: ${error.stack}`;
        status = 0;
        console.error(output);
    }
        
    return { testCaseNum: testCaseNum, status_code: status, output: `${output}` };
}
