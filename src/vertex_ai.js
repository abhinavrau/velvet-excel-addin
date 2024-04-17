
export async function callVertexAISearch(rowNum, query, config) {

    try {

        const token = config.accessToken;
        const preamble = config.preamble;
        const model = config.model;
        const summaryResultCount = config.summaryResultCount;
        const extractiveContentSpec = config.extractiveContentSpec === null ? {} : config.extractiveContentSpec;
        const snippetSpec = config.snippetSpec === null ? {} : config.snippetSpec;
        const ignoreAdversarialQuery = config.ignoreAdversarialQuery;
        const ignoreNonSummarySeekingQuery = config.ignoreNonSummarySeekingQuery;
        const projectNumber = config.vertexAISearchProjectNumber;
        const datastoreName = config.vertexAISearchDataStoreName;

       // console.log('config: ' + JSON.stringify(config));

        var data = {
            query: query,
            page_size: "5",
            offset: 0,
            contentSearchSpec: {
                extractiveContentSpec,
                snippetSpec,
                summarySpec: {
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
            throw new Error(`callVertexAISearch: Request failed with with code:${response.code} status:${response.status} and message:${response.message}`);
        }
        const json = await response.json();
        return { rowNum: rowNum, response: json };

    } catch (error) {
        console.log('Error calling callVertexAISearch: ' + error);
        throw error;

    }
}

export async function calculateSimilarityUsingVertexAI(sentence1, sentence2, config) {

    try {

        const token = config.accessToken;
        const projectId = config.vertexAIProjectID;
        const location = config.vertexAILocation;

        var prompt = "You will get two answers to a question, you should determine if they are semantically similar or not. If any monetory numbers in the answers, they should be matched exactly." +
            "examples - answer_1: I was created by X. answer_2: X created me. output:same " +
            "answer_1:There are 52 days in a year. answer_2: A year is fairly long. output:different ";
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
            throw new Error(`calculateSimilarityUsingVertexAI: Request failed with code:${response.code} status:${response.status} and message:${response.message}`);
        }
        const json = await response.json();
        return json.predictions[0].content;

    } catch (error) {
        console.log('Error calling calculateSimilarityUsingVertexAI: ' + error);
        throw error;

    }
}
