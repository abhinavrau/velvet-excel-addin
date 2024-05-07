import expect from 'expect';
import fetchMock from 'fetch-mock';
import fs from 'fs';

import { calculateSimilarityUsingVertexAI, callVertexAISearch } from '../src/vertex_ai.js';

describe('calculateSimilarityUsingVertexAI', () => {
    beforeEach(() => {
        fetchMock.reset();
    });
    it('should return a similarity  between two sentences', async () => {
        const sentence1 = 'sentece 1';
        const sentence2 = 'sentence 2';

        const config = {
            accessToken: 'YOUR_ACCESS_TOKEN',
            vertexAIProjectID: 'YOUR_PROJECT_ID',
            vertexAILocation: 'YOUR_LOCATION',
        };

        var response = {
            predictions: [
                {
                    content: 'same',
                },
            ],
        };
        const url = `https://${config.vertexAILocation}-aiplatform.googleapis.com/v1/projects/${config.vertexAIProjectID}/locations/${config.vertexAILocation}/publishers/google/models/text-bison:predict`;
        fetchMock.postOnce(url, {
            status: 200,
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(response)
        });


        const result = await calculateSimilarityUsingVertexAI(1, sentence1, sentence2, config);
        const expectedResponse = {
            testCaseNum: 1,
            status_code: 200,
            output: 'same',
        };
        expect(fetchMock.called()).toBe(true);
        expect(fetchMock.lastUrl().toLowerCase()).toBe(url.toLowerCase());
        expect(result).toEqual(expectedResponse);
    });
    it('should fail when you get an error from Vertex AI', async () => {
        const sentence1 = 'sentece 1';
        const sentence2 = 'sentence 2';

        const config = {
            accessToken: 'YOUR_ACCESS_TOKEN',
            vertexAIProjectID: 'YOUR_PROJECT_ID',
            vertexAILocation: 'YOUR_LOCATION',
        };

        var response = {
            error:
            {
                message: 'Call failed with status code 500 and status message: Internal Server Error',
            },
        };
        const url = `https://${config.vertexAILocation}-aiplatform.googleapis.com/v1/projects/${config.vertexAIProjectID}/locations/${config.vertexAILocation}/publishers/google/models/text-bison:predict`;
        fetchMock.postOnce(url, {
            status: 500,
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(response)
        });


        const result = await calculateSimilarityUsingVertexAI(1, sentence1, sentence2, config);
        const expectedResponse = {
            testCaseNum: 1,
            status_code: 500,
            output: "calculateSimilarityUsingVertexAI: Request failed with for testCase#: 1 error: " + response.error.message,
        };
        expect(fetchMock.called()).toBe(true);
        expect(fetchMock.lastUrl().toLowerCase()).toBe(url.toLowerCase());
        expect(result).toEqual(expectedResponse);
    });

    it('should fail when fetch throws exception', async () => {
        const sentence1 = 'sentece 1';
        const sentence2 = 'sentence 2';

        const config = {
            accessToken: 'YOUR_ACCESS_TOKEN',
            vertexAIProjectID: 'YOUR_PROJECT_ID',
            vertexAILocation: 'YOUR_LOCATION',
        };

        var response = {
            error:
            {
                message: 'Call failed with status code 500 and status message: Internal Server Error',
            },
        };
        const url = `https://${config.vertexAILocation}-aiplatform.googleapis.com/v1/projects/${config.vertexAIProjectID}/locations/${config.vertexAILocation}/publishers/google/models/text-bison:predict`;
        fetchMock.postOnce(url, {
            throws: new Error('Mocked error')
        });


        const result = await calculateSimilarityUsingVertexAI(1, sentence1, sentence2, config);
        const expectedResponse = {
            testCaseNum: 1,
            status_code: 0,
            output: "calculateSimilarityUsingVertexAI: Caught Exception for testCase#: 1 error: Error: Mocked error with stack: Error: Mocked error",
        };
        expect(fetchMock.called()).toBe(true);
        expect(fetchMock.lastUrl().toLowerCase()).toBe(url.toLowerCase());
        expect(result.status_code).toEqual(expectedResponse.status_code);
        // check is result.output is a subset of expectedResponse.output
        expect(result.output.includes(expectedResponse.output)).toBe(true);
    });

});

 describe('callVertexAISearch', () => {
    beforeEach(() => {
        fetchMock.reset();
    });
    it('should return a list of search results for Extractive Answer', async () => {
        const query = 'query';
        const config = {
            accessToken: 'YOUR_ACCESS_TOKEN',
            preamble: 'preamble',
            extractiveContentSpec: {
                maxExtractiveAnswerCount: 2,
            },
            summaryResultCount: 2,
            model: "preview",
            useSemanticChunks: true,
            ignoreAdversarialQuery: false,
            ignoreNonSummarySeekingQuery: false,
            vertexAISearchProjectNumber: 'YOUR_PROJECT_NUMBER',
            vertexAISearchDataStoreName: 'YOUR_DATASTORE_NAME',
            vertexAILocation: 'YOUR_LOCATION',
        };


        await testRequestResponse(1, 200,config, query,
            './test/data/extractive_answer/test_vai_search_extractive_answer_request.json',
            './test/data/extractive_answer/test_vai_search_extractive_answer_response.json');


    });
     it('should return a list of search results for Extractive Segments', async () => {
        const query = 'query';
        const config = {
            accessToken: 'YOUR_ACCESS_TOKEN',
            preamble: 'You are an expert financial analyst. Be brief.',
            extractiveContentSpec: {
                maxExtractiveSegmentCount: 2,
            },
            summaryResultCount: 2,
            useSemanticChunks: true,
            model: "preview",
            ignoreAdversarialQuery: true,
            ignoreNonSummarySeekingQuery: true,
            vertexAISearchProjectNumber: 'YOUR_PROJECT_NUMBER',
            vertexAISearchDataStoreName: 'YOUR_DATASTORE_NAME',
            vertexAILocation: 'YOUR_LOCATION',
        };


        await testRequestResponse(1, 200, config, query,
            './test/data/extractive_segment/test_vai_search_extractive_segment_request.json',
            './test/data/extractive_segment/test_vai_search_extractive_segment_response.json');
    });
    it('should return a list of search results for Snippets', async () => {
        const query = 'query';
        const config = {
            accessToken: 'YOUR_ACCESS_TOKEN',
            preamble: 'You are an expert financial analyst. Be brief.',
            snippetSpec: {
                maxSnippetCount: 2,
            },
            summaryResultCount: 2,
            useSemanticChunks: true,
            model: "preview",
            ignoreAdversarialQuery: true,
            ignoreNonSummarySeekingQuery: true,
            vertexAISearchProjectNumber: 'YOUR_PROJECT_NUMBER',
            vertexAISearchDataStoreName: 'YOUR_DATASTORE_NAME',
            vertexAILocation: 'YOUR_LOCATION',
        };


        await testRequestResponse(1, 200, config, query,
            './test/data/snippets/test_vai_search_snippet_request.json',
            './test/data/snippets/test_vai_search_snippet_response.json');
    });
    it('should return aan error when Vertex AI Search fails', async () => {
        const query = 'query';
        const config = {
            accessToken: 'YOUR_ACCESS_TOKEN',
            preamble: 'preamble',
            extractiveContentSpec: {
                maxExtractiveAnswerCount: 2,
            },
            summaryResultCount: 2,
            model: "preview",
            useSemanticChunks: true,
            ignoreAdversarialQuery: false,
            ignoreNonSummarySeekingQuery: false,
            vertexAISearchProjectNumber: 'YOUR_PROJECT_NUMBER',
            vertexAISearchDataStoreName: 'YOUR_DATASTORE_NAME',
            vertexAILocation: 'YOUR_LOCATION',
        };


        await testRequestResponse(1, 401, config, query,
            './test/data/extractive_answer/test_vai_search_extractive_answer_request.json',
            './test/data/not_authenticated.json');


    }); 
     it('should fail when fetch throws exception', async () => {
         const query = 'query';
         const config = {
             accessToken: 'YOUR_ACCESS_TOKEN',
             preamble: 'preamble',
             extractiveContentSpec: {
                 maxExtractiveAnswerCount: 2,
             },
             summaryResultCount: 2,
             model: "preview",
             useSemanticChunks: true,
             ignoreAdversarialQuery: false,
             ignoreNonSummarySeekingQuery: false,
             vertexAISearchProjectNumber: 'YOUR_PROJECT_NUMBER',
             vertexAISearchDataStoreName: 'YOUR_DATASTORE_NAME',
             vertexAILocation: 'YOUR_LOCATION',
         };


         const url = `https://discoveryengine.googleapis.com/v1alpha/projects/${config.vertexAISearchProjectNumber}/locations/global/collections/default_collection/dataStores/${config.vertexAISearchDataStoreName}/servingConfigs/default_search:search`;

         var response = fetchMock.postOnce(url, {
             throws: new Error('Mocked error')
         });
         const result = await callVertexAISearch(1, query, config);

         const expectedResponse = {
             testCaseNum: 1,
             status_code: 0,
             output: "callVertexAISearch: Caught Exception for testCase#: 1 error: Error: Mocked error with stack: Error: Mocked error",
         };
         
         expect(fetchMock.called()).toBe(true);
         expect(fetchMock.lastUrl().toLowerCase()).toBe(url.toLowerCase());
         expect(result.status_code).toEqual(expectedResponse.status_code);
         // check is result.output is a subset of expectedResponse.output
         expect(result.output.includes(expectedResponse.output)).toBe(true);
     }); 

});

async function testRequestResponse(testCaseNum, expected_status_code, config, query, expectedRequestFile, expectedResponseFile) {
    const { requestJson, url, expectedResponse } = prepareVertexAISearchRequestResponse(testCaseNum, expected_status_code, expectedRequestFile, expectedResponseFile, config);


    const result = await callVertexAISearch(1, query, config);

    expect(fetchMock.called()).toBe(true);
    // Assert request body is correct
    expect(JSON.parse(fetchMock.lastCall()[1].body)).toEqual(requestJson);
    // Assert URL is correct
    expect(fetchMock.lastUrl().toLowerCase()).toBe(url.toLowerCase());
    // Assert response is correct
    if (expected_status_code == 200) {
        expect(result).toEqual(expectedResponse);
    } else if(expected_status_code >= 400) {
        expect(JSON.stringify(result.output).includes("callVertexAISearch: Request failed for testCase")).toBe(true);
    } else {
        expect(JSON.stringify(result.output).includes("callVertexAISearch: Caught Exception for testCase")).toBe(true);
    }
    

}

export function prepareVertexAISearchRequestResponse(testCaseNum, expected_status_code, expectedRequestFile, expectedResponseFile, config) {
    const requestData = fs.readFileSync(expectedRequestFile);
    const requestJson = JSON.parse(requestData);

    // Read response  json from file into variable 
    const responseData = fs.readFileSync(expectedResponseFile);
    const responseJson = JSON.parse(responseData);

    // expected response with row number
    const expectedResponse = {
        testCaseNum: testCaseNum,
        status_code: expected_status_code,
        output: responseJson
    };

    const url = `https://discoveryengine.googleapis.com/v1alpha/projects/${config.vertexAISearchProjectNumber}/locations/global/collections/default_collection/dataStores/${config.vertexAISearchDataStoreName}/servingConfigs/default_search:search`;

    var response = fetchMock.postOnce(url, {
        status: expected_status_code,
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(responseJson)
    });
    return { requestJson, url, expectedResponse };
}

 