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

        var data = {
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
            body: JSON.stringify(data)
        });


        const result = await calculateSimilarityUsingVertexAI(sentence1, sentence2, config);

        expect(fetchMock.called()).toBe(true);
        expect(fetchMock.lastUrl().toLowerCase()).toBe(url.toLowerCase());
        expect(result).toEqual('same');
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
            ignoreAdversarialQuery: false,
            ignoreNonSummarySeekingQuery: false,
            vertexAISearchProjectNumber: 'YOUR_PROJECT_NUMBER',
            vertexAISearchDataStoreName: 'YOUR_DATASTORE_NAME',
            vertexAILocation: 'YOUR_LOCATION',
        };

        
        await testRequestResponse(config, query,
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
            model: "preview",
            ignoreAdversarialQuery: true,
            ignoreNonSummarySeekingQuery: true,
            vertexAISearchProjectNumber: 'YOUR_PROJECT_NUMBER',
            vertexAISearchDataStoreName: 'YOUR_DATASTORE_NAME',
            vertexAILocation: 'YOUR_LOCATION',
        };

       
        await testRequestResponse(config, query,
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
            model: "preview",
            ignoreAdversarialQuery: true,
            ignoreNonSummarySeekingQuery: true,
            vertexAISearchProjectNumber: 'YOUR_PROJECT_NUMBER',
            vertexAISearchDataStoreName: 'YOUR_DATASTORE_NAME',
            vertexAILocation: 'YOUR_LOCATION',
        };

     
        await testRequestResponse(config, query,
            './test/data/snippets/test_vai_search_snippet_request.json',
            './test/data/snippets/test_vai_search_snippet_response.json');
    });


});

async function testRequestResponse(config, query, expectedRequestFile, expectedResponseFile) {
    const { requestJson, url, expectedResponse } = prepareVertexAISearchRequestResponse(expectedRequestFile, expectedResponseFile, config);


    const result = await callVertexAISearch(1, query, config);

    expect(fetchMock.called()).toBe(true);
    // Assert request body is correct
    expect(JSON.parse(fetchMock.lastCall()[1].body)).toEqual(requestJson);
    // Assert URL is correct
    expect(fetchMock.lastUrl().toLowerCase()).toBe(url.toLowerCase());
    // Assert response is correct
    expect(result).toEqual(expectedResponse);
}

export function prepareVertexAISearchRequestResponse(expectedRequestFile, expectedResponseFile, config) {
    const requestData = fs.readFileSync(expectedRequestFile);
    const requestJson = JSON.parse(requestData);

    // Read response  json from file into variable 
    const responseData = fs.readFileSync(expectedResponseFile);
    const responseJson = JSON.parse(responseData);

    // expected response with row number
    const expectedResponse = {
        rowNum: 1,
        response: responseJson,
    };

    const url = `https://discoveryengine.googleapis.com/v1alpha/projects/${config.vertexAISearchProjectNumber}/locations/global/collections/default_collection/dataStores/${config.vertexAISearchDataStoreName}/servingConfigs/default_search:search`;

    var response = fetchMock.postOnce(url, {
        status: 200,
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(responseJson)
    });
    return { requestJson, url, expectedResponse };
}

