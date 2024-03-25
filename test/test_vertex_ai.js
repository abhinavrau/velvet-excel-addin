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
    it('should return a list of search results', async () => {
        const query = 'query';
        const config = {
            accessToken: 'YOUR_ACCESS_TOKEN',
            preamble: 'preamble',
            summaryResultCount: 2,
            maxExtractiveAnswerCount: 2,
            ignoreAdversarialQuery: false,
            ignoreNonSummarySeekingQuery: false,
            vertexAISearchProjectNumber: 'YOUR_PROJECT_NUMBER',
            vertexAISearchDataStoreName: 'YOUR_DATASTORE_NAME',
            vertexAILocation: 'YOUR_LOCATION',
        };

        // Read json from file into variable 
        const rawData = fs.readFileSync('./test/test_response.json');
        const data = JSON.parse(rawData);

        const response = {
            rowNum: 1,
            response: data,
        };

        const url = `https://discoveryengine.googleapis.com/v1alpha/projects/${config.vertexAISearchProjectNumber}/locations/global/collections/default_collection/dataStores/${config.vertexAISearchDataStoreName}/servingConfigs/default_search:search`;

        fetchMock.postOnce(url, {
            status: 200,
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(data)
        });


        const result = await callVertexAISearch(1, query, config);

        expect(fetchMock.called()).toBe(true);
        expect(fetchMock.lastUrl().toLowerCase()).toBe(url.toLowerCase());
        expect(result).toEqual(response);
    });

});