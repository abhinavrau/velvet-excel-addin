import fetchMock from 'fetch-mock';
import fs from 'fs';





export function mockVertexAISearchRequestResponse(testCaseNum, expected_status_code,
    expectedRequestFile, expectedResponseFile, config) {
    const requestData = fs.readFileSync(expectedRequestFile);
    const requestJson = JSON.parse(requestData);

    // Read response  json from file into variable 
    const responseData = fs.readFileSync(expectedResponseFile);
    const responseJson = JSON.parse(responseData);

    // expected response with row number
    const expectedResponse = {
        id: testCaseNum,
        status_code: expected_status_code,
        output: responseJson
    };

    const url = `https://discoveryengine.googleapis.com/v1alpha/projects/${config.vertexAISearchProjectNumber}/locations/global/collections/default_collection/dataStores/${config.vertexAISearchDataStoreName}/servingConfigs/default_search:search`;

    // mock the call with our response we want to return
    var response =  fetchMock.postOnce(url, {
        status: expected_status_code,
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(responseJson)
    });
    return { requestJson, url, expectedResponse };
}