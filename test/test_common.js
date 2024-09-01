import fetchMock from "fetch-mock";
import fs from "fs";

export function mockVertexAISearchRequestResponse(
  testCaseNum,
  expected_status_code,
  expectedRequestFile,
  expectedResponseFile,
  config,
) {
  const requestData = fs.readFileSync(expectedRequestFile);
  const requestJson = JSON.parse(requestData);

  // Read response  json from file into variable
  const responseData = fs.readFileSync(expectedResponseFile);
  const responseJson = JSON.parse(responseData);

  // expected response with row number
  const expectedResponse = {
    id: testCaseNum,
    status_code: expected_status_code,
    output: responseJson,
  };

  const url = `https://discoveryengine.googleapis.com/v1alpha/projects/${config.vertexAISearchProjectNumber}/locations/global/collections/default_collection/dataStores/${config.vertexAISearchDataStoreName}/servingConfigs/default_search:search`;

  // mock the call with our response we want to return
  var response = fetchMock.post(url, {
    status: expected_status_code,
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify(responseJson),
  });
  return { requestJson, url, expectedResponse };
}

export function mockGeminiRequestResponse(
  testCaseNum,
  expected_status_code,
  expectedRequestFile,
  expectedResponseFile,
  model_id,
  config,
) {
  const requestData = fs.readFileSync(expectedRequestFile);
  const requestJson = JSON.parse(requestData);

  // Read response  json from file into variable
  const responseData = fs.readFileSync(expectedResponseFile);
  const responseJson = JSON.parse(responseData);

  // expected response with row number
  const expectedResponse = {
    id: testCaseNum,
    status_code: expected_status_code,
    output: responseJson,
  };

  const url = `https://${config.vertexAILocation}-aiplatform.googleapis.com/v1/projects/${config.vertexAIProjectID}/locations/${config.vertexAILocation}/publishers/google/models/${model_id}:generateContent`;

  // mock the call with our response we want to return
  var response = fetchMock.post(url, {
    status: expected_status_code,
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify(responseJson),
  });
  return { requestJson, url, expectedResponse };
}

export function getRequestResponseJsonFromFile(requestJsonFilePath, responseJsonFilePath) {
  const request = fs.readFileSync(requestJsonFilePath);
  const response = fs.readFileSync(responseJsonFilePath);
  return {
    response_json: JSON.parse(response),
    request_json: JSON.parse(request),
  };
}
