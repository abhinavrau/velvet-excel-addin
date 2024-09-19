import expect from "expect";
import fetchMock from "fetch-mock";
import { default as $, default as JQuery } from "jquery";
import sinon from "sinon";
import {
  NotAuthenticatedError,
  QuotaError,
  ResourceNotFoundError,
  VertexAIError,
  summaryMatching_examples,
  summaryMatching_prompt,
} from "../src/common.js";

import {
  calculateSimilarityUsingPalm2,
  callCheckGrounding,
  callVertexAISearch,
} from "../src/vertex_ai.js";
import { mockDiscoveryEngineRequestResponse } from "./test_common.js";

global.$ = $;
global.JQuery = JQuery;

describe("When calculateSimilarityUsingVertexAI is called ", () => {
  var $stub;

  beforeEach(() => {
    // stub out jQuery calls
    $stub = sinon.stub(globalThis, "$").returns({
      empty: sinon.stub(),
      append: sinon.stub(),
      val: sinon.stub(),
      tabulator: sinon.stub(),
    });

    fetchMock.reset();
  });

  afterEach(() => {
    $stub.restore();
    sinon.reset();
  });

  it("should return a similarity  between two sentences", async () => {
    const sentence1 = "sentece 1";
    const sentence2 = "sentence 2";

    const config = {
      accessToken: "YOUR_ACCESS_TOKEN",
      vertexAIProjectID: "YOUR_PROJECT_ID",
      vertexAILocation: "YOUR_LOCATION",
      summaryMatchingAdditionalPrompt: "addional prompt",
    };

    var response = {
      predictions: [
        {
          content: "same",
        },
      ],
    };
    const url = `https://${config.vertexAILocation}-aiplatform.googleapis.com/v1/projects/${config.vertexAIProjectID}/locations/${config.vertexAILocation}/publishers/google/models/text-bison:predict`;
    fetchMock.postOnce(url, {
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(response),
    });

    const result = await calculateSimilarityUsingPalm2(1, sentence1, sentence2, config);
    const expectedResponse = {
      id: 1,
      status_code: 200,
      output: "same",
    };
    expect(fetchMock.called()).toBe(true);
    expect(fetchMock.lastUrl().toLowerCase()).toBe(url.toLowerCase());
    expect(fetchMock.lastCall()[1].headers).toEqual({
      Authorization: `Bearer ${config.accessToken}`,
      "Content-Type": "application/json",
    });

    // validate header and body
    var prompt =
      summaryMatching_prompt + config.summaryMatchingAdditionalPrompt + summaryMatching_examples;

    var full_prompt = `${prompt} answer_1: ${sentence1} answer_2: ${sentence2} output:`;

    expect(JSON.parse(fetchMock.lastCall()[1].body)).toEqual({
      instances: [{ prompt: `${full_prompt}` }],
      parameters: {
        temperature: 0.2,
        maxOutputTokens: 256,
        topK: 40,
        topP: 0.95,
        logprobs: 2,
      },
    });

    expect(result).toEqual(expectedResponse);
  });
  it("should fail when you get an authentication error from Vertex AI", async () => {
    const sentence1 = "sentece 1";
    const sentence2 = "sentence 2";

    const config = {
      accessToken: "YOUR_ACCESS_TOKEN",
      vertexAIProjectID: "YOUR_PROJECT_ID",
      vertexAILocation: "YOUR_LOCATION",
    };

    var response = {
      error: {
        message: "Call failed with status code 500 and status message: Internal Server Error",
      },
    };
    const url = `https://${config.vertexAILocation}-aiplatform.googleapis.com/v1/projects/${config.vertexAIProjectID}/locations/${config.vertexAILocation}/publishers/google/models/text-bison:predict`;
    fetchMock.postOnce(url, {
      status: 401,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(response),
    });

    try {
      const result = await calculateSimilarityUsingPalm2(1, sentence1, sentence2, config);
      assert.fail();
    } catch (err) {
      expect(fetchMock.called()).toBe(true);
      expect(fetchMock.lastUrl().toLowerCase()).toBe(url.toLowerCase());
      expect(err).toBeInstanceOf(NotAuthenticatedError);
    }
  });

  it("should fail when fetch throws exception", async () => {
    const sentence1 = "sentece 1";
    const sentence2 = "sentence 2";

    const config = {
      accessToken: "YOUR_ACCESS_TOKEN",
      vertexAIProjectID: "YOUR_PROJECT_ID",
      vertexAILocation: "YOUR_LOCATION",
    };

    const url = `https://${config.vertexAILocation}-aiplatform.googleapis.com/v1/projects/${config.vertexAIProjectID}/locations/${config.vertexAILocation}/publishers/google/models/text-bison:predict`;
    fetchMock.postOnce(url, {
      throws: new Error("Mocked error"),
    });

    try {
      const result = await calculateSimilarityUsingPalm2(1, sentence1, sentence2, config);
      assert.fail();
    } catch (err) {
      expect(fetchMock.called()).toBe(true);
      expect(fetchMock.lastUrl().toLowerCase()).toBe(url.toLowerCase());
      expect(err.message).toBe("Network or unexpected error: Mocked error");
    }
  });
});

describe("When callVertexAISearch is called", () => {
  var $stub;
  beforeEach(() => {
    // stub out jQuery calls
    $stub = sinon.stub(globalThis, "$").returns({
      empty: sinon.stub(),
      append: sinon.stub(),
      val: sinon.stub(),
      tabulator: sinon.stub(),
    });

    fetchMock.reset();
  });

  afterEach(() => {
    $stub.restore();
  });
  it("should return a list of search results for Extractive Answer", async () => {
    const query = "What is Google's revenue for the year ending December 31, 2021";
    const config = {
      accessToken: "YOUR_ACCESS_TOKEN",
      preamble:
        "You are an expert financial analyst.  Only use the data returned from documents. All finance numbers must be reported in billions, millions or thousands. Be brief. Answer should be no more than 2 sentences please.",
      extractiveContentSpec: {
        maxExtractiveAnswerCount: "2",
        maxExtractiveSegmentCount: "0",
      },
      summaryResultCount: 2,
      model: "gemini-1.0-pro-001/answer_gen/v1",
      useSemanticChunks: false,
      ignoreAdversarialQuery: true,
      ignoreNonSummarySeekingQuery: true,
      vertexAISearchProjectNumber: "YOUR_PROJECT_NUMBER",
      vertexAISearchAppId: "YOUR_DATASTORE_NAME",
      vertexAILocation: "YOUR_LOCATION",
    };

    const url = `https://discoveryengine.googleapis.com/v1alpha/projects/${config.vertexAISearchProjectNumber}/locations/global/collections/default_collection/engines/${config.vertexAISearchAppId}/servingConfigs/default_search:search`;
    await testRequestResponse(
      1,
      url,
      200,
      config,
      query,
      "./test/data/search/search_extractive_answer/test_vai_search_extractive_answer_request.json",
      "./test/data/search/search_extractive_answer/test_vai_search_extractive_answer_response.json",
    );
  });
  it("should return a list of search results for Extractive Segments", async () => {
    const query = "query";
    const config = {
      accessToken: "YOUR_ACCESS_TOKEN",
      preamble: "You are an expert financial analyst. Be brief.",
      extractiveContentSpec: {
        maxExtractiveSegmentCount: 2,
      },
      summaryResultCount: 2,
      useSemanticChunks: true,
      model: "preview",
      ignoreAdversarialQuery: true,
      ignoreNonSummarySeekingQuery: true,
      vertexAISearchProjectNumber: "YOUR_PROJECT_NUMBER",
      vertexAISearchAppId: "YOUR_DATASTORE_NAME",
      vertexAILocation: "YOUR_LOCATION",
    };
    const url = `https://discoveryengine.googleapis.com/v1alpha/projects/${config.vertexAISearchProjectNumber}/locations/global/collections/default_collection/engines/${config.vertexAISearchAppId}/servingConfigs/default_search:search`;
    await testRequestResponse(
      1,
      url,
      200,
      config,
      query,
      "./test/data/search/search_extractive_segment/test_vai_search_extractive_segment_request.json",
      "./test/data/search/search_extractive_segment/test_vai_search_extractive_segment_response.json",
    );
  });
  it("should return a list of search results for Snippets", async () => {
    const query = "query";
    const config = {
      accessToken: "YOUR_ACCESS_TOKEN",
      preamble: "You are an expert financial analyst. Be brief.",
      snippetSpec: {
        maxSnippetCount: 2,
      },
      summaryResultCount: 2,
      useSemanticChunks: true,
      model: "preview",
      ignoreAdversarialQuery: true,
      ignoreNonSummarySeekingQuery: true,
      vertexAISearchProjectNumber: "YOUR_PROJECT_NUMBER",
      vertexAISearchAppId: "YOUR_DATASTORE_NAME",
      vertexAILocation: "YOUR_LOCATION",
    };
    const url = `https://discoveryengine.googleapis.com/v1alpha/projects/${config.vertexAISearchProjectNumber}/locations/global/collections/default_collection/engines/${config.vertexAISearchAppId}/servingConfigs/default_search:search`;
    await testRequestResponse(
      1,
      url,
      200,
      config,
      query,
      "./test/data/search/search_snippets/test_vai_search_snippet_request.json",
      "./test/data/search/search_snippets/test_vai_search_snippet_response.json",
    );
  });

  it("should return support score for CheckGrouding API", async () => {
    const answerCandidate =
      "Google's total revenue for the year ending December 31, 2021 was $257,637 million. This represents a 41% increase from the previous year. The majority of Google's revenue comes from its advertising business, which includes Google Search, YouTube ads, and Google Network. In 2021, Google's advertising revenue was $209,497 million. Google's other revenue streams include Google Cloud, which generated $19,206 million in revenue in 2021, and Other Bets, which generated $753 million.";
    const config = {
      accessToken: "YOUR_ACCESS_TOKEN",
      vertexAIProjectID: "YOUR_PROJECT_ID",
    };
    const url = `https://discoveryengine.googleapis.com/v1/projects/${config.vertexAIProjectID}/locations/global/groundingConfigs/default_grounding_config:check`;
    const { requestJson, expectedResponse } = mockDiscoveryEngineRequestResponse(
      1,
      url,
      200,
      "./test/data/search/eval_check_grounding/2_test_check_grouding_request.json",
      "./test/data/search/eval_check_grounding/2_test_check_grouding_response.json",
      config,
    );

    const factsJson = [
      {
        factText:
          "Deferred revenues primarily relate to Google Cloud and Google other. Total deferred revenue as of December 31, 2020 was $3.0 billion, of which \u003cb\u003e$2.3 billion\u003c/b\u003e was recognized as revenues for the year ending December 31, 2021.",
      },
      {
        factText:
          "Year Ended December 31, 2020 2021 Google Services $ $ 91855 Google Cloud ) (3099) Other Bets ) (5281) Corporate costs, unallocated(1) (3299) ) Total income from operations $ 41224 $ (1) Unallocated corporate costs primarily include corporate initiatives, corporate shared costs, such as finance and legal, including certain fines and settlements, as well as costs associated with certain shared R&amp;D activities.",
      },
      {
        factText:
          "Deferred revenues primarily relate to Google Cloud and Google other. Total deferred revenue as of December 31, 2021 was \u003cb\u003e$3.8 billion\u003c/b\u003e, of which $2.5 billion was recognized as revenues for the year ending December 31, 2022.",
      },
      {
        factText:
          "Financial Results Revenues The following table presents revenues by type (in millions): Year Ended December 31, 2021 2022 Google Search &amp; other $ 148951 $ 162450 YouTube ads 28845 29243 Google Network 31701 32780 Google advertising 209497 224473 Google other 28032 29055 Google Services total 237529 253528 Google Cloud 19206 26280 Other Bets 753 1068 Hedging gains (losses) 149 1960 Total revenues $ 257637 $ 282836.",
      },
    ];

    const { id, status_code, output } = await callCheckGrounding(
      config,
      answerCandidate,
      factsJson,
      1,
    );

    expect(fetchMock.called()).toBe(true);
    // Assert request body is correct
    expect(JSON.parse(fetchMock.lastCall()[1].body)).toEqual(requestJson);
    // Assert URL is correct
    expect(fetchMock.lastUrl().toLowerCase()).toBe(url.toLowerCase());
    // Assert response is correct
    expect(output).toEqual(expectedResponse.output);

    expect(id).toEqual(1);
    expect(status_code).toEqual(200);
  });

  it("should return error when Vertex AI Search is not Authenticated", async () => {
    const query = "What is Google's revenue for the year ending December 31, 2021";
    const config = {
      accessToken: "YOUR_ACCESS_TOKEN",
      preamble: "",
      extractiveContentSpec: {
        maxExtractiveAnswerCount: "2",
        maxExtractiveSegmentCount: "0",
      },
      summaryResultCount: 2,
      model: "gemini-1.0-pro-001/answer_gen/v1",
      useSemanticChunks: false,
      ignoreAdversarialQuery: true,
      ignoreNonSummarySeekingQuery: true,
      vertexAISearchProjectNumber: "YOUR_PROJECT_NUMBER",
      vertexAISearchAppId: "YOUR_DATASTORE_NAME",
      vertexAILocation: "YOUR_LOCATION",
    };
    const url = `https://discoveryengine.googleapis.com/v1alpha/projects/${config.vertexAISearchProjectNumber}/locations/global/collections/default_collection/engines/${config.vertexAISearchAppId}/servingConfigs/default_search:search`;
    const { requestJson, expectedResponse } = mockDiscoveryEngineRequestResponse(
      1,
      url,
      401,
      "./test/data/search/search_extractive_answer/test_vai_search_extractive_answer_request.json",
      "./test/data/not_authenticated.json",
      config,
    );

    try {
      const result = await callVertexAISearch(1, query, config);
      assert.fail();
    } catch (err) {
      expect(err instanceof NotAuthenticatedError).toBe(true);
      expect(err.message).toEqual(
        "Request is missing required authentication credential. Expected OAuth 2 access token, login cookie or other valid authentication credential. See https://developers.google.com/identity/sign-in/web/devconsole-project.",
      );
    }
  });
  it("should return error when Vertex AI Search return Quota error", async () => {
    const query = "What is Google's revenue for the year ending December 31, 2021";
    const config = {
      accessToken: "YOUR_ACCESS_TOKEN",
      preamble: "",
      extractiveContentSpec: {
        maxExtractiveAnswerCount: "2",
        maxExtractiveSegmentCount: "0",
      },
      summaryResultCount: 2,
      model: "gemini-1.0-pro-001/answer_gen/v1",
      useSemanticChunks: false,
      ignoreAdversarialQuery: true,
      ignoreNonSummarySeekingQuery: true,
      vertexAISearchProjectNumber: "YOUR_PROJECT_NUMBER",
      vertexAISearchAppId: "YOUR_DATASTORE_NAME",
      vertexAILocation: "YOUR_LOCATION",
    };
    const url = `https://discoveryengine.googleapis.com/v1alpha/projects/${config.vertexAISearchProjectNumber}/locations/global/collections/default_collection/engines/${config.vertexAISearchAppId}/servingConfigs/default_search:search`;
    const { requestJson, expectedResponse } = mockDiscoveryEngineRequestResponse(
      1,
      url,
      429,
      "./test/data/search/search_extractive_answer/test_vai_search_extractive_answer_request.json",
      "./test/data/not_authenticated.json",
      config,
    );

    try {
      const result = await callVertexAISearch(1, query, config);
      assert.fail();
    } catch (err) {
      expect(err instanceof QuotaError).toBe(true);
    }
  });
  it("should fail when fetch throws exception", async () => {
    const query = "query";
    const config = {
      accessToken: "YOUR_ACCESS_TOKEN",
      preamble: "preamble",
      extractiveContentSpec: {
        maxExtractiveAnswerCount: 2,
      },
      summaryResultCount: 2,
      model: "preview",
      useSemanticChunks: true,
      ignoreAdversarialQuery: false,
      ignoreNonSummarySeekingQuery: false,
      vertexAISearchProjectNumber: "YOUR_PROJECT_NUMBER",
      vertexAISearchAppId: "YOUR_DATASTORE_NAME",
      vertexAILocation: "YOUR_LOCATION",
    };

    const url = `https://discoveryengine.googleapis.com/v1alpha/projects/${config.vertexAISearchProjectNumber}/locations/global/collections/default_collection/engines/${config.vertexAISearchAppId}/servingConfigs/default_search:search`;

    var response = fetchMock.postOnce(url, {
      throws: new Error("Mocked error"),
    });

    try {
      const result = await callVertexAISearch(1, query, config);
      assert.fail();
    } catch (err) {
      expect(fetchMock.called()).toBe(true);
      expect(fetchMock.lastUrl().toLowerCase()).toBe(url.toLowerCase());
      expect(err.message).toBe("Network or unexpected error: Mocked error");
    }
  });

  it("should return error when Vertex AI Search returns 404 error", async () => {
    const query = "What is Google's revenue for the year ending December 31, 2021";
    const config = {
      accessToken: "YOUR_ACCESS_TOKEN",
      preamble: "",
      extractiveContentSpec: {
        maxExtractiveAnswerCount: "2",
        maxExtractiveSegmentCount: "0",
      },
      summaryResultCount: 2,
      model: "gemini-1.0-pro-001/answer_gen/v1",
      useSemanticChunks: false,
      ignoreAdversarialQuery: true,
      ignoreNonSummarySeekingQuery: true,
      vertexAISearchProjectNumber: "YOUR_PROJECT_NUMBER",
      vertexAISearchAppId: "YOUR_DATASTORE_NAME",
      vertexAILocation: "YOUR_LOCATION",
    };
    const url = `https://discoveryengine.googleapis.com/v1alpha/projects/${config.vertexAISearchProjectNumber}/locations/global/collections/default_collection/engines/${config.vertexAISearchAppId}/servingConfigs/default_search:search`;
    const { requestJson, expectedResponse } = mockDiscoveryEngineRequestResponse(
      1,
      url,
      404,
      "./test/data/search/search_extractive_answer/test_vai_search_extractive_answer_request.json",
      "./test/data/not_authenticated.json",
      config,
    );

    try {
      const result = await callVertexAISearch(1, query, config);
      assert.fail();
    } catch (err) {
      expect(err instanceof ResourceNotFoundError).toBe(true);
    }
  });
  it("should return error when Vertex AI Search returns any other error", async () => {
    const query = "What is Google's revenue for the year ending December 31, 2021";
    const config = {
      accessToken: "YOUR_ACCESS_TOKEN",
      preamble: "",
      extractiveContentSpec: {
        maxExtractiveAnswerCount: "2",
        maxExtractiveSegmentCount: "0",
      },
      summaryResultCount: 2,
      model: "gemini-1.0-pro-001/answer_gen/v1",
      useSemanticChunks: false,
      ignoreAdversarialQuery: true,
      ignoreNonSummarySeekingQuery: true,
      vertexAISearchProjectNumber: "YOUR_PROJECT_NUMBER",
      vertexAISearchAppId: "YOUR_DATASTORE_NAME",
      vertexAILocation: "YOUR_LOCATION",
    };
    const url = `https://discoveryengine.googleapis.com/v1alpha/projects/${config.vertexAISearchProjectNumber}/locations/global/collections/default_collection/engines/${config.vertexAISearchAppId}/servingConfigs/default_search:search`;
    const { requestJson, expectedResponse } = mockDiscoveryEngineRequestResponse(
      1,
      url,
      405,
      "./test/data/search/search_extractive_answer/test_vai_search_extractive_answer_request.json",
      "./test/data/not_authenticated.json",
      config,
    );

    try {
      const result = await callVertexAISearch(1, query, config);
      assert.fail();
    } catch (err) {
      expect(err instanceof VertexAIError).toBe(true);
    }
  });

  async function testRequestResponse(
    testCaseNum,
    url,
    expected_status_code,
    config,
    query,
    expectedRequestFile,
    expectedResponseFile,
  ) {
    const { requestJson, expectedResponse } = mockDiscoveryEngineRequestResponse(
      testCaseNum,
      url,
      expected_status_code,
      expectedRequestFile,
      expectedResponseFile,
      config,
    );

    const result = await callVertexAISearch(1, query, config);

    expect(fetchMock.called()).toBe(true);
    // Assert request body is correct
    expect(JSON.parse(fetchMock.lastCall()[1].body)).toEqual(requestJson);
    // Assert URL is correct
    expect(fetchMock.lastUrl().toLowerCase()).toBe(url.toLowerCase());
    // Assert response is correct
    expect(result).toEqual(expectedResponse);
  }
});
