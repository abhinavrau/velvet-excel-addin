import expect from "expect";
import fetchMock from "fetch-mock";
import fs from "fs";
import { default as $, default as JQuery } from "jquery";
import sinon from "sinon";
import { NotAuthenticatedError } from "../src/common.js";
import { showStatus } from "../src/ui.js";
import { callGeminiMultitModal } from "../src/vertex_ai.js";

// mock the UI components
global.showStatus = showStatus;
global.$ = $;
global.JQuery = JQuery;

describe("When callGeminiMultiModal is called", () => {
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

  it("should return a response from Gemini for a valid prompt and file", async () => {
    // read the request from json file
    const requestData = fs.readFileSync("./test/data/multi_modal/test_multi_modal_request.json");
    const requestJson = JSON.parse(requestData);

    // Read response  json from file into variable
    const responseData = fs.readFileSync("./test/data/multi_modal/test_multi_modal_response.json");
    const responseJson = JSON.parse(responseData);

    // get the correct fields to pass in
    const prompt = requestJson.contents[0].parts[0].text;
    const fileUri = requestJson.contents[0].parts[1].fileData.fileUri;
    const mimeType = requestJson.contents[0].parts[1].fileData.mimeType;

    const config = {
      accessToken: "YOUR_ACCESS_TOKEN",
      vertexAIProjectID: "YOUR_PROJECT_ID",
      vertexAILocation: "YOUR_LOCATION",
      model: "gemini-1.5-flash-001",
      systemInstruction: requestJson.systemInstruction.parts[0].text,
      responseMimeType: "application/json",
    };

    const url = `https://${config.vertexAILocation}-aiplatform.googleapis.com/v1/projects/${config.vertexAIProjectID}/locations/${config.vertexAILocation}/publishers/google/models/${config.model}:generateContent`;
    fetchMock.postOnce(url, {
      status: 200,
      headers: { "Content-Type": `${mimeType}` },
      body: JSON.stringify(responseJson),
    });

    const result = await callGeminiMultitModal(1, prompt, fileUri, mimeType, config);

    // make sure our mock is called
    expect(fetchMock.called()).toBe(true);
    // make sure url is formatted correctly
    expect(fetchMock.lastUrl().toLowerCase()).toBe(url.toLowerCase());

    // make sure the request was sent correctly
    expect(JSON.parse(fetchMock.lastCall()[1].body)).toEqual(requestJson);
    // make sure we got the right status
    expect(result.status_code).toEqual(200);
    // make sure we got the right response
    expect(result.output.candidates[0].content.parts[0].text).toEqual(
      responseJson.candidates[0].content.parts[0].text,
    );
    
  });
  it("should fail when you get an authentication error from Vertex AI", async () => {
    const prompt = "What is the sentiment of this text?";
    const fileUri = "https://example.com/file.txt";
    const mimeType = "text/plain";
    const model_id = "gemini-1.5-flash-001";
    const config = {
      accessToken: "YOUR_ACCESS_TOKEN",
      vertexAIProjectID: "YOUR_PROJECT_ID",
      vertexAILocation: "YOUR_LOCATION",
      model: "gemini-1.5-flash-001",
      systemInstruction: null,
      responseMimeType: "application/json",
    };

    var response = {
      error: {
        message: "Call failed with status code 500 and status message: Internal Server Error",
      },
    };
    const url = `https://${config.vertexAILocation}-aiplatform.googleapis.com/v1/projects/${config.vertexAIProjectID}/locations/${config.vertexAILocation}/publishers/google/models/${config.model}:generateContent`;
    fetchMock.postOnce(url, {
      status: 401,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(response),
    });

    try {
      const result = await callGeminiMultitModal(1, prompt, fileUri, mimeType, config);
      assert.fail();
    } catch (err) {
      expect(fetchMock.called()).toBe(true);
      expect(fetchMock.lastUrl().toLowerCase()).toBe(url.toLowerCase());
      expect(err).toBeInstanceOf(NotAuthenticatedError);
    }
  });

  it("should fail when fetch throws exception", async () => {
    const prompt = "What is the sentiment of this text?";
    const fileUri = "https://example.com/file.txt";
    const mimeType = "text/plain";
    const model_id = "gemini-1.5-flash-001";
    const config = {
      accessToken: "YOUR_ACCESS_TOKEN",
      vertexAIProjectID: "YOUR_PROJECT_ID",
      vertexAILocation: "YOUR_LOCATION",
      model: "gemini-1.5-flash-001",
      systemInstruction: null,
      responseMimeType: "application/json",
    };

    const url = `https://${config.vertexAILocation}-aiplatform.googleapis.com/v1/projects/${config.vertexAIProjectID}/locations/${config.vertexAILocation}/publishers/google/models/${config.model}:generateContent`;
    fetchMock.postOnce(url, {
      throws: new Error("Mocked error"),
    });

    try {
      const result = await callGeminiMultitModal(1, prompt, fileUri, mimeType, config);
      assert.fail();
    } catch (err) {
      expect(fetchMock.called()).toBe(true);
      expect(fetchMock.lastUrl().toLowerCase()).toBe(url.toLowerCase());
      expect(err.message).toBe("Mocked error");
    }
  });
});
