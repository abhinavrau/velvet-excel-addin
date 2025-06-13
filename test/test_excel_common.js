import expect from "expect";
import { vertex_ai_search_configValues } from "../src/common.js";
import { populateConfigTableValues } from "../src/excel/excel_common.js";

describe("populateConfigTableValues", () => {
  let originalConfigValues;

  beforeEach(() => {
    // Create a deep copy of the original array to reset it before each test
    originalConfigValues = JSON.parse(JSON.stringify(vertex_ai_search_configValues));
  });

  afterEach(() => {
    // Restore the original array after each test
    vertex_ai_search_configValues.length = 0;
    vertex_ai_search_configValues.push(...originalConfigValues);
  });

  it("should populate the config table with values from the config object", () => {
    const config = {
      vertexAISearchAppId: "my-search-app",
      vertexAIProjectID: "my-gcp-project",
      vertexAILocation: "us-central1",
      model: "gemini-pro",
      preamble: "This is a test preamble.",
      summaryResultCount: 3,
      genereateGrounding: "TRUE",
      useSemanticChunks: "FALSE",
      ignoreAdversarialQuery: "TRUE",
      ignoreNonSummarySeekingQuery: "FALSE",
      summaryMatchingAdditionalPrompt: "Additional prompt for matching.",
      batchSize: 5,
      timeBetweenCallsInSec: 2,
      maxSnippetCount: 2,
      extractiveContentSpec: {
        maxExtractiveAnswerCount: 3,
        maxExtractiveSegmentCount: 4,
      },
    };

    populateConfigTableValues(vertex_ai_search_configValues, config);

    
    expect(
      vertex_ai_search_configValues.find((row) => row[0] === "Vertex AI Search App ID")[1],
    ).toBe("my-search-app");
    expect(vertex_ai_search_configValues.find((row) => row[0] === "Vertex AI Project ID")[1]).toBe(
      "my-gcp-project",
    );
    expect(vertex_ai_search_configValues.find((row) => row[0] === "Vertex AI Location")[1]).toBe(
      "us-central1",
    );
    expect(vertex_ai_search_configValues.find((row) => row[0] === "Answer Model")[1]).toBe(
      "gemini-pro",
    );
    expect(
      vertex_ai_search_configValues.find((row) => row[0] === "Preamble (Customized Summaries)")[1],
    ).toBe("This is a test preamble.");
    expect(
      vertex_ai_search_configValues.find((row) => row[0] === "summaryResultCount (1-5)")[1],
    ).toBe(3);
    expect(
      vertex_ai_search_configValues.find((row) => row[0] === "Generate Grounding Score")[1],
    ).toBe("TRUE");
    expect(
      vertex_ai_search_configValues.find(
        (row) => row[0] === "useSemanticChunks (True or False)",
      )[1],
    ).toBe("FALSE");
    expect(
      vertex_ai_search_configValues.find(
        (row) => row[0] === "ignoreAdversarialQuery (True or False)",
      )[1],
    ).toBe("TRUE");
    expect(
      vertex_ai_search_configValues.find(
        (row) => row[0] === "ignoreNonSummarySeekingQuery (True or False)",
      )[1],
    ).toBe("FALSE");
    expect(
      vertex_ai_search_configValues.find((row) => row[0] === "SummaryMatchingAdditionalPrompt")[1],
    ).toBe("Additional prompt for matching.");
    expect(vertex_ai_search_configValues.find((row) => row[0] === "Batch Size (1-10)")[1]).toBe(5);
    expect(
      vertex_ai_search_configValues.find(
        (row) => row[0] === "Time between Batches in Seconds (1-10)",
      )[1],
    ).toBe(2);
    expect(vertex_ai_search_configValues.find((row) => row[0] === "maxSnippetCount (1-5)")[1]).toBe(
      2,
    );
    expect(
      vertex_ai_search_configValues.find((row) => row[0] === "maxExtractiveAnswerCount (1-5)")[1],
    ).toBe(3);
    expect(
      vertex_ai_search_configValues.find((row) => row[0] === "maxExtractiveSegmentCount (1-5)")[1],
    ).toBe(4);
  });

  it("should handle null values in the config object", () => {
    const config = {
      vertexAISearchAppId: null,
      vertexAIProjectID: null,
      vertexAILocation: null,
      model: null,
      preamble: null,
      summaryResultCount: null,
      genereateGrounding: null,
      useSemanticChunks: null,
      ignoreAdversarialQuery: null,
      ignoreNonSummarySeekingQuery: null,
      summaryMatchingAdditionalPrompt: null,
      batchSize: null,
      timeBetweenCallsInSec: null,
      maxSnippetCount: null,
      extractiveContentSpec: {
        maxExtractiveAnswerCount: null,
        maxExtractiveSegmentCount: null,
      },
    };

    populateConfigTableValues(vertex_ai_search_configValues, config);

    expect(
      vertex_ai_search_configValues.find((row) => row[0] === "Vertex AI Search App ID")[1],
    ).toBe(0);
    expect(vertex_ai_search_configValues.find((row) => row[0] === "Vertex AI Project ID")[1]).toBe(
      0,
    );
    expect(vertex_ai_search_configValues.find((row) => row[0] === "Vertex AI Location")[1]).toBe(0);
    expect(vertex_ai_search_configValues.find((row) => row[0] === "Answer Model")[1]).toBe(0);
    expect(
      vertex_ai_search_configValues.find((row) => row[0] === "Preamble (Customized Summaries)")[1],
    ).toBe(0);
    expect(
      vertex_ai_search_configValues.find((row) => row[0] === "summaryResultCount (1-5)")[1],
    ).toBe(0);
    expect(
      vertex_ai_search_configValues.find((row) => row[0] === "Generate Grounding Score")[1],
    ).toBe(0);
    expect(
      vertex_ai_search_configValues.find(
        (row) => row[0] === "useSemanticChunks (True or False)",
      )[1],
    ).toBe(0);
    expect(
      vertex_ai_search_configValues.find(
        (row) => row[0] === "ignoreAdversarialQuery (True or False)",
      )[1],
    ).toBe(0);
    expect(
      vertex_ai_search_configValues.find(
        (row) => row[0] === "ignoreNonSummarySeekingQuery (True or False)",
      )[1],
    ).toBe(0);
    expect(
      vertex_ai_search_configValues.find((row) => row[0] === "SummaryMatchingAdditionalPrompt")[1],
    ).toBe(0);
    expect(vertex_ai_search_configValues.find((row) => row[0] === "Batch Size (1-10)")[1]).toBe(0);
    expect(
      vertex_ai_search_configValues.find(
        (row) => row[0] === "Time between Batches in Seconds (1-10)",
      )[1],
    ).toBe(0);
    expect(vertex_ai_search_configValues.find((row) => row[0] === "maxSnippetCount (1-5)")[1]).toBe(
      0,
    );
    expect(
      vertex_ai_search_configValues.find((row) => row[0] === "maxExtractiveAnswerCount (1-5)")[1],
    ).toBe(0);
    expect(
      vertex_ai_search_configValues.find((row) => row[0] === "maxExtractiveSegmentCount (1-5)")[1],
    ).toBe(0);
  });

  it("should handle undefined config object", () => {
    populateConfigTableValues(vertex_ai_search_configValues, undefined);
    // If no error is thrown, the test passes.
    // We can add more specific checks if needed, but for now, just ensuring no error is sufficient.
  });

  it("should handle null config object", () => {
    populateConfigTableValues(vertex_ai_search_configValues, null);
    // If no error is thrown, the test passes.
    // We can add more specific checks if needed, but for now, just ensuring no error is sufficient.
  });

  it("should handle undefined vertex_ai_search_configValues", () => {
    populateConfigTableValues(undefined, {});
    // If no error is thrown, the test passes.
    // We can add more specific checks if needed, but for now, just ensuring no error is sufficient.
  });

  it("should handle null vertex_ai_search_configValues", () => {
    populateConfigTableValues(null, {});
    // If no error is thrown, the test passes.
    // We can add more specific checks if needed, but for now, just ensuring no error is sufficient.
  });
});
