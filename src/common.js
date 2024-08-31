// Vertex AI Search Table Format

export const vertex_ai_search_configValues = [
  ["Config", "Value"],
  ["Vertex AI Search Project Number", "384473000457"],
  ["Vertex AI Search DataStore Name", "alphabet-pdfs_1695783402380"],
  ["Vertex AI Project ID", "argolis-arau"],
  ["Vertex AI Location", "us-central1"],
  ["maxExtractiveAnswerCount (1-5)", "2"], //maxExtractiveAnswerCount
  ["maxExtractiveSegmentCount (1-5)", "0"], //maxExtractiveSegmentCount
  ["maxSnippetCount (1-5)", "0"], //maxSnippetCount
  ["Preamble (Customized Summaries)", ""],
  ["Summarization Model", "gemini-1.0-pro-001/answer_gen/v1"],
  ["summaryResultCount (1-5)", "2"], //summaryResultCount
  ["useSemanticChunks (True or False)", "False"], //useSemanticChunks
  ["ignoreAdversarialQuery (True or False)", "True"], // ignoreAdversarialQuery
  ["ignoreNonSummarySeekingQuery (True or False)", "True"], // ignoreNonSummarySeekingQuery
  [
    "SummaryMatchingAdditionalPrompt",
    "If there are monetary numbers in the answers, they should be matched exactly.",
  ],
  ["Batch Size (1-10)", "2"], // BatchSize
  ["Time between Batches in Seconds (1-10)", "2"],
];

export const vertex_ai_search_testTableHeader = [
  [
    "ID",
    "Query",
    "Expected Summary",
    "Actual Summary",
    "Expected Link 1",
    "Expected Link 2",
    "Expected Link 3",
    "Summary Match",
    "First Link Match",
    "Link in Top 2",
    "Actual Link 1",
    "Actual Link 2",
    "Actual Link 3",
  ],
];

export var summaryMatching_prompt =
  "You will get two answers to a question, you should determine if they are semantically similar or not. ";
export var summaryMatching_examples =
  " examples - answer_1: I was created by X. answer_2: X created me. output:same " +
  "answer_1:There are 52 days in a year. answer_2: A year is fairly long. output:different ";

// Synthetic Q&A  Table Format
export const synth_q_and_a_configValues = [
  ["Config", "Value"],
  ["Vertex AI Project ID", "argolis-arau"],
  ["Vertex AI Location", "us-central1"],
  ["Gemini Model ID", "gemini-1.5-flash-001"],
  [
    "System Instructions",
    "You are an expert in reading call center policy and procedure documents." +
      "Given the attached document, generate a question and answer that customers are likely to ask a call center agent." +
      "The question should only be sourced from the provided the document.Do not use any other information other than the attached document. " +
      "Explain your reasoning for the answer by quoting verbatim where in the document the answer is found. Return the results in JSON format.Example: " +
      "{'question': 'Here is a question?', 'answer': 'Here is the answer', 'reasoning': 'Quote from document'}",
  ],
  ["Batch Size (1-10)", "4"], // BatchSize
  ["Time between Batches in Seconds (1-10)", "2"],
];

export const synth_q_and_a_TableHeader = [
  [
    "ID",
    "GCS File URI",
    "Mime Type",
    "Generated Question",
    "Expected Answer",
    "Reasoning",
    "Status",
    "Response Time",
  ],
];

// Summarization  Table Format
export const summarization_configValues = [
  ["Config", "Value"],
  ["Vertex AI Project ID", "argolis-arau"],
  ["Vertex AI Location", "us-central1"],
  ["Gemini Model ID", "gemini-1.5-flash-001"],
  ["Instructions", "Summarize the text."],
  ["summarization_quality", "TRUE"],
  ["summarization_helpfulness", "TRUE"],
  ["summarization_verbosity", "FALSE"],
  ["groundedness", "FALSE"],
  ["fulfillment", "FALSE"],
  ["Batch Size (1-10)", "4"], // BatchSize
  ["Time between Batches in Seconds (1-10)", "2"],
];

export const summarization_TableHeader = [
  [
    "ID",
    "Context",
    "Summary",
    "summarization_quality",
    "summarization_helpfulness",
    "summarization_verbosity",
    "groundedness",
    "fulfillment",
  ],
];

// Eval Maps

export const mapSummaryQualityScore = new Map();
mapSummaryQualityScore.set(1, "1-Very Bad");
mapSummaryQualityScore.set(2, "2-Bad");
mapSummaryQualityScore.set(3, "3-OK");
mapSummaryQualityScore.set(4, "4-Good");
mapSummaryQualityScore.set(5, "5-Very Good");

export const mapSummaryHelpfulnessScore = new Map();
mapSummaryHelpfulnessScore.set(1, "1-Unhelpful");
mapSummaryHelpfulnessScore.set(2, "2-Somewhat Unhelpful");
mapSummaryHelpfulnessScore.set(3, "3-Neutral");
mapSummaryHelpfulnessScore.set(4, "4-Somewhat Helpful");
mapSummaryHelpfulnessScore.set(5, "5-Helpful");

export const mapSummaryVerbosityScore = new Map();
mapSummaryVerbosityScore.set(-2, "-2-Terse");
mapSummaryVerbosityScore.set(-1, "-1-Somewhat Terse");
mapSummaryVerbosityScore.set(0, "0-Optimal");
mapSummaryVerbosityScore.set(1, "1-Somewhat Verbose");
mapSummaryVerbosityScore.set(2, "2-Verbose");

export const mapTextgenGroundednessScore = new Map();
mapTextgenGroundednessScore.set(0, "0-Ungrounded");
mapTextgenGroundednessScore.set(1, "1-Grounded");

export const mapTextgenFulfillmentScore = new Map();
mapTextgenFulfillmentScore.set(1, "1-No fulfillment");
mapTextgenFulfillmentScore.set(2, "2-Poor fulfillment");
mapTextgenFulfillmentScore.set(3, "3-Some fulfillment");
mapTextgenFulfillmentScore.set(4, "4-Good fulfillment");
mapTextgenFulfillmentScore.set(5, "5-Complete fulfillmentl");

export class VertexAIError extends Error {
  constructor(id, message = "Processing Error") {
    super(message);
    this.name = "VertexAIError";
    this.id = id;
  }
}

export class NotAuthenticatedError extends Error {
  constructor(id, message = "User is not authenticated") {
    super(message);
    this.name = "NotAuthenticatedError";
    this.id = id;
    this.statusCode = 401; // Optional: HTTP status code for API errors
  }
}

export class QuotaError extends Error {
  constructor(id, message = "Quota Exceeded") {
    super(message);
    this.name = "QuotaError";
    this.id = id;
    this.statusCode = 429; // Optional: HTTP status code for API errors
  }
}

export class PermissionDeniedError extends Error {
  constructor(id, message = "Permission Denied") {
    super(message);
    this.name = "PermissionDeniedError";
    this.id = id;
    this.statusCode = 403; // Optional: HTTP status code for API errors
  }
}

export class ResourceNotFoundError extends Error {
  constructor(id, message = "Resource Not Found ") {
    super(message);
    this.name = "ResourceNotFoundError";
    this.id = id;
    this.statusCode = 404; // Optional: HTTP status code for API errors
  }
}

