export function findIndexByColumnsNameIn2DArray(array2D, searchValue) {
  for (let i = 0; i < array2D.length; i++) {
    if (array2D[i][0] === searchValue) {
      // Found the value in the first index, return the entire sub-array
      return i;
    }
  }
  // Value not found, return null or handle it as needed
  return -1;
}

export const test_search_runs_table = [
  [
    "Sheet",
    "Time of Run",
    "Project ID",
    "Search App ID",
    "Successful Queries",
    "Failed Queries",
    "% Summary Match Accuracy",
    "% First Link Match",
    "% Link in Top 2",
    "Avg. Grounding Score",
    "Num Calls to Vertex AI",
    "Time Taken",
    "Stopped Reason",
  ],
];

export const synth_qa_runs_table = [
  [
    "Sheet",
    "Time of Run",
    "Project ID",
    "Num Success",
    "Num Failed",
    "Avg. Synthetic Q&A Quality (0-5)",
  ],
];

export const summary_runs_table = [
  [
    "Sheet",
    "Time of Run",
    "Project ID",
    "Num Success",
    "Num Failed",
    "Avg. Summarization Quality (0-5)",
  ],
];
// Vertex AI Search Table Format
export const vertex_ai_search_configValues = [
  ["Config", "Value"],

  ["GCP PARAMETERS", ""],
  ["Vertex AI Search App ID", "l300-arau_1695783344117"],
  ["Vertex AI Project ID", "test_project"],
  ["Vertex AI Location", "us-central1"],

  ["SEARCH SETTINGS", ""],
  ["maxExtractiveAnswerCount (1-5)", "2"], //maxExtractiveAnswerCount
  ["maxExtractiveSegmentCount (1-5)", "0"], //maxExtractiveSegmentCount
  ["maxSnippetCount (1-5)", "0"], //maxSnippetCount
  ["useSemanticChunks (True or False)", "False"], //useSemanticChunks
  ["ignoreAdversarialQuery (True or False)", "True"], // ignoreAdversarialQuery
  ["ignoreNonSummarySeekingQuery (True or False)", "True"], // ignoreNonSummarySeekingQuery

  ["SUMMARY GENERATION SETTINGS", ""],
  ["summaryResultCount (1-5)", "2"], //summaryResultCount
  [
    "Preamble (Customized Summaries)",
    `You are an expert financial analyst.  Only use the data returned from documents. All finance numbers must be reported in billions, millions or thousands. Be brief. Answer should be no more than 2 sentences please.`,
  ],
  ["Answer Model", "gemini-2.0-flash-001/answer_gen/v1"],
  [
    "SummaryMatchingAdditionalPrompt",
    "If there are monetary numbers in the answers, they should be matched exactly.",
  ],
  ["Generate Grounding Score", "True"], //generate grounding score

  ["BATCHING SETTINGS", ""],
  ["Batch Size (1-10)", "5"], // BatchSize
  ["Time between Batches in Seconds (1-10)", "1"],
];

export const vertex_ai_search_summary_Table = [
  ["Eval Metric", "Value"],
  ["% Summary Match Accuracy", ""],
  ["% First Link Match", ""],
  ["% Link in Top 2", ""],
  ["Avg. Grounding Score", ""],
];

export const vertex_ai_search_testTableHeader = [
  [
    "ID",
    "Query",
    "Expected Summary",
    "Actual Summary",
    "Summary Match",
    "First Link Match",
    "Link in Top 2",
    "Grounding Score",
    "Expected Link 1",
    "Expected Link 2",
    "Expected Link 3",
    "Actual Link 1",
    "Actual Link 2",
    "Actual Link 3",
  ],
];

export const getAccuracyFormula = (worksheetName, columnName) =>
  `=IF(COUNTIF(${worksheetName}.TestCasesTable[${columnName}], TRUE) + COUNTIF(${worksheetName}.TestCasesTable[${columnName}], FALSE) > 0, COUNTIF(${worksheetName}.TestCasesTable[${columnName}], TRUE) / (COUNTIF(${worksheetName}.TestCasesTable[${columnName}], TRUE) + COUNTIF(${worksheetName}.TestCasesTable[${columnName}], FALSE)), 0)`;

export const getAverageFormula = (worksheetName, columnName) =>
  `=IF(COUNTA(${worksheetName}.TestCasesTable[${columnName}])>0,AVERAGE(${worksheetName}.TestCasesTable[${columnName}]), 0)`;

export var summaryMatching_prompt =
  "You will get two answers to a question, you should determine if they are semantically similar or not. ";
export var summaryMatching_examples = `examples - answer_1: I was created by X. answer_2: X created me. output:same 
  answer_1:There are 52 days in a year. answer_2: A year is fairly long. output:different `;

// Synthetic Q&A  Table Format
export const synth_q_and_a_configValues = [
  ["Config", "Value"],
  ["GCP PARAMETERS", ""],
  ["Vertex AI Project ID", "test_project"],
  ["Vertex AI Location", "us-central1"],
  ["GENERATIVE AI MODEL SETTINGS", ""],
  ["Gemini Model ID", "gemini-1.5-flash-001"],
  [
    "System Instructions",
    `Given the attached document, generate a question and an answer.The question should only be sourced from the provided the document. Do not use any other information other than the attached document. Explain your reasoning for the answer by quoting verbatim where in the document the answer is found. Return the results in JSON format.Example: {'question': 'Here is a question?', 'answer': 'Here is the answer', 'reasoning': 'Quote from document'}`,
  ],
  [
    "Prompt",
    `You are an expert in reading call center policy and procedure documents. Generate question and answer a customer would ask from a Bank using the attached document.`,
  ],
  ["Q&A EVAL SETTINGS]", ""],
  ["Generate Q & A Quality", "TRUE"],
  [
    "Q & A Quality Prompt",
    `# Instruction
You are an expert evaluator. Your task is to evaluate the quality of the responses generated by AI models.
We will provide you with the user prompt and an AI-generated responses.
You should first read the user prompt carefully for analyzing the task, and then evaluate the quality of the responses based on and rules provided in the Evaluation section below.

# Evaluation
## Metric Definition
You will be assessing question answering quality, which measures the overall quality of the answer to the question in user prompt. Pay special attention to length constraints, such as in X words or in Y sentences. The instruction for performing a question-answering task is provided in the user prompt. The response should not contain information that is not present in the context (if it is provided).

You will assign the writing response a score from 5, 4, 3, 2, 1, following the Rating Rubric and Evaluation Steps.
Give step-by-step explanations for your scoring, and only choose scores from 5, 4, 3, 2, 1.

## Criteria Definition
Instruction following: The response demonstrates a clear understanding of the question answering task instructions, satisfying all of the instruction's requirements.
Groundedness: The response contains information included only in the context if the context is present in user prompt. The response does not reference any outside information.
Completeness: The response completely answers the question with suffient detail.
Fluent: The response is well-organized and easy to read.

## Rating Rubric
5: (Very good). The answer follows instructions, is grounded, complete, and fluent.
4: (Good). The answer follows instructions, is grounded, complete, but is not very fluent.
3: (Ok). The answer mostly follows instructions, is grounded, answers the question partially and is not very fluent.
2: (Bad). The answer does not follow the instructions very well, is incomplete or not fully grounded.
1: (Very bad). The answer does not follow the instructions, is wrong and not grounded.

## Evaluation Steps
STEP 1: Assess the response in aspects of instruction following, groundedness,completeness, and fluency according to the crtieria.
STEP 2: Score based on the rubric.

Return result in JSON format. example output: { 'rating': 2 , evaluation: 'reason'}`,
  ],
  ["Q & A Quality Model ID", "gemini-2.0-flash-001"],
  ["BATCHING SETTINGS", ""],
  ["Batch Size (1-10)", "5"], // BatchSize
  ["Time between Batches in Seconds (1-10)", "1"],
];

export const synth_q_and_a_TableHeader = [
  ["ID", "GCS File URI", "Generated Question", "Expected Answer", "Q & A Quality"],
];

// Summarization  Table Format
export const summarization_configValues = [
  ["Config", "Value"],
  ["Vertex AI Project ID", "test_project"],
  ["Vertex AI Location", "us-central1"],
  ["Gemini Model ID", "gemini-1.5-flash-001"],
  ["System Instructions", ""],
  ["Prompt", "Summarize the text."],
  ["Generate Summarization Quality", "TRUE"],
  [
    "Summarization Quality Prompt",
    `# Instruction
You are an expert evaluator. Your task is to evaluate the quality of the responses generated by AI models.
We will provide you with the user input and an AI-generated responses.
You should first read the user input carefully for analyzing the task, and then evaluate the quality of the responses based on the Criteria provided in the Evaluation section below.
You will assign the response a rating following the Rating Rubric and Evaluation Steps. Give step-by-step explanations for your rating, and only choose ratings from the Rating Rubric.

# Evaluation
## Metric Definition
You will be assessing summarization quality, which measures the overall ability to summarize text. Pay special attention to length constraints, such as in X words or in Y sentences. The instruction for performing a summarization task and the context to be summarized are provided in the user prompt. The response should be shorter than the text in the context. The response should not contain information that is not present in the context.

## Criteria
Instruction following: The response demonstrates a clear understanding of the summarization task instructions, satisfying all of the instruction's requirements.
Groundedness: The response contains information included only in the context. The response does not reference any outside information.
Conciseness: The response summarizes the relevant details in the original text without a significant loss in key information without being too verbose or terse.
Fluency: The response is well-organized and easy to read.

## Rating Rubric
5: (Very good). The summary follows instructions, is grounded, is concise, and fluent.
4: (Good). The summary follows instructions, is grounded, concise, and fluent.
3: (Ok). The summary mostly follows instructions, is grounded, but is not very concise and is not fluent.
2: (Bad). The summary is grounded, but does not follow the instructions.
1: (Very bad). The summary is not grounded.

## Evaluation Steps
STEP 1: Assess the response in aspects of instruction following, groundedness, conciseness, and verbosity according to the crtieria.
STEP 2: Score based on the rubric.

Return result in JSON format. example output: { 'rating': 2 , evaluation: 'reason'}`,
  ],
  ["Summarization Quality Model ID", "gemini-2.0-flash-001"],
  ["Batch Size (1-10)", "5"], // BatchSize
  ["Time between Batches in Seconds (1-10)", "1"],
];

export const summarization_TableHeader = [["ID", "Context", "Summary", "Summary Quality"]];

export const mapGeminiSupportedMimeTypes = {
  ".flv": "video/x-flv",
  ".mov": "video/mov",
  ".mpeg": "video/mpeg",
  ".mpegps": "video/mpegps",
  ".mpg": "video/mpg",
  ".mp4": "video/mp4",
  ".webm": "video/webm",
  ".wmv": "video/wmv",
  ".3gpp": "video/3gpp",
  ".png": "image/png",
  ".jpeg": "image/jpeg",
  ".aac": "audio/aac",
  ".flac": "audio/flac",
  ".mp3": "audio/mp3",
  ".m4a": "audio/m4a", // Note: Corrected from 'MPA' to 'M4A'
  ".mpga": "audio/mpga",
  ".opus": "audio/opus",
  ".pcm": "audio/pcm",
  ".wav": "audio/wav",
  ".pdf": "application/pdf",
  "": "upsupportedMimeType",
};

export function getFileExtensionFromUri(uri) {
  const lastDotIndex = uri.lastIndexOf(".");
  if (lastDotIndex === -1) {
    return ""; // No extension found
  } else {
    return uri.substring(lastDotIndex);
  }
}
// Eval Maps
export const mapQuestionAnsweringScore = new Map();
mapQuestionAnsweringScore.set("1", "1-Very Bad");
mapQuestionAnsweringScore.set("2", "2-Bad");
mapQuestionAnsweringScore.set("3", "3-OK");
mapQuestionAnsweringScore.set("4", "4-Good");
mapQuestionAnsweringScore.set("5", "5-Very Good");
mapQuestionAnsweringScore.set(1, "1-Very Bad");
mapQuestionAnsweringScore.set(2, "2-Bad");
mapQuestionAnsweringScore.set(3, "3-OK");
mapQuestionAnsweringScore.set(4, "4-Good");
mapQuestionAnsweringScore.set(5, "5-Very Good");

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
