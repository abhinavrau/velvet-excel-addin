export const configValues = [
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
    ["SummaryMatchingAdditionalPrompt", "If there are monetary numbers in the answers, they should be matched exactly."],
    ["Batch Size (1-10)", "2"], // BatchSize
    ["Time between Batches in Seconds (1-10)", "2"],
];


export const testCaseData = [
    ["ID", "Query", "Expected Summary", "Actual Summary", "Expected Link 1", "Expected Link 2", "Expected Link 3", "Summary Match", "First Link Match", "Link in Top 2", "Actual Link 1", "Actual Link 2", "Actual Link 3"],
];

export var summaryMatching_prompt = "You will get two answers to a question, you should determine if they are semantically similar or not. ";
export var summaryMatching_examples =  " examples - answer_1: I was created by X. answer_2: X created me. output:same "
    + "answer_1:There are 52 days in a year. answer_2: A year is fairly long. output:different ";


export class NotAuthenticatedError extends Error {
    constructor(message = 'User is not authenticated') {
        super(message);
        this.name = 'NotAuthenticatedError';
        this.statusCode = 401; // Optional: HTTP status code for API errors
    }
}

export class QuotaError extends Error {
    constructor(message = 'Quota Exceeded') {
        super(message);
        this.name = 'QuotaError';
        this.statusCode = 429; // Optional: HTTP status code for API errors
    }
}