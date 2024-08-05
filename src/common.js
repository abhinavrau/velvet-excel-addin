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
    ["1", "query", "", "", "link1", "link2", "link3", "TRUE", "TRUE", "TRUE", "", "", ""],
];
