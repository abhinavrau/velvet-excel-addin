

![](images/velvet_excel.png)

# Velvet Excel
Run and measure [Vertex AI Search Agent](https://cloud.google.com/enterprise-search) accuracy using Excel. 

## Note
This is not an official Google product

## Why?
- 🧑‍💻 Run your search acceptance test cases without leaving Excel.
- ✅ Enables Business and Compliance teams to measure search accuracy without any training or developer skills using the tool they use every day.
- 🗒️ Allows comparing of different test runs using different settings to determine the correct values for a specific use case.
- 🚀 Accelerates search application validation from weeks/months to hours.
- 🤖 Coz testing manually is hard, boring and no fun!

## Features

- Excel Office Add-on. Works with Excel Web (Microsoft 365) and Excel Desktop.
- Creates a test case table that measures accuracy for the following metrics:
    - **Summary Match (True/False)**: Semantically match the actual summary returned with the expected summary. This uses the PaLM2 model to measure semantic similarity using [the following prompt](https://github.com/abhinavrau/velvet-excel-addin/blob/f532763488eb03c93b24e372ab650997e1acbee0/src/vertex_ai.js#L66).
    - **First Link Match (True/False)**: The actual first document link returned matches the expected document link. 
    - **Link in Top 2**: Match if the top 2 actual document links returned are in the top 2 expected links
- Ability to run hundreds of test cases in batch mode.

- Ability to change the following Vertex AI Search Agent settings from within Excel:
    - Search Types (any one of):
        - [Extractive Answers](https://cloud.google.com/generative-ai-app-builder/docs/snippets#extractive-answers): maxExtractiveAnswerCount (1-5)
        - [Extractive Segments](https://cloud.google.com/generative-ai-app-builder/docs/snippets#extractive-segments): maxExtractiveSegmentCount  (1-5)
        - [Snippets](https://cloud.google.com/generative-ai-app-builder/docs/snippets#snippets): maxSnippetCount  (1-5)
    - Search Summary Settings:
        - [Summary Result Count](https://cloud.google.com/generative-ai-app-builder/docs/get-search-summaries#get-search-summary): (1-5)
        - [Use Semantic Chunks](https://cloud.google.com/generative-ai-app-builder/docs/get-search-summaries#semantic-chunks):  (True/False)
        - [Summary Model](https://cloud.google.com/generative-ai-app-builder/docs/answer-generation-models) gemini-1.0-pro-002/answer_gen/v1 (default)
            
        - [Ignore Adversarial Query](https://cloud.google.com/generative-ai-app-builder/docs/get-search-summaries#ignore-adversarial-queries):   (True/False)
        - [Ignore Non Summary Seeking Query](https://cloud.google.com/generative-ai-app-builder/docs/get-search-summaries#ignore-non-summary-seeking-queries):   (True/False)
    - Other
        - Summary Match Additional Prompt: Additional prompt to pass to the PaLM2 model for semantic similarity matching. Useful when the need to match exact monetary quantities (millions, billions, etc).

## Installation
Only [sideloading](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/test-debug-office-add-ins#sideload-an-office-add-in-for-testing) is supported at this time.


### Excel Web (Microsoft 365)

- Download the [manfest.xml](manifest.xml) file to your local machine.
- Open Excel in your browser.
- Click on the "Add-ins" button and select "More Add-Ins".

![](images/Web_Install_step_1.png)
- Click on the "Upload My Add-in" link 

![](images/Web_Install_step_2.png)
- Select the manifest.xml file from your local machine and click on the "Open" button.

![](images/Web_Install_step_3.png)
- You should now see the Velvet Add-in listed in the Add-ins menu.

![](images/Web_Install_step_4.png)

To remove it just clear your browser cache.

### Excel Desktop 

- Download the [manfest.xml](manifest.xml) file to your local machine.
- Place the manifest.xml file in the following location:
    - **Windows**: Follow instructions [here](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins).
    - **Mac**:
        - Use Finder to sideload the manifest file. Open Finder and then enter Command+Shift+G to open the Go to folder dialog.
        - Navigate to the following location: /Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef
        - If the wef folder doesn't exist on your computer, create it.
        - Save the manifest.xml file in the wef folder.
- Open Excel.
- Confirm that the Velvet Add-in is listed in the Home Ribbon.





