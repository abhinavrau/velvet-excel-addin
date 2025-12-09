# Gemini Project Context: velvet-excel-addin

This document provides context for the Gemini agent to effectively assist with development in the `velvet-excel-addin` project.

## Project Overview

This project is an Excel Add-in named "Velvet Search Eval". It serves as an evaluation tool for Google Cloud Vertex AI Search, allowing users to interact with Vertex AI services directly from within Microsoft Excel.

## Key Technologies

- **Frontend:** Microsoft Excel Add-in (using Office.js), HTML, CSS, jQuery
- **Backend/Logic:** JavaScript (ESM), Node.js
- **AI:** Google Cloud Vertex AI
- **Testing:** Mocha
- **Formatting:** Prettier

## Project Structure

- `src/`: Contains the main application logic.
  - `excel/`: Holds the code specific to Excel integration and functionality.
  - `vertex_ai.js`: Manages interactions with the Vertex AI API.
  - `common.js`, `ui.js`: Handle common tasks and UI logic.
- `test/`: Contains the test suite, with test files corresponding to the source files. Tests are written using Mocha.
- `assets/`: Contains static assets like icons and sample datasets.
- `manifest-*.xml`: The manifest files that define the Excel Add-in's properties and commands.
- `package.json`: Defines project scripts, dependencies, and metadata.

## Common Commands

- **Run tests:** To execute the test suite, run the following command from the project root:
  ```bash
  ./node_modules/mocha/bin/mocha.js test/*.js
  ```
