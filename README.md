# Test Automation Agent

A local desktop friendly web app that turns Azure DevOps test cases into Playwright test drafts. Point it at a TypeScript Playwright project and it uses that project's fixtures, tests, helpers, and source as generation context. It supports Quick Generate scaffolds and Smart Generate with GitHub Copilot CLI or OpenAI Codex CLI.

The project is designed to be adapted to your own Azure DevOps organization and Playwright framework. It does not include a target application, test credentials, or framework code.

## Requirements

- Node.js 22 or newer and npm.
- An Azure DevOps organization and project with test cases.
- A TypeScript Playwright project with its dependencies installed, including TypeScript.
- GitHub Copilot CLI or OpenAI Codex CLI for Smart Generate. Quick Generate does not require an AI CLI.
- Optional: Code Review Graph and its Python package for graph assisted source lookup.

## Run locally

```sh
npm ci
npm start
```

Open the loopback URL printed in the terminal (normally <http://127.0.0.1:4173>). Keep the terminal open while using the app and press Ctrl+C to stop it. The app listens on loopback only; it is not a hosted service.

## Configure your project

Open **Workspace setup** and provide:

1. Your Azure DevOps organization and project.
2. The absolute path to your TypeScript Playwright project.
3. The folder where generated test drafts should be written.
4. The CLI command and model to use for Smart Generate.

Create an Azure DevOps personal access token with only the work item read permissions needed by your organization. Save it in the app; it uses the operating system credential store. The token is not written to this repository or the app's settings file.

Sign in to Copilot CLI or Codex CLI separately using the CLI's supported sign-in process. The app invokes the selected CLI locally. For better results, keep project setup guidance in your own `README.md`, `AGENTS.md`, and nearby tests so the agent can discover your conventions.

If you use Code Review Graph, enter the Python executable for the environment where its package is installed. The project path must contain the graph database for graph lookups. Local RAG can retrieve project source and test examples without that integration.

Settings are stored in the current user's application configuration directory. The Azure DevOps token is kept in the OS credential store.

## Generate and review

1. Save workspace settings and verify the Azure DevOps connection.
2. Enter a test case ID and fetch the work item.
3. Review the title, steps, and attachments. Add only scenario files needed for this test.
4. Choose **Quick Generate** for a scaffold or **Smart Generate** for a CLI generated draft.
5. Review the saved TypeScript file, adapt it to your project's fixture conventions, and run the appropriate Playwright checks yourself.

Generated code is a draft. The app performs limited static checks, but it does not run generated tests or confirm application behavior. Smart Generate uses the configured AI CLI and sends it the selected test case, attachments, and relevant project context. Do not submit material you are not allowed to share with that CLI.

## Features

- Fetches Azure DevOps test case titles, descriptions, steps, and attachments.
- Generates reviewable Playwright TypeScript scaffolds or AI assisted drafts.
- Retrieves relevant project source and existing tests with a local lexical index.
- Stores the Azure DevOps PAT with the operating system credential manager.
- Supports optional read only Code Review Graph context.
- Provides draft refinement, progress reporting, cancellation, and a local output preview.
- Includes a generic explainer video at [`media/automation-agent-explainer.mp4`](media/automation-agent-explainer.mp4).

## Project files

- `server.mjs` — local HTTP server, Azure DevOps connection, source retrieval, generation orchestration, and validation.
- `app.js`, `index.html`, `styles.css` — browser interface.
- `generation-policy.mjs` — shared generation and refinement instructions.
- `scenario-workbook.mjs` — extracts matching scenario data from selected Excel/CSV attachments.
- `source-evidence.mjs` — reads relevant TypeScript project sources for static context.
- `generated-test-check.mjs` — performs limited TypeScript and Playwright source checks.
- `outlook-draft.mjs`, `open-outlook.ps1` — optional local Outlook feedback draft support.

## Privacy and safety

- The web server binds to `127.0.0.1`.
- The Azure DevOps PAT is stored in the OS credential store and is not sent to the AI CLI.
- The app does not execute generated tests. Review generated source and attachments before using them.
- Temporary prompts and attachments are removed after generation; selected attachments are held in local memory while the app is running.
- Generated files are written only to the output directory you configure.
