# MS Teams Email Sentiment Analyzer & Summarizer

This MS Teams app allows users to view their emails and analyze their sentiments using Azure Open AI Text Analytics. It also provides the ability to summarize the emails using Azure Open AI Completion model.

## Features

- View Emails: Display a list of user's emails from their Outlook inbox.
- Sentiment Analysis: Analyze the sentiment (positive, negative, neutral) of each email.
- Email Summarization: Generate a summary of the email content using the completion model.
- Filter by Sentiment: Filter the emails based on sentiment (positive, negative, neutral).

![MS Teams App Demo](Animation.gif)

## Configuration

1. Create an Azure Text Analytics resource and obtain the API endpoint and access key.
2. Create an Azure Open AI resource and obtain the API endpoint and access key for the Completion model.
3. Update the following variables with your Azure API and Open AI credentials:
   - `textAnalyticsEndpoint`: Azure Text Analytics API endpoint
   - `textAnalyticsApiKey`: Azure Text Analytics access key
   - `summaryEndpoint`: Open AI Completion API endpoint

## Prerequisites

- [Node.js](https://nodejs.org/), supported versions: 16, 18
- An M365 account. If you do not have M365 account, apply one from [M365 developer program](https://developer.microsoft.com/microsoft-365/dev-program)
- [Set up your dev environment for extending Teams apps across Microsoft 365](https://aka.ms/teamsfx-m365-apps-prerequisites)
  > Please note that after you enrolled your developer tenant in Office 365 Target Release, it may take couple days for the enrollment to take effect.
- [Teams Toolkit Visual Studio Code Extension](https://aka.ms/teams-toolkit) version 5.0.0 and higher or [TeamsFx CLI](https://aka.ms/teamsfx-cli)

## Getting Started

Follow below instructions to get started with this application template for local debugging.

### Test your application with Visual Studio Code

1. Press `F5` or use the `Run and Debug Activity Panel` in Visual Studio Code.
1. Select a target Microsoft 365 application where the personal tabs can run: `Debug in Teams`, `Debug in Outlook` or `Debug in the Microsoft 365 app` and click the `Run and Debug` green arrow button.

### Test your application with TeamsFx CLI

1. Executing the command `teamsfx provision --env local` in your project directory.
1. Executing the command `teamsfx deploy --env local` in your project directory.
1. Executing the command `teamsfx preview --env local --m365-host <m365-host>` in your project directory, where options for `m365-host` are `teams`, `outlook` or `office`.

## References

- [Extend a Teams personal tabs across Microsoft 365](https://docs.microsoft.com/microsoftteams/platform/m365-apps/extend-m365-teams-personal-tab?tabs=manifest-teams-toolkit)
- [Teams Toolkit Documentations](https://docs.microsoft.com/microsoftteams/platform/toolkit/teams-toolkit-fundamentals)
- [Teams Toolkit CLI](https://docs.microsoft.com/microsoftteams/platform/toolkit/teamsfx-cli)
- [TeamsFx SDK](https://docs.microsoft.com/microsoftteams/platform/toolkit/teamsfx-sdk)
- [Teams Toolkit Samples](https://github.com/OfficeDev/TeamsFx-Samples)
