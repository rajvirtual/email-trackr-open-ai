import React, { useContext, useState } from "react";
import {
  Text,
  Stack,
  Dropdown,
  Spinner,
  SpinnerSize,
  PrimaryButton,
} from "@fluentui/react";
import Email from "./Email";
import { TeamsFxContext } from "./Context";
import { useData } from "@microsoft/teamsfx-react";
import { useGraphWithCredential } from "@microsoft/teamsfx-react";
import { Providers, ProviderState } from "@microsoft/mgt-element";
import { TeamsFxProvider } from "@microsoft/mgt-teamsfx-provider";

import "./Main.css";
import {
  TextAnalyticsClient,
  AzureKeyCredential,
} from "@azure/ai-text-analytics";
import axios from "axios";

const MainPage = () => {
  // Enter your own Open AI endpoints and API Keys
  const textAnalyticsApiKey = "";
  const textAnalyticsEndpoint = "";

  const summaryEndpoint = "";

  const [filter, setFilter] = useState("All");

  const handleFilterChange = (event, option) => {
    setFilter(option.text);
  };

  const { teamsUserCredential } = useContext(TeamsFxContext);
  const textAnalyticsClient = new TextAnalyticsClient(
    textAnalyticsEndpoint,
    new AzureKeyCredential(textAnalyticsApiKey)
  );
  const { loading, data, error } = useData(async () => {
    if (teamsUserCredential) {
      const userInfo = await teamsUserCredential.getUserInfo();
      return userInfo;
    }
  });

  const userName = loading || error ? "" : data?.displayName;

  const {
    loading: graphLoading,
    error: graphError,
    data: graphData,
    reload,
  } = useGraphWithCredential(
    async (graph, teamsUserCredential, scope) => {
      // Call graph api directly to get user profile information
      const profile = await graph.api("/me").get();
      let emails = [];

      // Initialize Graph Toolkit TeamsFx provider
      const provider = new TeamsFxProvider(teamsUserCredential, scope);
      Providers.globalProvider = provider;
      Providers.globalProvider.setState(ProviderState.SignedIn);

      try {
        const today = new Date();
        const twoDaysAgo = new Date(
          today.setDate(today.getDate() - 2)
        ).toISOString();

        const result = await graph
          .api("me/mailFolders/inbox/messages")
          .header("Prefer", 'outlook.body-content-type="text"')
          // .filter(`isRead eq false and receivedDateTime ge ${twoDaysAgo}`)
          // .top(20)
          .select("id, subject,sender, body")
          .get();

        if (result?.value) {
          emails = await Promise.all(
            result?.value.map(async (val) => {
              let summary = await summarizeEmail(val);
              let sentiment = await analyzeSentiments(val);
              return {
                id: val.id,
                subject: val.subject,
                sentiment: sentiment,
                summary: summary,
                body: val.body.content,
                sender: val.sender,
              };
            })
          );
        }
      } catch (error) {
        console.error("Error in Graph request: " + error);
      }
      return { profile, emails };
    },
    { scope: ["User.Read", "mail.read"], credential: teamsUserCredential }
  );

  const analyzeSentiments = async (email) => {
    let sentiment = "";

    let fullEmail = `Subject : ${email.subject} ${email.body.content}`;

    const documents = [{ id: email.id, text: fullEmail }];

    const results = await textAnalyticsClient.analyzeSentiment(documents);

    if (results && results[0].confidenceScores) {
      sentiment =
        results[0].confidenceScores.negative > 0.49 ? "Negative" : "Positive";
    }
    return sentiment;
  };

  const summarizeEmail = async (email) => {
    const requestBody = {
      prompt: `Summarize this text: ${email.body.content}`,
      max_tokens: 50,
      temperature: 0.4,
      top_p: 1,
      frequency_penalty: 0.2,
      presence_penalty: 0.2,
    };

    const headers = {
      "Content-Type": "application/json",
    };

    try {
      const response = await axios.post(summaryEndpoint, requestBody, {
        headers,
      });

      const { choices } = response.data;

      const summaryText = choices[0].text.trim();

      return summaryText;
    } catch (error) {
      console.error("Error:", error);
    }
  };

  const filterEmails = (emails, filter) => {
    if (filter === "All") {
      return emails;
    } else if (filter === "Positive") {
      return emails.filter((email) => email.sentiment === "Positive");
    } else if (filter === "Negative") {
      return emails.filter((email) => email.sentiment === "Negative");
    }
    return emails;
  };

  const sortBySentiment = (a, b) => {
    if (a.sentiment === "Negative" && b.sentiment !== "Negative") {
      return -1; // Move negative sentiment emails to the top
    }
    if (a.sentiment !== "Negative" && b.sentiment === "Negative") {
      return 1; // Move positive sentiment emails below negative sentiment emails
    }
    return 0; // Preserve the original order for emails with the same sentiment
  };

  if (graphLoading) {
    return <Spinner size={SpinnerSize.large} label="Loading..." />;
  }
  const filteredEmails = filterEmails(graphData.emails, filter);

  const sortedEmails = filteredEmails.sort(sortBySentiment);

  console.log("Sorted emails", sortedEmails);

  return (
    <Stack tokens={{ childrenGap: 20 }}>
      <div className="filter-header">
        <Text
          className="user-header"
          variant="xLarge"
        >{`Hello, ${userName}`}</Text>
        <div className="filter-dropdown">
          <Text className="filter-label">Filter by Sentiment:</Text>
          <Dropdown
            className="sentiment-filter"
            options={[
              { key: "All", text: "All" },
              { key: "Positive", text: "Positive" },
              { key: "Negative", text: "Negative" },
            ]}
            selectedKey={filter} // Set the selected filter
            onChange={handleFilterChange} // Handle filter change
          />
        </div>
        <PrimaryButton
          className="refresh-button"
          onClick={reload}
          style={{ backgroundColor: "blue", color: "white" }}
        >
          Refresh
        </PrimaryButton>
      </div>
      <Stack tokens={{ childrenGap: 10 }}>
        {graphData &&
          sortedEmails.map((email) => (
            <Email
              key={email.id}
              subject={email.subject}
              body={email.body}
              summary={email.summary}
              sender={email.sender}
              sentiment={email.sentiment}
            />
          ))}
      </Stack>
    </Stack>
  );
};

export default MainPage;
