import React, { useState } from "react";
import { Stack, Text, Checkbox } from "@fluentui/react";
import "./Email.css";

const Email = ({ sender, subject, body, summary, sentiment }) => {
  const [showSummary, setShowSummary] = useState(false);

  const handleCheckboxChange = (event) => {
    setShowSummary(event.target.checked);
  };

  return (
    <div className="email">
      <Stack verticalAlign="start" tokens={{ childrenGap: 10 }}>
        <div className="email-header">
          <Text
            variant="mediumPlus"
            className={`email-sentiment ${
              sentiment === "Positive" ? "positive" : "negative"
            }`}
          >
            {sentiment}
          </Text>
          <Text variant="mediumPlus" className="email-sender">
            From: {sender.emailAddress.name}
          </Text>
        </div>
        <div className="email-details">
          <div className="email-subject-container">
            <Text variant="large" className="email-subject-label">
              Subject:
            </Text>
            <Text className="email-subject-value">{subject}</Text>
          </div>
          <div className="email-body-container">
            <Checkbox
              label="Summarize"
              checked={showSummary}
              onChange={handleCheckboxChange}
              className="email-summary-checkbox"
            />
            <Text className="email-body">{showSummary ? summary : body}</Text>
          </div>
        </div>
      </Stack>
    </div>
  );
};

export default Email;
