import "./Graph.css";
import { useGraphWithCredential } from "@microsoft/teamsfx-react";
import { Providers, ProviderState } from "@microsoft/mgt-element";
import { TeamsFxProvider } from "@microsoft/mgt-teamsfx-provider";
import { Button } from "@fluentui/react-components";
import { Design } from "./Design";
import { PersonCardFluentUI } from "./PersonCardFluentUI";
import { PersonCardGraphToolkit } from "./PersonCardGraphToolkit";
import { useContext ,useState } from "react";
import { TeamsFxContext } from "../Context";
import Email from "../Email";

export function Graph() {
  const { teamsUserCredential } = useContext(TeamsFxContext);
  const [emails, setEmails] = useState([]);
  const [graphError, setGraphError] = useState([]);
  const [rawValue, SetRawValue] = useState([]);
  const { loading, error, data, reload } = useGraphWithCredential(
    async (graph, teamsUserCredential, scope) => {
      // Call graph api directly to get user profile information
      const profile = await graph.api("/me").get();

      // Initialize Graph Toolkit TeamsFx provider
      const provider = new TeamsFxProvider(teamsUserCredential, scope);
      Providers.globalProvider = provider;
      Providers.globalProvider.setState(ProviderState.SignedIn);

      let photoUrl = "";
      try {
        // const photo = await graph.api("/me/photo/$value").get();
        // photoUrl = URL.createObjectURL(photo);

        const today = new Date();
        const twoDaysAgo = new Date(today.setDate(today.getDate() - 2)).toISOString();
    
        const result = await graph.api('/me/messages')
          // .filter(`isRead eq false and receivedDateTime ge ${twoDaysAgo}`)
          // .top(10)
          .select('id, subject,sender, bodyPreview')
          .get();

          SetRawValue(result);
          setEmails(result.value);

      } catch(error) {
        setGraphError(error)
        // Could not fetch photo from user's profile, return empty string as placeholder.
      }
      return { profile, photoUrl };
    },
    { scope: ["User.Read","mail.read"], credential: teamsUserCredential }
  );

  return (
    <div>
      <p>Emails</p>
      <p>Raw Value: {JSON.stringify(rawValue, null, 2)}</p>
      <p>Errors: {JSON.stringify(graphError, null, 2)}</p>
        <div>
      {emails.map((email) => (
                <Email
                key={email.id}
                sender={email.sender}
                subject={email.subject}
                body={email.bodyPreview}
                sentiment={email.sentiment}
              />
        // <div key={email.id}>
        //   <h3>{email.subject}</h3>
        //   <h3>{email.sender.displayName || email.sender.emailAddress.name}</h3>
        //   <p>{email.bodyPreview}</p>
        // </div>
      ))}
    </div>
      {/* <Design /> */}
      <h3>Emails</h3>
      <h3>Example: Get the user's profile</h3>
      <div className="section-margin">
        <p>Click below to authorize button to grant permission to using Microsoft Graph.</p>
        <pre>{`credential.login(scope);`}</pre>
        <Button appearance="primary" disabled={loading} onClick={reload}>
          Authorize
        </Button>

        <p>
          Below are two different implementations of retrieving profile photo for currently
          signed-in user using Fluent UI component and Graph Toolkit respectively.
        </p>
        <h4>1. Display user profile using Fluent UI Component</h4>
        <PersonCardFluentUI loading={loading} data={data} error={error} />
        {/* <h4>2. Display user profile using Graph Toolkit</h4>
        <PersonCardGraphToolkit loading={loading} data={data} error={error} /> */}
      </div>
    </div>
  );
}
