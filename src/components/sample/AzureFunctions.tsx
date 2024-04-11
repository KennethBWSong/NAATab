import { useContext, useState } from "react";
import { Button, Spinner } from "@fluentui/react-components";
import { useData } from "@microsoft/teamsfx-react";
import * as axios from "axios";
import { TeamsFxContext } from "../Context";
import config from "./lib/config";
import { PublicClientApplication } from "@azure/msal-browser";
import { app, authentication } from "@microsoft/teams-js";

const functionName = config.apiName || "myFunc";

export function AzureFunctions(props: { codePath?: string; docsUrl?: string }) {
  const [needConsent, setNeedConsent] = useState(false);
  const { codePath, docsUrl } = {
    codePath: `api/src/functions/${functionName}.ts`,
    docsUrl: "https://aka.ms/teamsfx-azure-functions",
    ...props,
  };
  const teamsUserCredential = useContext(TeamsFxContext).teamsUserCredential;
  const { loading, data, error, reload } = useData(async () => {
    await app.initialize();
    const msalConfig = {
      auth: {
          clientId: `${config.clientId}`,
          authority: "https://login.microsoftonline.com/72f988bf-86f1-41af-91ab-2d7cd011db47",
          supportsNestedAppAuth: true, // Enable native bridging.
      },
      cache: {
        cacheLocation: "localStorage",
      },
    };
    const msalClient = await PublicClientApplication.createPublicClientApplication(msalConfig);
    if (!teamsUserCredential) {
      throw new Error("TeamsFx SDK is not initialized.");
    }
    if (needConsent) {
      // await teamsUserCredential!.login(["User.Read"]);
      const res = await authentication.authenticate({
        url: `${config.initiateLoginEndpoint}?clientId=${config.clientId}&scope=${encodeURI("User.Read")}&loginHint=bowsong@microsoft.com`,
        width: 600,
        height: 535, 
      });
      console.log("-------------------------call login success-------------------------");
      const account = JSON.parse(res).account
      msalClient.setActiveAccount(account);
      setNeedConsent(false);
    }
    try {
      const account = msalClient.getActiveAccount();
      const request = {
        scopes: ["User.Read"],
        account: account ?? undefined
      };
      const res = await msalClient.acquireTokenSilent(request);
      console.log("-------------------------call silentAuth success-------------------------");
      console.log(JSON.stringify(res));
    } catch (error: any) {
      console.log("-------------------------call silentAuth failed-------------------------");
      console.log(JSON.stringify(error));
      setNeedConsent(true);
    }
  });
  return (
    <div>
      <h2>Call your Azure Functions</h2>
      <p>
        An Azure Functions app is running. Authorize this app and click below to call it for a
        response:
      </p>
      {!loading && (
        <Button appearance="primary" disabled={loading} onClick={reload}>
          Authorize and call Azure Functions
        </Button>
      )}
      {loading && (
        <pre className="fixed">
          <Spinner />
        </pre>
      )}
      {!loading && !!data && !error && <pre className="fixed">{JSON.stringify(data, null, 2)}</pre>}
      {!loading && !data && !error && <pre className="fixed"></pre>}
      {!loading && !!error && <div className="error fixed">{(error as any).toString()}</div>}
      <h4>How to edit the Azure Functions</h4>
      <p>
        See the code in <code>{codePath}</code> to add your business logic.
      </p>
      {!!docsUrl && (
        <p>
          For more information, see the{" "}
          <a href={docsUrl} target="_blank" rel="noreferrer">
            docs
          </a>
          .
        </p>
      )}
    </div>
  );
}
