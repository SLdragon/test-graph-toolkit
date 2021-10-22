import React, { useState } from "react";
import { Image, Menu } from "@fluentui/react-northstar";
import "./Welcome.css";
import { EditCode } from "./EditCode";
import { AzureFunctions } from "./AzureFunctions";
import { Graph } from "./Graph";
import { CurrentUser } from "./CurrentUser";
import { useTeamsFx } from "./lib/useTeamsFx";
import { TeamsUserCredential } from "@microsoft/teamsfx";
import { useData } from "./lib/useData";
import { Deploy } from "./Deploy";
import { Publish } from "./Publish";
import {Providers, SimpleProvider, ProviderState } from "@microsoft/mgt";


export function Welcome(props) {
  const { showFunction, environment } = {
    showFunction: true,
    environment: window.location.hostname === "localhost" ? "local" : "azure",
    ...props,
  };
  const friendlyEnvironmentName =
    {
      local: "local environment",
      azure: "Azure environment",
    }[environment] || "local environment";

  const steps = ["local", "azure", "publish"];
  const friendlyStepsName = {
    local: "1. Build your app locally",
    azure: "2. Provision and Deploy to the Cloud",
    publish: "3. Publish to Teams",
  };
  const [selectedMenuItem, setSelectedMenuItem] = useState("local");
  const items = steps.map((step) => {
    return {
      key: step,
      content: friendlyStepsName[step] || "",
      onClick: () => setSelectedMenuItem(step),
    };
  });

  const { isInTeams } = useTeamsFx();
  const userProfile = useData(async () => {
    const credential = new TeamsUserCredential();
    initGraphToolkit(credential, ["User.Read"]);
    return isInTeams ? await credential.getUserInfo() : undefined;
  })?.data;
  const userName = userProfile ? userProfile.displayName : "";
  async function initGraphToolkit(credential, scope) {
    async function getAccessToken(scopes) {
      let tokenObj = await credential.getToken(scopes);
      return tokenObj.token;
    }

    async function login() {
      try {
       await credential.login(scope);
      } catch (err) {
        alert("Login failed: " + err);
         return;
       }

      Providers.globalProvider.setState(ProviderState.SignedIn)
    }

    function logout() {}

    Providers.globalProvider = new SimpleProvider(getAccessToken, login, logout);
    Providers.globalProvider.setState(ProviderState.SignedOut)
  }

  return (
    <div className="welcome page">
      <div className="narrow page-padding">
        <Image src="hello.png" />
        <h1 className="center">Congratulations{userName ? ", " + userName : ""}!</h1>
        <p className="center">Your app is running in your {friendlyEnvironmentName}</p>
        <Menu defaultActiveIndex={0} items={items} underlined secondary />
        <mgt-login></mgt-login>
        <div className="sections">
          {selectedMenuItem === "local" && (
            <div>
              <EditCode showFunction={showFunction} />
              {isInTeams && <CurrentUser userName={userName} />}
              <Graph />
              {showFunction && <AzureFunctions />}
            </div>
          )}
          {selectedMenuItem === "azure" && (
            <div>
              <Deploy />
            </div>
          )}
          {selectedMenuItem === "publish" && (
            <div>
              <Publish />
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
