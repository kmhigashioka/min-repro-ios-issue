import "./Graph.css";
import { useData, useGraphWithCredential } from "@microsoft/teamsfx-react";
import { Providers, ProviderState } from '@microsoft/mgt-element';
import { TeamsFxProvider } from '@microsoft/mgt-teamsfx-provider';
import { authentication, app } from "@microsoft/teams-js";
import { Button } from "@fluentui/react-northstar";
import { Design } from './Design';
import { PersonCardFluentUI } from './PersonCardFluentUI';
import { PersonCardGraphToolkit } from './PersonCardGraphToolkit';
import { useContext, useEffect, useState } from "react";
import { TeamsFxContext } from "../Context";

export function Graph() {
  const [info, setInfo] = useState()
  const { teamsUserCredential } = useContext(TeamsFxContext);
  // const { loading, error, data, reload } = useGraphWithCredential(
  //   async (graph, teamsUserCredential, scope) => {
  //     // Call graph api directly to get user profile information
  //     const profile = await graph.api("/me").get();

  //     // Initialize Graph Toolkit TeamsFx provider
  //     const provider = new TeamsFxProvider(teamsUserCredential, scope);
  //     Providers.globalProvider = provider;
  //     Providers.globalProvider.setState(ProviderState.SignedIn);

  //     let photoUrl = "";
  //     try {
  //       const photo = await graph.api("/me/photo/$value").get();
  //       photoUrl = URL.createObjectURL(photo);
  //     } catch {
  //       // Could not fetch photo from user's profile, return empty string as placeholder.
  //     }

  //     const authToken = await authentication.getAuthToken()
  //     const context = await app.getContext()

  //     return { profile, photoUrl, authToken, context };
  //   },
  //   { scope: ["User.Read"], credential: teamsUserCredential }
  // );

  const { loading, data, error } = useData(async () => {
    const userInfo = await teamsUserCredential!.getUserInfo();
    const authToken = await authentication.getAuthToken()
    const tid = userInfo.tenantId
    const token = authToken
    // @ts-ignore
    setInfo({ tid, token })
    fetch(`/api/me?tid=${tid}&token=${token}`, {
      headers: {
        'Content-Type': 'application/json'
      },
      mode: 'cors',
      cache: 'default'
    })
      .then((response) => {
        if (response.ok) {
          return response.json();
        } else {
          // @ts-ignore
          throw response.error
        }
      })
      .then((responseJson) => {
        if (responseJson.error) {
          throw responseJson.error;
        } else {
          const profile = responseJson;
          return profile
        }
      });
  });

  return (
    <div>
      <Design />
      <h3>Example: Get the user's profile</h3>
      <div className="section-margin">
        <p>Click below to authorize button to grant permission to using Microsoft Graph.</p>
        <pre>{`credential.login(scope);`}</pre>
        {/* <Button primary content="Authorize" disabled={loading} onClick={reload} /> */}

        <p>Below are two different implementations of retrieving profile photo for currently signed-in user using Fluent UI component and Graph Toolkit respectively.</p>
        <h4>1. Display user profile using Fluent UI Component</h4>
        {/* @ts-ignore */}
        <PersonCardFluentUI loading={loading} data={data} error={error} />
        <h4>2. Display user profile using Graph Toolkit</h4>
        {/* @ts-ignore */}
        <PersonCardGraphToolkit loading={loading} data={data} error={error} />
        <pre>{JSON.stringify(data, undefined, 4)}</pre>
        <h4>info</h4>
        <pre>{JSON.stringify(info, undefined, 4)}</pre>
      </div>
    </div>
  );
}
