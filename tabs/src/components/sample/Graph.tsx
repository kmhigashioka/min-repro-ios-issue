import "./Graph.css";
import { useGraph, useGraphWithCredential } from "@microsoft/teamsfx-react";
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
  const { teamsUserCredential } = useContext(TeamsFxContext);
  const [error, setError] = useState()
  const [data, setData] = useState("")
  const [info, setInfo] = useState({})

  useEffect(() => {
    async function getInfo() {
      const authToken = await authentication.getAuthToken({
        silent: false,
      })
      const context = await app.getContext()
      setInfo({ authToken, tenantId: context.user?.tenant?.id })
    }

    getInfo()
  }, [])

  const { loading, error: graphError, data: graphData, reload } = useGraph(
    async (graph, teamsUserCredential, scope) => {
      // Call graph api directly to get user profile information
      const profile = await graph.api("/me").get();

      let photoUrl = "";
      try {
        const photo = await graph.api("/me/photo/$value").get();
        photoUrl = URL.createObjectURL(photo);
      } catch {
        // Could not fetch photo from user's profile, return empty string as placeholder.
      }

      return { profile, photoUrl };

      const authToken = await authentication.getAuthToken()

      const context = await app.getContext()

      return { profile, photoUrl, authToken, tenantId: context.user?.tenant?.id };
    },
    { scope: ["User.Read"] }
  );

  // useEffect(() => {
  //   const isNotYetSignedIn = !data && !loading && !error
  //   if (isNotYetSignedIn) {
  //     reload()
  //   }
  // }, [data, error, loading, reload])

  function authorize() {
    const url = new URL(`${window.location.origin}/auth-start.html`);
    url.searchParams.set("clientId", process.env.REACT_APP_CLIENT_ID || "")
    url.searchParams.set("scope", "User.Read email openid profile offline_access")

    authentication.authenticate({
      url: url.toString(),
      width: 600,
      height: 535})
    .then((result) => {
      console.log("Login succeeded: " + result);
      setData(result)
    })
    .catch((reason) => {
      console.log("Login failed: " + reason);
      setError(reason)
    });
  }

  return (
    <div>
      <Design />
      <h3>Example: Get the user's profile</h3>
      <div className="section-margin">
        <p>Click below to authorize button to grant permission to using Microsoft Graph.</p>
        <pre>{`credential.login(scope);`}</pre>
        <Button primary content="Authorize" onClick={authorize} />

        <p>Below are two different implementations of retrieving profile photo for currently signed-in user using Fluent UI component and Graph Toolkit respectively.</p>
        <h4>1. Display user profile using Fluent UI Component</h4>
        {/* <PersonCardFluentUI loading={loading} data={data} error={error} />
        <h4>2. Display user profile using Graph Toolkit</h4>
        <PersonCardGraphToolkit loading={loading} data={data} error={error} /> */}
        <h4>data</h4>
        <pre>{JSON.stringify(data, undefined, 4)}</pre>
        <h4>graphData</h4>
        <pre>{JSON.stringify(graphData, undefined, 4)}</pre>
        <h4>info</h4>
        <pre>{JSON.stringify(info, undefined, 4)}</pre>
        <h4>error</h4>
        <pre>{JSON.stringify(error, undefined, 4)}</pre>
        <h4>graphError</h4>
        <pre>{JSON.stringify(graphError, undefined, 4)}</pre>
      </div>
    </div>
  );
}
