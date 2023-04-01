# repios-teamsapp

You have a new Teams project scaffolded! To understand more about the structure of the project, you can read the readme files listed below to get further information.

Microsoft Teams apps bring key information, common tools, and trusted processes to where people increasingly gather, learn, and work.Apps are how you extend Teams to fit your needs. Create something brand new for Teams or integrate an existing app.

There are multiple ways to extend Teams, so every app is unique. Some only have one capability, while others have more than one feature to give users various options. For example, your app can display data in a central location, that is, the tab and present that same information through a conversational interface, that is, the bot.

[What is Teams app capabilities](https://aka.ms/teamsfx-capabilities-overview)

## Capabilities scaffolded in this project

- Tab capabilities: [README](./tabs/README.md)
- Bot capabilities: [README](./bot/README.md)

## Requirements

- at least npm v18.14.2
- [Teams Toolkit](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/teams-toolkit-fundamentals?pivots=visual-studio-code) with VSCode, or [teamsfx CLI](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/teamsfx-cli)

## Notes

### `.vscode/`

Includes pre-built:
- [`launch.json`](https://code.visualstudio.com/docs/editor/debugging#_launch-configurations): for various setups to launch the app locally using VSCode's "Run and Debug"
- [`settings.json`](https://code.visualstudio.com/docs/getstarted/settings#_workspace-settings): for shared VSCode workspace settings
- [`tasks.json`](https://code.visualstudio.com/docs/editor/tasks): for developer tasks, somewhat equivalent to Makefile commands

### `templates/`

- `appPackage/`: Includes teams app manifest and Azure AD manifest, set up for substitutions using configs in `.fx/configs/`
- `azure/`: Declares [cloud resources to be provisioned in Azure via Teams Toolkit](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/provision?pivots=visual-studio-code) via [.bicep files](https://learn.microsoft.com/en-us/azure/azure-resource-manager/bicep/overview?tabs=bicep)
