# Teams Agent

Slack agent is **sample code** that explores how to  build a typed Slack agent with **TypeChat**. This package includes the schema definition and implementation needed for a Slack agent that interacts with Slack channels and messaging using the Slack API.  This [article](https://api.slack.com/) explores how to work with typescript and Slack APIs.

This agent depends on the utility library [graph-utils](../graphUtils/src/slackClient.ts) to implement different slack actions.

The teams agent uses the Microsoft Graph API to interact with the user's teams conversations. The agent uses `@microsoft/microsoft-graph-client` library to interact with the Microsoft Graph API. The agent enables operations to create group conversations and send teams messages.

To build the email agent, it needs to provide a manifest and an instantiation entry point.  
These are declared in the `package.json` as export paths:

- `./agent/manifest` - The location of the JSON file for the manifest.
- `./agent/handlers` - an ESM module with an instantiation entry point.

### Prerequisites

This Slack agent uses the Slack API to access your Slack workspace. To get started, create a Slack app and configure it with the necessary permissions, following Slack's API Quickstart. After creating your Slack app and generating a bot token, add the following variables to your `.env` file:

```
SLACK_BOT_TOKEN
```


### Sample User Requests

```
Create a channel for Q4 planning.

Send a message to Q4 planning asking for a status update.
```

## Trademarks

This project may contain trademarks or logos for projects, products, or services. Authorized use of Microsoft
trademarks or logos is subject to and must follow
[Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks/usage/general).
Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion or imply Microsoft sponsorship.
Any use of third-party trademarks or logos are subject to those third-party's policies.
