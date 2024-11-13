import { createSlackClient, SlackClient } from "graph-utils";
import chalk from "chalk";
import { SlackAction, SendMessageInSlackAction, CreateChannelInSlackAction } from "./slackActionsSchema.js";
import { ActionContext, AppAction, AppAgent, SessionContext } from "@typeagent/agent-sdk";
import { createActionResultFromHtmlDisplay, createActionResultFromError } from "@typeagent/agent-sdk/helpers/action";
import { CommandHandlerNoParams, CommandHandlerTable, getCommandInterface } from "@typeagent/agent-sdk/helpers/command";

// Command handler for Slack login
export class SlackClientLoginCommandHandler implements CommandHandlerNoParams {
    public readonly description = "Log into Slack to access workspace channels and messages";
    public async run(context: ActionContext<SlackActionContext>) {
        const slackClient: SlackClient | undefined = context.sessionContext.agentContext.slackClient;
        if (!slackClient?.isSlackClientInitialized()) {
            await slackClient?.initSlackClient(true);
        }
    }
}

// Command handlers table
const handlers: CommandHandlerTable = {
    description: "Slack login command",
    defaultSubCommand: "login",
    commands: {
        login: new SlackClientLoginCommandHandler(),
    },
};

// Agent instantiation function
export function instantiate(): AppAgent {
    return {
        initializeAgentContext: initializeSlackContext,
        updateAgentContext: updateSlackContext,
        executeAction: executeSlackAction,
        ...getCommandInterface(handlers),
    };
}

// Type for the Slack action context
type SlackActionContext = {
    slackClient: SlackClient | undefined;
};

// Initial context setup
async function initializeSlackContext() {
    return {
        slackClient: undefined,
    };
}

// Context update for enabling/disabling the Slack client
async function updateSlackContext(
    enable: boolean,
    context: SessionContext<SlackActionContext>,
): Promise<void> {
    if (enable) {
        context.agentContext.slackClient = await createSlackClient();
    } else {
        context.agentContext.slackClient = undefined;
    }
}

// Main action execution function
async function executeSlackAction(
    action: AppAction,
    context: ActionContext<SlackActionContext>,
) {
    const { slackClient } = context.sessionContext.agentContext;
    if (!slackClient || !slackClient.isSlackClientInitialized()) {
        return createActionResultFromError("Use @slack login to log into Slack.");
    }

    let result = await handleSlackAction(action as SlackAction, context);
    if (result) {
        return createActionResultFromHtmlDisplay(result);
    }
}

// Handler function for Slack actions
async function handleSlackAction(
    action: SlackAction,
    context: ActionContext<SlackActionContext>,
) {
    const { slackClient } = context.sessionContext.agentContext;
    if (!slackClient) {
        return "<div>Slack client not initialized ...</div>";
    }

    let res;
    switch (action.actionName) {
        case "sendMessageInSlack": {
            const { channelName, message, attachments } = action.parameters as SendMessageInSlackAction["parameters"];
            console.log(chalk.green("Handling sendMessage action ..."));

            // Resolve channel ID using partial matching if necessary
            const resolvedChannelId = await slackClient.findChannelIdByName(channelName);
            if (!resolvedChannelId) {
                return "<div>Error: Channel not found for name: " + channelName + "</div>";
            }

            // Send message to the resolved channel ID
            res = await slackClient.sendMessageAsync(resolvedChannelId, message, attachments ?? []);
            return res
                ? "<div>Message sent successfully in Slack...</div>"
                : "<div>Error encountered when sending message in Slack!</div>";
        }

        case "createChannelInSlack": {
            const { channelName, isPrivate } = action.parameters as CreateChannelInSlackAction["parameters"];
            console.log(chalk.green("Handling createChannel action ..."));

            res = await slackClient.createChannelAsync(channelName, isPrivate);
            return res
                ? "<div>Channel created successfully...</div>"
                : "<div>Error encountered when creating channel!</div>";
        }

        default:
            throw new Error(`Unknown action: ${(action as any).actionName}`);
    }
}
