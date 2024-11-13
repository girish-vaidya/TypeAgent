// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { createTeamsGraphClient, TeamsClient } from "graph-utils";
import { TeamsAction, SendMessageAction } from "./teamsActionsSchema.js";
import { ActionContext, AppAction, AppAgent, SessionContext } from "@typeagent/agent-sdk";
import { createActionResultFromError, createActionResultFromHtmlDisplay } from "@typeagent/agent-sdk/helpers/action";
import { CommandHandlerNoParams, CommandHandlerTable, getCommandInterface } from "@typeagent/agent-sdk/helpers/command";

// Command handler for Teams login
export class TeamsClientLoginCommandHandler implements CommandHandlerNoParams {
    public readonly description = "Log into MS Graph to access Teams";
    public async run(context: ActionContext<TeamsActionContext>) {
        const teamsClient: TeamsClient | undefined = context.sessionContext.agentContext.teamsClient;
        if (!teamsClient?.isGraphClientInitialized()) {
            await teamsClient?.initGraphClient(true);
        }
    }
}

// Command handlers table
const handlers: CommandHandlerTable = {
    description: "Teams login command",
    defaultSubCommand: "login",
    commands: {
        login: new TeamsClientLoginCommandHandler(),
    },
};

// Agent instantiation function
export function instantiate(): AppAgent {
    return {
        initializeAgentContext: initializeTeamsContext,
        updateAgentContext: updateTeamsContext,
        executeAction: executeTeamsAction,
        ...getCommandInterface(handlers),
    };
}

// Type for the Teams action context
type TeamsActionContext = {
    teamsClient: TeamsClient | undefined;
    callerId: string | undefined; // Add callerId to hold the authenticated user's ID
};

// Initial context setup
async function initializeTeamsContext(): Promise<TeamsActionContext> {
    const teamsClient = await createTeamsGraphClient();
    return {
        teamsClient: teamsClient,
        callerId: undefined,
    };
}

// Context update for enabling/disabling the Teams client
async function updateTeamsContext(
    enable: boolean,
    context: SessionContext<TeamsActionContext>,
): Promise<void> {
    if (enable) {
        context.agentContext.teamsClient = await createTeamsGraphClient();
    } else {
        context.agentContext.teamsClient = undefined;
    }
}

// Main action execution function
async function executeTeamsAction(
    action: AppAction,
    context: ActionContext<TeamsActionContext>,
) {
    const { teamsClient } = context.sessionContext.agentContext;
    if (!teamsClient || !teamsClient?.isGraphClientInitialized()) {
        return createActionResultFromError("Use @teams login to log into MS Graph.");
    }

    // Initialize callerId if not already initialized
    if (!context.sessionContext.agentContext.callerId) {
        const callerId = await teamsClient.getCallerIdAsync();
        
        if (!callerId) {
            return createActionResultFromError("Unable to retrieve caller ID.");
        }

        // Update context with the retrieved callerId
        context.sessionContext.agentContext.callerId = callerId;
    }

    // Directly return the result from handleActions to avoid wrapping in another ActionResult
    return await handleActions(action as TeamsAction, context);
}

// Handler function for Teams actions
async function handleActions(
    action: TeamsAction,
    context: ActionContext<{ teamsClient: TeamsClient | undefined; callerId: string | undefined; }>
) {
    const { teamsClient, callerId } = context.sessionContext.agentContext;

    if (!teamsClient || !callerId) {
        return createActionResultFromError("Teams client or caller ID not initialized.");
    }

    const { recipients, message, attachments } = action.parameters as SendMessageAction["parameters"];

    // Find user IDs based on recipient names
    const userIds = await teamsClient.findUserIdsByNamesAsync(recipients);

    if (userIds.length === 0) {
        return createActionResultFromError("No valid users found for conversation.");
    }

    // Determine if it's a group or individual chat
    const chatId = userIds.length > 1
        ? await teamsClient.createOrFindConversationAsync(callerId, userIds)  // Pass callerId and userIds for group chat
        : await teamsClient.createOrFind1on1ChatAsync(callerId, userIds[0]);  // Individual chat

    if (!chatId) {
        return createActionResultFromError("Failed to create or retrieve chat with specified users.");
    }

    // Send the message with any attachments
    const result = await teamsClient.sendMessageAsync([chatId], message, attachments ?? []);
    return result
        ? createActionResultFromHtmlDisplay("<div>Message sent successfully in Teams.</div>")
        : createActionResultFromError("Error sending message in Teams.");
}

