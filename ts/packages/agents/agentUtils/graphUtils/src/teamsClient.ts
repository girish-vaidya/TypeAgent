// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { GraphClient, DynamicObject } from "./graphClient.js";
import chalk from "chalk";
import registerDebug from "debug";
import fs from "fs";
import path from "path";


export class TeamsClient {
    private readonly logger = registerDebug("typeagent:graphUtils:teamsClient");
    private graphClient: GraphClient | undefined = undefined;

    public constructor() {}

    public isGraphClientInitialized(): boolean {
        return this.graphClient && this.graphClient.getClient() ? true : false;
    }

    public async initGraphClient(fLogin: boolean): Promise<void> {
        if (this.graphClient === undefined) {
            this.graphClient = await GraphClient.getInstance();
            this.graphClient?.authenticateUser();
        } else if (fLogin) {
            await this.graphClient.ensureTokenIsValid();
        }
        return;
    }

    // Retrieve the caller's user ID
    public async getCallerIdAsync(): Promise<string | undefined> {
        if (this.graphClient === undefined) return undefined;

        this.graphClient.ensureTokenIsValid();
        try {
            const response = await this.graphClient.getClient()?.api("/me").select("id").get();
            return response?.id;
        } catch (error) {
            this.logger(chalk.red("Error retrieving caller ID:", error));
            return undefined;
        }
    }

    // Find user IDs based on friendly or given names
    public async findUserIdsByNamesAsync(names: string[]): Promise<string[]> {
        if (this.graphClient === undefined) return [];

        this.graphClient.ensureTokenIsValid();
        const userIds: string[] = [];

        for (const name of names) {
            try {
                const users = await this.graphClient
                    .getClient()
                    ?.api("/users")
                    .filter(`startswith(displayName, '${name}') or startswith(givenName, '${name}')`)
                    .select("id")
                    .top(1)
                    .get();

                if (users && users.value && users.value.length > 0) {
                    userIds.push(users.value[0].id);
                } else {
                    this.logger(chalk.red(`No user found with name: ${name}`));
                }
            } catch (error) {
                this.logger(chalk.red(`Error finding user by name (${name}):`, error));
            }
        }

        return userIds;
    }

    // Create or find a conversation with a list of users (group chat)
    public async createOrFindConversationAsync(callerId: string, userIds: string[]): Promise<string | undefined> {
        if (this.graphClient === undefined) return undefined;
    
        this.graphClient.ensureTokenIsValid();
        
        // Include the caller ID in the group membership
        const allUserIds = [callerId, ...userIds];
    
        try {
            // Step 1: Check for an existing group chat with the specified users, including the caller
            const existingChats = await this.graphClient
                .getClient()
                ?.api("/me/chats")
                .filter(`chatType eq 'group'`)
                .expand("members")
                .get();
    
            for (const chat of existingChats.value) {
                const memberIds = chat.members.map((member: any) => member.userId);
                if (allUserIds.every(id => memberIds.includes(id))) {
                    this.logger(chalk.green("Found existing group conversation with specified users."));
                    return chat.id;
                }
            }
    
            // Step 2: No existing chat found, so create a new group chat
            const chatPayload = {
                chatType: "group",
                members: allUserIds.map(userId => ({
                    '@odata.type': '#microsoft.graph.aadUserConversationMember',
                    roles: ["owner"],
                    'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${userId}')`,
                })),
                topic: "Group Conversation",
            };
    
            const newChat = await this.graphClient
                .getClient()
                ?.api("/chats")
                .post(chatPayload);
    
            if (newChat?.id) {
                this.logger(chalk.green("Created new group conversation with specified users."));
                return newChat.id;
            }
        } catch (error) {
            this.logger(chalk.red("Error creating or finding group conversation:", error));
        }
    
        return undefined;
    }
    

    // Create a 1-on-1 chat if it doesn't exist
    public async createOrFind1on1ChatAsync(callerId: string, otherUserId: string): Promise<string | undefined> {
        if (this.graphClient === undefined) return undefined;

        this.graphClient.ensureTokenIsValid();
        try {
            // Step 1: Attempt to find an existing 1-on-1 chat with the other user
            const existingChats = await this.graphClient
                .getClient()
                ?.api("/me/chats")
                .filter(`chatType eq 'oneOnOne'`)
                .expand("members")
                .get();

            // Check if any existing chat contains both the caller and the other user
            for (const chat of existingChats.value) {
                const memberIds = chat.members.map((member: any) => member.userId);
                if (memberIds.includes(callerId) && memberIds.includes(otherUserId)) {
                    this.logger(chalk.green(`Found existing 1-on-1 chat with user ID ${otherUserId}.`));
                    return chat.id;
                }
            }

            // Step 2: No existing chat found, so create a new 1-on-1 chat
            const chatPayload = {
                chatType: "oneOnOne",
                members: [
                    {
                        '@odata.type': '#microsoft.graph.aadUserConversationMember',
                        roles: ["owner"],
                        'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${callerId}')`,
                    },
                    {
                        '@odata.type': '#microsoft.graph.aadUserConversationMember',
                        roles: ["owner"],
                        'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${otherUserId}')`,
                    }
                ],
            };

            const chat = await this.graphClient
                .getClient()
                ?.api("/chats")
                .post(chatPayload);

            if (chat?.id) {
                this.logger(chalk.green(`1-on-1 chat created with user ID ${otherUserId}`));
                return chat.id;
            }
        } catch (error) {
            this.logger(chalk.red("Error creating 1-on-1 chat:", error));
        }
        return undefined;
    }

    // Send a message to multiple people or an existing conversation

    public async sendMessageAsync(
        chatIds: string[],
        content: string,
        attachments: DynamicObject[] = []
    ): Promise<boolean> {
        if (this.graphClient === undefined) return false;

        this.graphClient.ensureTokenIsValid();
        let allSent = true;

        // Prepare the attachments for the API
        const preparedAttachments = await Promise.all(
            attachments.map(async (attachment) => {
                if (attachment.filePath) {
                    try {
                        // Read file content and encode it as base64
                        const fileContent = await fs.promises.readFile(path.resolve(attachment.filePath));
                        const base64FileContent = fileContent.toString("base64");

                        // Set up attachment structure for Microsoft Graph API
                        return {
                            "@odata.type": "#microsoft.graph.fileAttachment",
                            contentBytes: base64FileContent,
                            contentType: attachment.contentType || "application/octet-stream",
                            name: attachment.fileName || path.basename(attachment.filePath),
                        };
                    } catch (error) {
                        this.logger(chalk.red(`Error reading file ${attachment.filePath}: ${error}`));
                        return null;
                    }
                } else {
                    this.logger(chalk.red("Attachment missing filePath."));
                    return null;
                }
            })
        );

        // Filter out any attachments that failed to load
        const validAttachments = preparedAttachments.filter((attachment) => attachment !== null);

        // Send message with attachments to each chat
        try {
            for (const chatId of chatIds) {
                const messagePayload = {
                    body: {
                        content: content,
                    },
                    attachments: validAttachments,
                };

                await this.graphClient
                    .getClient()
                    ?.api(`/chats/${chatId}/messages`)
                    .post(messagePayload)
                    .then(() => {
                        this.logger(chalk.green(`Message sent successfully to chat ${chatId}`));
                    })
                    .catch((error) => {
                        this.logger(chalk.red(`Error sending message to chat ${chatId}: ${error}`));
                        allSent = false;
                    });
            }
        } catch (error) {
            this.logger(chalk.red("Error sending message:", error));
            allSent = false;
        }

        return allSent;
    }
}

// Function to initialize Teams client
export async function createTeamsGraphClient(): Promise<TeamsClient> {
    return new TeamsClient();
}
