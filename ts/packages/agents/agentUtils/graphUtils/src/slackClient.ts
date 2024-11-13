// slackClient.ts

import registerDebug from "debug";
import { WebClient } from "@slack/web-api";
import chalk from "chalk";

export class SlackClient {
    private readonly logger = registerDebug("typeagent:graphUtils:slackClient");
    private client: WebClient;
    private token: string;

    constructor(token: string) {
        this.token = token;
        this.client = new WebClient(this.token);
    }

    // Checks if the Slack client is authenticated
    public isSlackClientInitialized(): boolean {
        return !!this.token;
    }

    // Initialize the Slack client
    public async initSlackClient(retry: boolean = false): Promise<void> {
        try {
            await this.client.auth.test();
            this.logger(chalk.green("Slack client authenticated successfully"));
        } catch (error) {
            this.logger(chalk.red("Slack client authentication failed:", error));
            if (retry) {
                await this.initSlackClient();
            }
        }
    }

    // Find channel ID by friendly name with partial matching
    public async findChannelIdByName(friendlyName: string): Promise<string | null> {
        try {
            const validChannelName = this.sanitizeChannelName(friendlyName);
            if (!validChannelName) {
                this.logger(chalk.red(`Channel name "${friendlyName}" is invalid and cannot be created.`));
                return null;
            }
            const response = await this.client.conversations.list();

            if (response.ok && response.channels) {
                // Use partial matching by checking if friendlyName is a substring of the channel's name
                const channel = response.channels.find(
                    (ch) => ch.name?.includes(validChannelName) || ch.name_normalized?.includes(validChannelName)
                );

                if (channel && channel.id) {
                    this.logger(chalk.green(`Found channel ID ${channel.id} for ${validChannelName}`));
                    return channel.id;
                }
            }
            this.logger(chalk.red(`No channel found containing the name: ${validChannelName}`));
            return null;
        } catch (error) {
            this.logger(chalk.red("Error finding channel by name:", error));
            return null;
        }
    }

    // Sends a message to a Slack channel
    public async sendMessageAsync(channelId: string, text: string, attachments?: any[]): Promise<boolean> {
        try {
            const response = await this.client.chat.postMessage({
                channel: channelId,
                text,
                attachments,
            });

            if (response.ok) {
                this.logger(chalk.green(`Message sent to channel ${channelId}`));
                return true;
            } else {
                this.logger(chalk.red(`Failed to send message to channel ${channelId}:`, response.error || 'Unknown error'));
                return false;
            }
        } catch (error) {
            this.logger(chalk.red("Error sending message:", error));
            return false;
        }
    }

    // Creates a new channel in Slack
    public async createChannelAsync(channelName: string, isPrivate: boolean = false): Promise<boolean> {
        // Sanitize and validate the channel name
        const validChannelName = this.sanitizeChannelName(channelName);
        if (!validChannelName) {
            this.logger(chalk.red(`Channel name "${channelName}" is invalid and cannot be created.`));
            return false;
        }

        try {
            const response = await this.client.conversations.create({
                name: validChannelName,
                is_private: isPrivate,
            });

            if (response.ok) {
                this.logger(chalk.green(`Channel ${validChannelName} created successfully`));
                return true;
            } else {
                this.logger(chalk.red(`Failed to create channel ${validChannelName}:`, response.error || 'Unknown error'));
                return false;
            }
        } catch (error) {
            this.logger(chalk.red("Error creating channel:", error));
            return false;
        }
    }

    // Helper function to sanitize and validate channel names
    private sanitizeChannelName(name: string): string | null {
        // Remove invalid characters and convert to lowercase
        const sanitized = name.toLowerCase().replace(/[^a-z0-9-_]/g, '');
        
        // Ensure the name meets Slack’s requirements
        if (sanitized.length < 1 || sanitized.length > 80) {
            this.logger(chalk.red(`Channel name "${name}" is invalid. It must be between 1 and 80 characters.`));
            return null;
        }
        
        return sanitized;
    }
}

// Properly export the createSlackClient function
export async function createSlackClient(): Promise<SlackClient> {
    return new SlackClient(process.env.SLACK_BOT_TOKEN || "");
}
