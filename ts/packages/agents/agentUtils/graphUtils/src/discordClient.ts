// DiscordClient.ts

import { Client, Intents, TextChannel } from "discord.js";
import chalk from "chalk";

export class DiscordClient {
    private client: Client;
    private token: string;

    constructor(token: string) {
        this.token = token;
        this.client = new Client({ intents: [Intents.FLAGS.GUILDS, Intents.FLAGS.GUILD_MESSAGES] });
    }

    // Initializes the Discord client and logs it in
    public async initDiscordClient(): Promise<void> {
        await this.client.login(this.token);
        console.log(chalk.green("Discord client authenticated successfully"));
    }

    // Sends a message to a Discord channel
    public async sendMessageAsync(channelId: string, message: string): Promise<boolean> {
        try {
            const channel = await this.client.channels.fetch(channelId) as TextChannel;
            if (channel) {
                await channel.send(message);
                console.log(chalk.green(`Message sent to channel ${channelId}`));
                return true;
            } else {
                console.log(chalk.red(`Channel with ID ${channelId} not found.`));
                return false;
            }
        } catch (error) {
            console.log(chalk.red("Error sending message to Discord:", error));
            return false;
        }
    }

    // Creates a new channel in a guild
    public async createChannelAsync(guildId: string, channelName: string, isPrivate: boolean = false): Promise<boolean> {
        try {
            const guild = await this.client.guilds.fetch(guildId);
            if (!guild) {
                console.log(chalk.red(`Guild with ID ${guildId} not found.`));
                return false;
            }
            await guild.channels.create(channelName, {
                type: "GUILD_TEXT",
                permissionOverwrites: isPrivate ? [{ id: guild.id, deny: ["VIEW_CHANNEL"] }] : [],
            });
            console.log(chalk.green(`Channel ${channelName} created in guild ${guildId}`));
            return true;
        } catch (error) {
            console.log(chalk.red("Error creating channel in Discord:", error));
            return false;
        }
    }
}
