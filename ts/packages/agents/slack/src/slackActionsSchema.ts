
export interface SendMessageInSlackAction {
    actionName: "sendMessageInSlack";
    parameters: {
        channelName: string; // Friendly name of the channel to send the message to
        message: string;   // The message content
        attachments?: any[]; // Optional attachments for the message
    };
}

export interface CreateChannelInSlackAction {
    actionName: "createChannelInSlack";
    parameters: {
        channelName: string; // Name of the channel to create
        isPrivate?: boolean; // Indicates if the channel is private
    };
}

export type SlackAction = SendMessageInSlackAction | CreateChannelInSlackAction;
