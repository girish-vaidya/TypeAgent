// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

type DynamicObject = { [key: string]: any }; // Flexible object type

export interface SendMessageAction {
    actionName: "sendMessage";
    parameters: {
        recipients: string[];  // Friendly names of recipients
        message: string;
        attachments?: DynamicObject[];  // Optional attachments
    };
}

export type TeamsAction = SendMessageAction; // Add more actions if needed
