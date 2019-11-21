// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import {
    ActivityHandler,
    ActionTypes,
    ActivityTypes,
    CardFactory,
    ConversationState,
    UserState,
} from 'botbuilder';

import { Culture, recognizeNumber, recognizeDateTime } from '@microsoft/recognizers-text-suite';

const CONVERSATION_FLOW_PROPERTY = 'conversation_prop';
const USER_PROFILE_PROPERTY = 'user_prop';

const question = {
    none: 'NONE',
    age: 'AGE',
    name: 'NAME',
    school: 'SCHOOL'
};

const action = {
    none: 'NONE',
    rsvp: 'RSVP',
    cancel: 'Cancel RSVP',
    friend: 'Invite a friend'
};

const friendQuestion = {
    none: 'NONE',
    name: 'NAME',
    email: 'EMAIL'
};

const cancelQuestion = {
    none: 'NONE',
    name: 'NAME'
};

export class EchoBot extends ActivityHandler {
    // Validates name input. Returns whether validation succeeded and either the parsed and normalized
// value or a message the bot can use to ask the user again.
    private static validateString(input) {
        const string = input && input.trim();
        return string !== undefined
            ? { success: true, string: string }
            : { success: false, message: 'Please enter a name that contains at least one character.' };
    };

// Validates age input. Returns whether validation succeeded and either the parsed and normalized
// value or a message the bot can use to ask the user again.
    private static validateAge(input) {
        // Try to recognize the input as a number. This works for responses such as "twelve" as well as "12".
        try {
            // Attempt to convert the Recognizer result to an integer. This works for "a dozen", "twelve", "12", and so on.
            // The recognizer returns a list of potential recognition results, if any.
            const results = recognizeNumber(input, Culture.English);
            let output;
            results.forEach(result => {
                // result.resolution is a dictionary, where the "value" entry contains the processed string.
                const value = result.resolution.value;
                if (value) {
                    const age = parseInt(value);
                    if (!isNaN(age) && age >= 18 && age <= 120) {
                        output = { success: true, age: age };
                        return;
                    }
                }
            });
            return output || { success: false, message: 'Please enter an age between 18 and 120.' };
        } catch (error) {
            return {
                success: false,
                message: "I'm sorry, I could not interpret that as an age. Please enter an age between 18 and 120."
            };
        }
    }

    private static async sendHeroCard(context) {
        const reply: any = { type: ActivityTypes.Message };
        const buttons = [
            { type: ActionTypes.ImBack, title: '1. RSVP', value: action.rsvp },
            { type: ActionTypes.ImBack, title: '2. Cancel RSVP', value: action.cancel },
            { type: ActionTypes.ImBack, title: '3. Invite a friend', value: action.friend }
        ];

        const card = CardFactory.heroCard('', undefined,
            buttons, { text: 'What do you want to do ?' });

        reply.attachments = [card];

        await context.sendActivity(reply);
    }

    private static async fillOutRSVP(flow, profile, turnContext) {
        const input = turnContext.activity.text;
        let result;
        switch (flow.lastQuestionAsked) {
            // If we're just starting off, we haven't asked the user for any information yet.
            // Ask the user for their name and update the conversation flag.
            case question.none:
                await turnContext.sendActivity('In order to RSVP you need to give out some information. What is your name?');
                flow.lastQuestionAsked = question.name;
                break;

            // If we last asked for their name, record their response, confirm that we got it.
            // Ask them for their age and update the conversation flag.
            case question.name:
                result = EchoBot.validateString(input);
                if (result.success) {
                    profile.name = result.string;
                    await turnContext.sendActivity(`Welcome among us ${ profile.name }.`);
                    await turnContext.sendActivity('How old are you?');
                    flow.lastQuestionAsked = question.age;
                    break;
                } else {
                    // If we couldn't interpret their input, ask them for it again.
                    // Don't update the conversation flag, so that we repeat this step.
                    await turnContext.sendActivity(result.message || "I'm sorry, I didn't understand that.");
                    break;
                }

            // If we last asked for their age, record their response, confirm that we got it.
            // Ask them for their date preference and update the conversation flag.
            case question.age:
                result = EchoBot.validateAge(input);
                if (result.success) {
                    profile.age = result.age;
                    await turnContext.sendActivity(`I have your age as ${ profile.age }.`);
                    await turnContext.sendActivity('What university are you from?');
                    flow.lastQuestionAsked = question.school;
                    break;
                } else {
                    // If we couldn't interpret their input, ask them for it again.
                    // Don't update the conversation flag, so that we repeat this step.
                    await turnContext.sendActivity(result.message || "I'm sorry, I didn't understand that.");
                    break;
                }

            // If we last asked for a date, record their response, confirm that we got it,
            // let them know the process is complete, and update the conversation flag.
            case question.school:
                result = EchoBot.validateString(input);
                if (result.success) {
                    profile.school = result.string;
                    if (profile.school === 'Stanford') {
                        await turnContext.sendActivity('I\'m sorry, students from Stanford are not accepted. Nobody\'s perfect but it is still time to apply to Berkeley next year following this link: https://grad.berkeley.edu/admissions/apply/');
                    } else {
                        await turnContext.sendActivity(`Amazing, so many people coming from ${profile.school}.`);
                        const reply: any = { type: ActivityTypes.Message };
                        reply.text = 'You can already take a look at the documentation here: https://docs.microsoft.com/en-us/azure/bot-service/bot-service-debug-emulator?view=azure-bot-service-4.0';
                        reply.attachments = [EchoBot.getInternetAttachment()];
                        await turnContext.sendActivity(reply);
                        await turnContext.sendActivity(`Thanks for completing the RSVP ${profile.name}.`);
                    }
                    flow.lastQuestionAsked = question.none;
                    flow.action = action.none;
                    profile = {};
                    EchoBot.sendHeroCard(turnContext);
                    break;
                } else {
                    // If we couldn't interpret their input, ask them for it again.
                    // Don't update the conversation flag, so that we repeat this step.
                    await turnContext.sendActivity(result.message || "I'm sorry, I didn't understand that.");
                    break;
                }
        }
    }

    private static async fillOutFriend(flow, profile, turnContext) {
        const input = turnContext.activity.text;
        let result;
        switch (flow.lastQuestionFriend) {
            // If we're just starting off, we haven't asked the user for any information yet.
            // Ask the user for their name and update the conversation flag.
            case friendQuestion.none:
                await turnContext.sendActivity('What is the name of your friend?');
                flow.lastQuestionFriend = friendQuestion.name;
                break;

            case friendQuestion.name:
                result = EchoBot.validateString(input);
                if (result.success) {
                    profile.firenName = result.string;
                    await turnContext.sendActivity('Great! What is your email address?');
                    flow.lastQuestionFriend = friendQuestion.email;
                    break;
                } else {
                    // If we couldn't interpret their input, ask them for it again.
                    // Don't update the conversation flag, so that we repeat this step.
                    await turnContext.sendActivity(result.message || "I'm sorry, I didn't understand that.");
                    break;
                }
            case friendQuestion.email:
                result = EchoBot.validateString(input);
                if (result.success) {
                    profile.email = result.string;
                    await turnContext.sendActivity('Great, we\'ll have a lot of fun!');
                    await turnContext.sendActivity('An email to v-dalhay@microsoft.com has been sent. You\'ll receive a response by email within 24 hours.');
                    flow.lastQuestionFriend = question.none;
                    flow.action = action.none;
                    profile = {};
                    EchoBot.sendHeroCard(turnContext);
                    break;
                } else {
                    // If we couldn't interpret their input, ask them for it again.
                    // Don't update the conversation flag, so that we repeat this step.
                    await turnContext.sendActivity(result.message || "I'm sorry, I didn't understand that.");
                    break;
                }

        }
    }

    private static async cancelRSVP(flow, profile, turnContext) {
        const input = turnContext.activity.text;
        let result;
        switch (flow.lastQuestionCancel) {
            // If we're just starting off, we haven't asked the user for any information yet.
            // Ask the user for their name and update the conversation flag.
            case cancelQuestion.none:
                await turnContext.sendActivity('What is your name to cancel your participation?');
                flow.lastQuestionCancel = cancelQuestion.name;
                break;

            case cancelQuestion.name:
                result = EchoBot.validateString(input);
                if (result.success) {
                    profile.cancelName = result.string;
                    await turnContext.sendActivity('Ok, let\'s cancel!');
                    flow.lastQuestionCancel = question.none;
                    flow.action = action.none;
                    profile = {};
                    EchoBot.sendHeroCard(turnContext);
                    break;
                } else {
                    // If we couldn't interpret their input, ask them for it again.
                    // Don't update the conversation flag, so that we repeat this step.
                    await turnContext.sendActivity(result.message || "I'm sorry, I didn't understand that.");
                    break;
                }
        }
    }

    private static getInternetAttachment() {
        // NOTE: The contentUrl must be HTTPS.
        return {
            name: 'architecture-resize.png',
            contentType: 'image/png',
            contentUrl: 'https://docs.microsoft.com/en-us/bot-framework/media/how-it-works/architecture-resize.png'
        };
    }

    private userState: UserState;
    private conversationState: ConversationState;
    private conversationFlow: any;
    private userProfile: any;

    constructor(conversationState: ConversationState, userState: UserState) {
        super();

        // The state property accessors for conversation flow and user profile.
        this.conversationFlow = conversationState.createProperty(CONVERSATION_FLOW_PROPERTY);
        this.userProfile = userState.createProperty(USER_PROFILE_PROPERTY);

        // The state management objects for the conversation and user.
        this.conversationState = conversationState;
        this.userState = userState;

        this.onDialog(async (context, next) => {
            // Save any state changes. The load happened during the execution of the Dialog.
            await this.conversationState.saveChanges(context, false);
            await this.userState.saveChanges(context, false);

            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });

        this.onMessage(async (context, next) => {
            const flow = await this.conversationFlow.get(context, {
                lastQuestionAsked: question.none,
                lastQuestionFriend: friendQuestion.none,
                lastQuestionCancel: cancelQuestion.none,
                action: action.none
            });
            const profile = await this.userProfile.get(context, {});

            const actions = new Set([action.rsvp, action.cancel, action.friend]);
            console.log('Text:', context.activity.text);
            console.log('Action:', flow.action);
            console.log('Last RSVP question:', flow.lastQuestionAsked);
            console.log('Last friend question:', flow.lastQuestionFriend);
            if (actions.has(context.activity.text)) {
                if (flow.action === action.none) {
                    flow.action = context.activity.text;
                } else {
                    context.sendActivity('You need to finish the current action before doing another one!');
                    await next();
                    return;
                }
            }
            switch (flow.action) {
                case action.rsvp:
                    await EchoBot.fillOutRSVP(flow, profile, context);
                    break;
                case action.cancel:
                    await EchoBot.cancelRSVP(flow, profile, context);
                    break;
                case action.friend:
                    await EchoBot.fillOutFriend(flow, profile, context);
                    break;
                default:
                    console.log('Default case');
            }

            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });

        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            for (const member of membersAdded) {
                if (member.id !== context.activity.recipient.id) {
                    await context.sendActivity('Welcome at the 2020 Microsoft AI Hackathon!');
                    EchoBot.sendHeroCard(context);
                }
            }
            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });
    }
}
