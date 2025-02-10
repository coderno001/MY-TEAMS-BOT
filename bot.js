const { ActivityHandler, MessageFactory } = require('botbuilder');

// Array to store case numbers added by users
const caseNumbers = [];

class MyTeamsBot extends ActivityHandler {
    constructor() {
        super();

        // This event fires when a message is received
        this.onMessage(async (context, next) => {
            const message = context.activity.text.toLowerCase();

            // Check if the user is adding a case number
            if (message.startsWith('add case ')) {
                const newCaseNumber = message.slice(9).trim(); // Extract the case number
                caseNumbers.push(newCaseNumber); // Add the case number to the list
                await context.sendActivity(MessageFactory.text(`Case number ${newCaseNumber} added to the watchlist.`));
            } else {
                // Check if the message contains any case number from the list
                const foundCaseNumbers = caseNumbers.filter(caseNumber => message.includes(caseNumber));
                
                if (foundCaseNumbers.length > 0) {
                    // Notify the user if any case number is found in the message
                    const response = foundCaseNumbers.map(caseNumber => `Case number ${caseNumber} was mentioned.`).join(' ');
                    await context.sendActivity(MessageFactory.text(response));
                } else {
                    // If no case number found, echo the message as a fallback
                    const replyText = `Echo: ${context.activity.text}`;
                    await context.sendActivity(MessageFactory.text(replyText));
                }
            }

            // Continue with the next middleware in the pipeline
            await next();
        });

        // This event fires when a user is added to the conversation
        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            const welcomeText = 'Hello and welcome!';
            for (let cnt = 0; cnt < membersAdded.length; ++cnt) {
                if (membersAdded[cnt].id !== context.activity.recipient.id) {
                    await context.sendActivity(MessageFactory.text(welcomeText));
                }
            }

            // Continue with the next middleware in the pipeline
            await next();
        });
    }
}

module.exports.MyTeamsBot = MyTeamsBot;
