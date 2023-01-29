import {
  TeamsActivityHandler,
  CardFactory,
  TurnContext,
  AdaptiveCardInvokeValue,
  AdaptiveCardInvokeResponse,
  TeamsInfo,
  TeamsMeetingInfo,
} from "botbuilder";
import rawWelcomeCard from "./adaptiveCards/welcome.json";
import { AdaptiveCards } from "@microsoft/adaptivecards-tools";
import { Configuration, OpenAIApi } from "openai";
import config from "./config";
import GraphHelper from "./helpers/graphHelper.js"

export class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();

    const configuration = new Configuration({
      apiKey: config.openaiApiKey,
      organization: config.organization
    });


     // Activity handler for meeting end event.
    this.onTeamsMeetingEndEvent(async (meeting, context, next) => {
      const meetingDetails = await TeamsInfo.getMeetingInfo(context);
      console.log(meeting);
      // var graphHelper = new GraphHelper();
      const openai = new OpenAIApi(configuration);
      console.log("OpenAI API Initialized" + openai);

      await context.sendActivity("Meeting ended");
      await next();
    });

    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.");
      const graphHelper = new GraphHelper();
      if(context.activity.text.includes("summary")) {
        const meetingDetails: any = await TeamsInfo.getMeetingInfo(context);
        console.log(meetingDetails);
        var result = await graphHelper.GetMeetingTranscriptionsAsync(meetingDetails.details.msGraphResourceId);
        console.log(result);
        if(result) {
          let txt = "please create meeting summary for the following transcript?\n" + result;
          const removedMentionText = TurnContext.removeRecipientMention(context.activity);
          if (removedMentionText) {
            // Remove the line break
            txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();
          }
          const openai = new OpenAIApi(configuration);
    
          const response = await openai.createCompletion({
            model: "text-davinci-003",
            prompt: txt,
            temperature: 0,
            max_tokens: 2048,
          });
    
          await context.sendActivity(response.data.choices[0].text);
          
        } else {
          await context.sendActivity("No transcript available!");
        }
      } else {
        let txt = context.activity.text;
          const removedMentionText = TurnContext.removeRecipientMention(context.activity);
          if (removedMentionText) {
            // Remove the line break
            txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();
          }
          const openai = new OpenAIApi(configuration);
    
          const response = await openai.createCompletion({
            model: "text-davinci-003",
            prompt: txt,
            temperature: 0,
            max_tokens: 2048,
          });
    
          await context.sendActivity(response.data.choices[0].text);
      }
      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });

    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id) {
          const card = AdaptiveCards.declareWithoutData(rawWelcomeCard).render();
          await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          break;
        }
      }
      await next();
    });
  }
}
