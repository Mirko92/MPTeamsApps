import { CardFactory, MessagingExtensionAttachment } from "botbuilder";
import * as ACData from "adaptivecards-templating";

export class CardUtils {

  public static getPlanetDetailCard(selectedPlanet: any): MessagingExtensionAttachment {
    // load card template
    const adaptiveCardSource: any = require("./planetDisplayCard.json");
    // Create a Template instance from the template payload
    const template = new ACData.Template(adaptiveCardSource);
    // bind the data to the card template
    const boundTemplate = template.expand({ $root: selectedPlanet });
    // load the adaptive card
    const adaptiveCard = CardFactory.adaptiveCard(boundTemplate);
    // return the adaptive card
    return adaptiveCard;
  }
}