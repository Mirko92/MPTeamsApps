import {
  TeamsActivityHandler,
  TurnContext,
  CardFactory, 

  MessageFactory,
  MessagingExtensionAction, 
  MessagingExtensionActionResponse, 
  MessagingExtensionAttachment
} from "botbuilder";

import * as Util from "util";
import * as debug from "debug";

const TextEncoder = Util.TextEncoder;
const log = debug("msteams");

export class PlanetBot extends TeamsActivityHandler {


  protected handleTeamsMessagingExtensionFetchTask(context: TurnContext, action: MessagingExtensionAction): Promise<MessagingExtensionActionResponse> {
    console.log("handleTeamsMessagingExtensionFetchTask")
    // load planets & sort them by their order from the sun
    const planets: any[] = require("./planets.json");

    const sortedPlanets: any = planets
      .sort( (a,b) => a.id > b.id ? 1 : -1 )
      .map((planet) => {
        return { value: planet.id, title: planet.name };
      });
  
    // load card template
    const adaptiveCardSource: any = require("./planetSelectorCard.json");
    // locate the planet selector
    const planetChoiceSet: any = adaptiveCardSource.body.find( x => x.id === "planetSelector");
    // update choice set with planets
    planetChoiceSet.choices = sortedPlanets;
    // load the adaptive card
    const adaptiveCard = CardFactory.adaptiveCard(adaptiveCardSource);
  
    const response: MessagingExtensionActionResponse = {
      task: {
        type: "continue",
        value: {
          card: adaptiveCard,
          title: "Planet Selector",
          height: 150,
          width: 500
        }
      }
    } as MessagingExtensionActionResponse;
  
    return Promise.resolve(response);
  }




  protected handleTeamsMessagingExtensionSubmitAction(context: TurnContext, action: MessagingExtensionAction): Promise<MessagingExtensionActionResponse> {
    console.log("handleTeamsMessagingExtensionSubmitAction")

    switch (action.commandId) {
      case "planetExpanderAction": {
        // load planets
        const planets: any = require("./planets.json");
        // get the selected planet
        const selectedPlanet: any = planets.filter((planet) => planet.id === action.data.planetSelector)[0];
        const adaptiveCard = this.getPlanetDetailCard(selectedPlanet);
  
        // generate the response
        return Promise.resolve({
          composeExtension: {
            type: "result",
            attachmentLayout: "list",
            attachments: [adaptiveCard]
          }
        } as MessagingExtensionActionResponse);
      }
      default:
        throw new Error("NotImplemented");
    }
  }

  private getPlanetDetailCard(selectedPlanet: any): MessagingExtensionAttachment {
    // load display card
    const adaptiveCardSource: any = require("./planetDisplayCard.json");
  
    // update planet fields in display card
    adaptiveCardSource.actions[0].url = selectedPlanet.wikiLink;
    
    find<any>(adaptiveCardSource.body, { id: "cardHeader" }).items[0].text = selectedPlanet.name;
    const cardBody: any = find<any>(adaptiveCardSource.body, { id: "cardBody" });
    find<any>(cardBody.items, { id: "planetSummary" }).text = selectedPlanet.summary;
    find<any>(cardBody.items, { id: "imageAttribution" }).text = "*Image attribution: " + selectedPlanet.imageAlt + "*";
    const cardDetails: any = find<any>(cardBody.items, { id: "planetDetails" });
    cardDetails.columns[0].items[0].url = selectedPlanet.imageLink;
    find<any>(cardDetails.columns[1].items[0].facts, { id: "orderFromSun" }).value = selectedPlanet.id;
    find<any>(cardDetails.columns[1].items[0].facts, { id: "planetNumSatellites" }).value = selectedPlanet.numSatellites;
    find<any>(cardDetails.columns[1].items[0].facts, { id: "solarOrbitYears" }).value = selectedPlanet.solarOrbitYears;
    find<any>(cardDetails.columns[1].items[0].facts, { id: "solarOrbitAvgDistanceKm" }).value = Number(selectedPlanet.solarOrbitAvgDistanceKm).toLocaleString();
  
    // return the adaptive card
    return CardFactory.adaptiveCard(adaptiveCardSource);
  }
}



function find<T>(object: T[], selector: Partial<T>): T {
  const key = Object.keys(selector)[0];

  return object.find( o => o[key] === selector[key])!;
}
