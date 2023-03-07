import {
  TeamsActivityHandler,
  TurnContext,
  CardFactory, 

  // ACTION
  MessagingExtensionAction, 
  MessagingExtensionActionResponse, 
  MessagingExtensionAttachment,

  // SEARCH
  MessagingExtensionQuery,
  MessagingExtensionResponse,
  MessagingExtensionResult,

  // LINK
  AppBasedLinkQuery
} from "botbuilder";

import * as Util  from "util";
import * as debug from "debug";
import { IPlanet } from "./IPlanet";
import { ContainerState } from "@microsoft/teams-js";

const TextEncoder = Util.TextEncoder;
const log = debug("msteams");

export class PlanetBot extends TeamsActivityHandler {

  //#region ACTION

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
  //#endregion


  //#region SEARCH

  protected handleTeamsMessagingExtensionQuery(context: TurnContext, query: MessagingExtensionQuery): Promise<MessagingExtensionResponse> {
    console.log("handleTeamsMessagingExtensionQuery");
    // get the search query
    let searchQuery = "";
    if (
        query?.parameters &&
        query.parameters[0].name === "searchKeyword" &&
        query.parameters[0].value) {
      searchQuery = query.parameters[0].value.trim().toLowerCase();
    }

    // load planets
    const planets: any = require("./planets.json");
    // search results
    let queryResults: string[] = [];

    switch (searchQuery) {
      case "inner":
        // get all planets inside asteroid belt
        queryResults = planets.filter((planet) => planet.id <= 4);
        break;
      case "outer":
        // get all planets outside asteroid belt
        queryResults = planets.filter((planet) => planet.id > 4);
        break;
      default:
        // get the specified planet
        queryResults.push(planets.filter((planet) => planet.name.toLowerCase() === searchQuery)[0]);
    }

    // get results as cards
    let searchResultsCards: MessagingExtensionAttachment[] = [];
    queryResults.forEach((planet) => {
      searchResultsCards.push(this.getPlanetResultCard(planet));
    });

    /**
     * The type property can be one of the following options:
     *  - result:   displays a list of the search results
     *  - message:  displays a plain text message
     *  - auth:     prompts the user to authenticate
     *  - config:   prompts the user to set up the messaging extension
     * 
     * If type is set to message, an extra property text can be used to set the plain text message displayed.
     * 
     * The attachmentLayout property can be either a list of results containing thumbnails, titles and text fields, 
     * or a grid of thumbnail images.
     * 
     * When type is set to auth or config, use the suggestedActions property to suggest extra actions to do.
     */
    let response: MessagingExtensionResponse = <MessagingExtensionResponse>{
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: searchResultsCards
      }
    };

    return Promise.resolve(response);
  }

  //#endregion

  //#region LINK

  protected async handleTeamsAppBasedLinkQuery(context: TurnContext, query: AppBasedLinkQuery): Promise<MessagingExtensionResponse> {
    console.log("handleTeamsAppBasedLinkQuery:", query.url);

    if (!query.url) {
      return Promise.reject(null);
    }

    const url = new URL(query.url);

    const planetName: string = url.pathname?.substring(url.pathname.lastIndexOf("/") + 1) ?? "";

    // get the selected planet
    const selectedPlanet: IPlanet = this.getPlanetByName(planetName);

    if (!selectedPlanet) {
      return Promise.reject(`Planet ${planetName} not found`);
    }

    const typeParam = url.searchParams.get('type')?.toLowerCase() ?? "herocard";
    
    let attachment; 

    switch(typeParam) {
      case "herocard":
        attachment = this.getPlanetResultCard(selectedPlanet);
        break;
      case "adaptivecard":
        attachment = this.createAdaptiveCard(selectedPlanet);
        break;
      case "thumbnailcard":
        break;
    }

    console.log(`${typeParam} JSON: ${JSON.stringify(attachment)}`);

    // generate the response
    return Promise.resolve(<MessagingExtensionActionResponse>{
      cacheInfo: {
        cacheDuration: 1,
      },
      composeExtension: {
        type: "result",
        attachmentLayout: "grid",
        attachments: [attachment]
      }
    });

  }
  
  //#endregion

  private createAdaptiveCard(planet: IPlanet) {
    const ac = CardFactory.adaptiveCard({
      "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
      "type"   : "AdaptiveCard",
      "version": "1.0",
      "body": [
        {
          "type": "TextBlock",
          "text": "ADAPTIVE CARD",
        },
        {
          "type": "TextBlock",
          "text": planet.name,
        },
        {
          "type": "TextBlock",
          "text": planet.summary,
        }
      ],
      "actions": [
        {
          "type": "Action.Submit",
          "title": "OK"
        }
      ]
    });

    return ac;
  }

  private getPlanetByName(planetName: string) {
    console.log(`Looking for planet named: ${planetName}`);
    
    // load planets
    const planets: any = require("./planets.json");

    return planets.filter(
      (planet) => planet.name.toLowerCase() === planetName.toLowerCase()
    )[0];
  }

  private getPlanetDetailCard(selectedPlanet: any): MessagingExtensionAttachment {
    // load display card
    const adaptiveCardSource: any = require("./planetDisplayCard.json");
  
    // update planet fields in display card
    adaptiveCardSource.actions[0].url = selectedPlanet.wikiLink;
    
    // HEADER
    find<any>(adaptiveCardSource.body, { id: "cardHeader" }).items[0].text = selectedPlanet.name + "CIAO";

    // BODY
    const cardBody: any = find<any>(adaptiveCardSource.body, { id: "cardBody" });
    find<any>(cardBody.items, { id: "planetSummary" }).text = "SUM" + selectedPlanet.summary;

    // IMG
    find<any>(cardBody.items, { id: "imageAttribution" }).text = "*Image attribution: " + selectedPlanet.imageAlt + "*";
    const cardDetails: any = find<any>(cardBody.items, { id: "planetDetails" });
    cardDetails.columns[0].items[0].url = selectedPlanet.imageLink;

    find<any>(cardDetails.columns[1].items[0].facts, { id: "orderFromSun" }).value            = selectedPlanet.id;
    find<any>(cardDetails.columns[1].items[0].facts, { id: "planetNumSatellites" }).value     = selectedPlanet.numSatellites;
    find<any>(cardDetails.columns[1].items[0].facts, { id: "solarOrbitYears" }).value         = selectedPlanet.solarOrbitYears;
    find<any>(cardDetails.columns[1].items[0].facts, { id: "solarOrbitAvgDistanceKm" }).value = Number(selectedPlanet.solarOrbitAvgDistanceKm).toLocaleString();
  
    // return the adaptive card
    return CardFactory.adaptiveCard(adaptiveCardSource);
  }

  private getThumbnailPlanetDetailCard(selectedPlanet: any): MessagingExtensionAttachment {
    // load display card
    const adaptiveCardSource: any = require("./planetDisplayCard.json");
  
    // update planet fields in display card
    adaptiveCardSource.actions[0].url = selectedPlanet.wikiLink;
    
    // HEADER
    find<any>(adaptiveCardSource.body, { id: "cardHeader" }).items[0].text = selectedPlanet.name + "CIAO";

    // BODY
    const cardBody: any = find<any>(adaptiveCardSource.body, { id: "cardBody" });
    find<any>(cardBody.items, { id: "planetSummary" }).text = "SUM" + selectedPlanet.summary;

    // IMG
    find<any>(cardBody.items, { id: "imageAttribution" }).text = "*Image attribution: " + selectedPlanet.imageAlt + "*";
    const cardDetails: any = find<any>(cardBody.items, { id: "planetDetails" });
    cardDetails.columns[0].items[0].url = selectedPlanet.imageLink;

    find<any>(cardDetails.columns[1].items[0].facts, { id: "orderFromSun" }).value            = selectedPlanet.id;
    find<any>(cardDetails.columns[1].items[0].facts, { id: "planetNumSatellites" }).value     = selectedPlanet.numSatellites;
    find<any>(cardDetails.columns[1].items[0].facts, { id: "solarOrbitYears" }).value         = selectedPlanet.solarOrbitYears;
    find<any>(cardDetails.columns[1].items[0].facts, { id: "solarOrbitAvgDistanceKm" }).value = Number(selectedPlanet.solarOrbitAvgDistanceKm).toLocaleString();
  
    // return the adaptive card
    return CardFactory.thumbnailCard(adaptiveCardSource);
  }
  
  private getPlanetResultCard(selectedPlanet: any): MessagingExtensionAttachment {
    return CardFactory.heroCard(selectedPlanet.name, selectedPlanet.summary, [selectedPlanet.imageLink]);
  }

}



function find<T>(object: T[], selector: Partial<T>): T {
  const key = Object.keys(selector)[0];

  return object.find( o => o[key] === selector[key])!;
}
