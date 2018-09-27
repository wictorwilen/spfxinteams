import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { escape, cloneDeep } from '@microsoft/sp-lodash-subset';

import styles from './CalendarWebPart.module.scss';
import * as strings from 'CalendarWebPartStrings';
import { SPHttpClientConfiguration, SPHttpClient } from '@microsoft/sp-http';
import { AdaptiveCard, HostConfig, SubmitAction, OpenUrlAction, Action } from 'adaptivecards';
import eventTemplate from './EventTemplate';
import { MSGraphClient } from '@microsoft/sp-http';

const hostConfigTeams = require('./host-config-teams.json');
const hostConfigSP = require('./host-config-sp.json');

//#region REST response declarations
interface EventResponse {
  '@odata.context': string;
  value: Event[];
}

interface Event {
  '@odata.type': string;
  '@odata.id': string;
  '@odata.etag': string;
  '@odata.editLink': string;
  FileSystemObjectType: number;
  Id: number;
  ServerRedirectedEmbedUri?: any;
  ServerRedirectedEmbedUrl: string;
  ID: number;
  ContentTypeId: string;
  Title: string;
  Modified: string;
  Created: string;
  AuthorId: number;
  EditorId: number;
  OData__UIVersionString: string;
  Attachments: boolean;
  GUID: string;
  ComplianceAssetId?: any;
  Location: string;
  Geolocation?: any;
  EventDate: string;
  EndDate: string;
  Description?: string;
  fAllDayEvent: boolean;
  fRecurrence: boolean;
  ParticipantsPickerId?: any;
  ParticipantsPickerStringId?: any;
  Category?: any;
  FreeBusy?: any;
  Overbook?: any;
  BannerUrl?: any;
}

interface ChannelsResponse {
  '@odata.context': string;
  value: Channel[];
}

interface Channel {
  id: string;
  displayName: string;
  description?: any;
}
//#endregion

export interface ICalendarWebPartProps {
  calendarId: string;
  //itemsToShow: number;
}

export default class CalendarWebPart extends BaseClientSideWebPart<ICalendarWebPartProps> {

  public render(): void {

    if (this.isTeamsRemoveDialog()) {
      // Implement Remove logic here;
      this.domElement.innerHTML = 'Thank you for using this awesome Web Part';
      console.info('Web Part rendered in Remove dialog');
      return;
    }

    if (this.isTeamsConfigureDialog()) {
      this.domElement.innerHTML = `You're in the Teams Tab configure dialog, so rendering something is just a waste of electrons`;
      console.info('Web Part rendered in Configure dialog');
      return;
    }

    if (this.isInTeams()) {
      // Show all the Teams details we have
      console.log((<any>this.context.pageContext).teams);
    }

    if (this.properties.calendarId === undefined || this.properties.calendarId.length === 0) {
      // Render a default message if the WP is not configured
      this.domElement.innerHTML = `You need to configure this Web Part to show a calendar`;
    } else {
      this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'calendar events');

      // Read the calendar events
      this.context.spHttpClient
        .get(`${this.context.pageContext.web.absoluteUrl}/_api/lists(guid'${this.properties.calendarId}')/Items`,
          SPHttpClient.configurations.v1)
        .then(res => {
          return res.json();
        }).then((events: EventResponse) => {
          if (events.value.length == 0) {
            this.domElement.innerHTML = "No events found";
            return;
          }

          // Create Adaptive cards out of each event
          const cards = events.value.map((event: Event) => {
            let card = new AdaptiveCard();
            // Add the action handlers for the adaptive card
            card.onExecuteAction = (action: Action) => {
              switch (action.id) {
                case "comment":
                  let submitAction: SubmitAction = action as SubmitAction;
                  let channel;
                  if (this.isInTeams()) {
                    // if we're in Teams we already know the channel
                    channel = (<any>this.context.pageContext).teams.channelId;
                  } else {
                    // If we're not in Teams then use the channel submitted in the action
                    channel = (<any>submitAction.data).channel;
                  }

                  // Use the Microsfot graph to create a new thread
                  this.context.msGraphClientFactory.getClient().then((client: MSGraphClient) => {
                    client
                      .api(`/teams/${this.context.pageContext.legacyPageContext.groupId}/channels/${channel}/chatthreads`)
                      .version('beta')
                      .post(
                        {
                          "rootMessage": {
                            "body": {
                              "contentType": 1,
                              "content": `${(<any>submitAction.data).comment}`
                            }
                          }
                        },
                        (error, response: any, rawResponse?: any) => {
                          alert('Done');
                        });
                  });
                  break;
                case "outlook":
                  alert('Add to Outlook calendar');
                  break;
              }
            };

            // set the UX of the adaptive card based on the host (Teams or SP)
            card.hostConfig = this.isInTeams ? new HostConfig(hostConfigTeams) : new HostConfig(hostConfigSP);

            // Clone the adaptive card event template
            let data = cloneDeep(eventTemplate);

            // Add the data to the template
            data.body[0].text = event.Title;
            data.body[1].text = event.Location;
            data.body[2].text = `${event.EventDate} - ${event.EndDate}`;

            // Modify the card based on where it's hosted
            return new Promise<HTMLElement>((resolve, reject) => {
              if (this.isInTeams()) {
                // If we're in Microsoft Teams - remove the channel selector container
                (<any[]>data.actions[1].card.body).shift();
                card.parse(data);
                let elm = card.render();
                resolve(elm);
              } else {
                // If we're in SharePoint, first find out the actual theme used on the site
                this.context.spHttpClient
                  .get(`${this.context.pageContext.web.absoluteUrl}/_api/SP.Web.GetContextWebThemeData?noImages=true&lcid=en-US`,
                    SPHttpClient.configurations.v1)
                  .then(res => { return res.json(); })
                  .then(json => {
                    // Update host config based on SharePoint theme
                    let theme = JSON.parse(json.value);
                    let themeColor;
                    if (theme.Palette.Colors.themePrimary) {
                      themeColor = theme.Palette.Colors.themePrimary.DefaultColor;
                    } else {
                      themeColor = theme.Palette.Colors.AccentText.DefaultColor;
                    }
                    card.hostConfig.containerStyles.default.foregroundColors["accent"].default = this.rgbToHex(themeColor.R, themeColor.G, themeColor.B);

                    if (!this.isO365Group()) {
                      // If we're in a Team site not connected to an Office 365 Group,
                      // then remove the comment function as we can't have any associated team
                      (<any[]>data.actions).pop(); // remove the comment function
                      card.parse(data);
                      let elm = card.render();
                      resolve(elm);
                    } else {
                      // If we're in a group connected site, try to get the channels
                      this.context.msGraphClientFactory.getClient().then((client: MSGraphClient) => {
                        client
                          .api(`/teams/${this.context.pageContext.legacyPageContext.groupId}/channels`)
                          .version('beta')
                          .get((error, response: ChannelsResponse, rawResponse?: any) => {
                            if (!error) {
                              // Add the channels
                              let choices = response.value.map(r => {
                                return {
                                  "title": r.displayName,
                                  "value": r.id
                                };
                              });
                              data.actions[1].card.body[0].items[1].choices = choices;
                            } else {
                              // if we get an error - assume there is no associated team site
                              (<any[]>data.actions).pop(); // remove the comment function
                            }
                            card.parse(data);
                            let elm = card.render();
                            resolve(elm);
                          });
                      });
                    }
                  });
              }
            });
          });

          // Render all the cards
          Promise.all(cards).then(elements => {
            this.domElement.innerHTML = `<div class="placeholder"></div>`;
            let placeholder = this.domElement.getElementsByClassName('placeholder')[0];
            elements.forEach(c => {
              placeholder.appendChild(c);
            });

          });
        });
    }
  }

  private componentToHex(c): string {
    var hex = c.toString(16);
    return hex.length == 1 ? "0" + hex : hex;
  }

  private rgbToHex(r, g, b): string {
    return "#" + this.componentToHex(r) + this.componentToHex(g) + this.componentToHex(b);
  }

  private isInTeams(): boolean {
    return (<any>this.context.pageContext).teams !== undefined;
  }

  private isO365Group(): boolean {
    return (<any>this.context.pageContext.legacyPageContext).groupId !== null;
  }

  private isTeamsRemoveDialog(): boolean {
    return this.getParameterByName('removeTab') != null;
  }

  private isTeamsConfigureDialog(): boolean {
    return this.getParameterByName('teams') != null &&
      this.getParameterByName('openPropertyPane') === 'true' &&
      this.getParameterByName('componentId') != null;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getParameterByName(name: string, url?: string): string {
    if (!url) url = window.location.href;
    name = name.replace(/[\[\]]/g, '\\$&');
    var regex = new RegExp('[?&]' + name + '(=([^&#]*)|&|#|$)'),
      results = regex.exec(url);
    if (!results) return null;
    if (!results[2]) return '';
    return decodeURIComponent(results[2].replace(/\+/g, ' '));
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyFieldListPicker('calendarId', {
                  label: 'Select a list',
                  baseTemplate: 106,
                  selectedList: this.properties.calendarId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
