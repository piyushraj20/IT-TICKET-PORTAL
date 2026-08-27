// =====================================================================
// ItTicketPortalWebPart.ts
// This is the SharePoint side of the solution. Its whole job is:
//   1. start PnPjs with the page context (onInit)
//   2. create the React component and drop it into the page (render)
//   3. describe the settings that appear in the property pane
// Keep business logic OUT of this file.
// =====================================================================

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPFI } from '@pnp/sp';
import TicketPortal from './components/TicketPortal';
import { ITicketPortalProps } from './components/ITicketPortalProps';
import { getSP } from './services/pnpjsConfig';

export interface IItTicketPortalWebPartProps {
  ticketsList: string;
  commentsList: string;
  historyList: string;
  supportGroupName: string;
}

export default class ItTicketPortalWebPart
  extends BaseClientSideWebPart<IItTicketPortalWebPartProps> {

  private _sp!: SPFI;

  protected async onInit(): Promise<void> {
    await super.onInit();
    this._sp = getSP(this.context);
  }

  public render(): void {
    const element: React.ReactElement<ITicketPortalProps> = React.createElement(
      TicketPortal,
      {
        sp: this._sp,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        ticketsList: this.properties.ticketsList || 'ITTickets',
        commentsList: this.properties.commentsList || 'TicketComments',
        historyList: this.properties.historyList || 'TicketHistory',
        supportGroupName: this.properties.supportGroupName || 'IT Support Team',
        userDisplayName: this.context.pageContext.user.displayName,
        userEmail: this.context.pageContext.user.email
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'Point the web part at your SharePoint lists.' },
          groups: [
            {
              groupName: 'Lists',
              groupFields: [
                PropertyPaneTextField('ticketsList', { label: 'Tickets list name' }),
                PropertyPaneTextField('commentsList', { label: 'Comments list name' }),
                PropertyPaneTextField('historyList', { label: 'History list name' })
              ]
            },
            {
              groupName: 'Permissions',
              groupFields: [
                PropertyPaneTextField('supportGroupName', {
                  label: 'Support group name',
                  description: 'Members of this SharePoint group get the agent controls.'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
