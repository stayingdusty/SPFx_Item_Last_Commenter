import { Log } from '@microsoft/sp-core-library';
import { override } from '@microsoft/decorators';
import {
  BaseFieldCustomizer,
  IFieldCustomizerCellEventParameters
} from '@microsoft/sp-listview-extensibility';
import { SPHttpClient } from '@microsoft/sp-http';
import * as React from 'react';
import * as ReactDOM from 'react-dom';

export interface ILastCommenterFieldCustomizerProperties {
  // No custom properties needed for basic implementation
}

const LOG_SOURCE: string = 'LastCommenterFieldCustomizer';

export default class LastCommenterFieldCustomizer
  extends BaseFieldCustomizer<ILastCommenterFieldCustomizerProperties> {

  private _commentCache: Map<number, string> = new Map();

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Activated LastCommenterFieldCustomizer');
    return Promise.resolve();
  }

  @override
  public onRenderCell(event: IFieldCustomizerCellEventParameters): void {
    const itemId = event.listItem.getValueByName('ID');
    
    // Check cache first
    if (this._commentCache.has(itemId)) {
      const cachedEmail = this._commentCache.get(itemId);
      this.renderEmail(event.domElement, cachedEmail || 'N/A');
    } else {
      // Show loading state
      this.renderEmail(event.domElement, 'Loading...');
      
      // Fetch last commenter
      this.getLastCommenterEmail(itemId)
        .then(email => {
          this._commentCache.set(itemId, email);
          this.renderEmail(event.domElement, email);
        })
        .catch(error => {
          console.error(`Error fetching comments for item ${itemId}:`, error);
          this.renderEmail(event.domElement, 'N/A');
        });
    }
  }

  private renderEmail(element: HTMLElement, email: string): void {
    const emailElement = React.createElement(
      'div',
      { 
        style: { 
          padding: '4px',
          fontSize: '12px',
          color: email === 'Loading...' ? '#666' : '#0078d4'
        } 
      },
      email
    );
    
    ReactDOM.render(emailElement, element);
  }

  private async getLastCommenterEmail(itemId: number): Promise<string> {
    try {
      const listId = this.context.pageContext.list?.id?.toString();
      if (!listId) {
        return 'No list';
      }

      // First try to get comments from list items (modern SharePoint comments)
      try {
        const commentsApiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${listId}')/items(${itemId})/Comments?$expand=author&$orderby=createdDate desc&$top=1`;

        console.log('Trying comments API:', commentsApiUrl);

        const commentsResponse = await this.context.spHttpClient.get(
          commentsApiUrl,
          SPHttpClient.configurations.v1
        );

        console.log('Comments API response status:', commentsResponse.status);

        if (commentsResponse.ok) {
          const commentsData = await commentsResponse.json();
          console.log('Comments data:', commentsData);

          if (commentsData.value && commentsData.value.length > 0) {
            const lastComment = commentsData.value[0];
            const email = lastComment.author?.email || lastComment.author?.title || 'No email';
            console.log('Found comment author:', email);
            return email;
          } else {
            console.log('No comments found for item', itemId);
          }
        } else {
          console.log('Comments API failed with status:', commentsResponse.status);
          const errorText = await commentsResponse.text();
          console.log('Comments API error:', errorText);
        }
      } catch (commentsError) {
        console.log('Comments API exception:', commentsError);
      }

      // Fallback: show the last modified user
      console.log('Falling back to last modified user for item', itemId);
      const itemApiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${listId}')/items(${itemId})?$select=Editor/EMail,Editor/Title&$expand=Editor`;

      const response = await this.context.spHttpClient.get(
        itemApiUrl,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const data = await response.json();
      const email = data.Editor?.EMail || data.Editor?.Title || 'No editor';
      console.log('Last modified user:', email);
      return email;
    } catch (error) {
      console.error('Error in getLastCommenterEmail:', error);
      return 'Error';
    }
  }

  @override
  public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
    ReactDOM.unmountComponentAtNode(event.domElement);
    super.onDisposeCell(event);
  }
}
