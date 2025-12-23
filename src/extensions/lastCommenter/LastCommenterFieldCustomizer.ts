import { Log } from '@microsoft/sp-core-library';
import { override } from '@microsoft/decorators';
import {
  BaseFieldCustomizer,
  IFieldCustomizerCellEventParameters
} from '@microsoft/sp-listview-extensibility';
import { SPHttpClient } from '@microsoft/sp-http';

export interface ILastCommenterFieldCustomizerProperties {
  // No custom properties needed for basic implementation
}

const LOG_SOURCE: string = 'LastCommenterFieldCustomizer';

export default class LastCommenterFieldCustomizer
  extends BaseFieldCustomizer<ILastCommenterFieldCustomizerProperties> {

  private _commentCache: Map<number, string> = new Map();
  private _itemDataCache: Map<number, { admin1: string; admin2: string }> = new Map();

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Activated LastCommenterFieldCustomizer');
    return Promise.resolve();
  }

  @override
  public onRenderCell(event: IFieldCustomizerCellEventParameters): void {
    try {
      // Get the item ID - field customizers have limited access to list item properties
      // The most reliable way is to use the field value if it's the ID field, or get it from the context
      let itemId: number | undefined;

      // If this customizer is associated with the ID field, use the field value
      if (event.fieldValue && typeof event.fieldValue === 'number') {
        itemId = event.fieldValue;
      } else if (event.fieldValue && typeof event.fieldValue === 'string') {
        itemId = parseInt(event.fieldValue);
      }

      // If that doesn't work, try to get it from listItem (may not be available)
      if (!itemId && event.listItem) {
        try {
          itemId = event.listItem.getValueByName('ID') || event.listItem.getValueByName('id');
        } catch (e) {
          // listItem access might fail in some contexts
        }
      }

      // If still no ID, try to extract from URL or other context
      if (!itemId) {
        // Try to get from page context or URL
        const urlParams = new URLSearchParams(window.location.search);
        const itemIdParam = urlParams.get('ID');
        if (itemIdParam) {
          itemId = parseInt(itemIdParam);
        }
      }

      if (!itemId || isNaN(itemId)) {
        event.domElement.innerHTML = '<div style="padding: 4px; font-size: 11px; color: #ff0000;">No ID</div>';
        return;
      }

      // Check cache first
      if (this._commentCache.has(itemId)) {
        const cachedEmail = this._commentCache.get(itemId);
        if (cachedEmail) {
          event.domElement.innerHTML = `<div style="padding: 4px; font-size: 11px; color: #000000; line-height: 1.3;">${cachedEmail}</div>`;
        } else {
          event.domElement.innerHTML = ''; // No comments - show nothing
        }
      } else {
        // Show loading state
        event.domElement.innerHTML = '<div style="padding: 4px; font-size: 11px; color: #666;">Loading...</div>';

        // Fetch last commenter
        this.getLastCommenterEmail(itemId, event.listItem)
          .then(email => {
            this._commentCache.set(itemId, email);
            if (email) {
              event.domElement.innerHTML = `<div style="padding: 4px; font-size: 11px; color: #000000; line-height: 1.3;">${email}</div>`;
            } else {
              event.domElement.innerHTML = ''; // No comments - show nothing
            }
          })
          .catch(error => {
            console.error(`Error fetching comments for item ${itemId}:`, error);
            event.domElement.innerHTML = ''; // Error - show nothing
          });
      }
    } catch (error) {
      console.error('Error in onRenderCell:', error);
      event.domElement.innerHTML = '<div style="padding: 4px; font-size: 11px; color: #ff0000;">Error</div>';
    }
  }

  private async getLastCommenterEmail(itemId: number, listItem?: any): Promise<string> {
    try {
      const listId = this.context.pageContext.list?.id?.toString();
      if (!listId) {
        return '';
      }

      // Fetch admin_1 and admin_2 values from the list item
      let admin1Email = '';
      let admin2Email = '';

      try {
        // Fetch from API with proper expansion for people picker fields
        const itemUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${listId}')/items(${itemId})?$select=admin_1/EMail,admin_2/EMail&$expand=admin_1,admin_2`;
        const itemResponse = await this.context.spHttpClient.get(
          itemUrl,
          SPHttpClient.configurations.v1
        );

        if (itemResponse.ok) {
          const itemData = await itemResponse.json();
          
          // People picker fields return objects with EMail property
          admin1Email = itemData.admin_1?.EMail || '';
          admin2Email = itemData.admin_2?.EMail || '';
          
          console.log(`Admin fields for item ${itemId}: admin_1=${admin1Email}, admin_2=${admin2Email}`);
        }

        // Cache the admin values
        this._itemDataCache.set(itemId, { admin1: admin1Email, admin2: admin2Email });
      } catch (itemError) {
        console.warn(`Could not fetch admin fields for item ${itemId}:`, itemError);
      }

      // Only show information if there are comments
      try {
        const commentsApiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${listId}')/items(${itemId})/Comments?$expand=author&$orderby=createdDate desc&$top=1`;

        const commentsResponse = await this.context.spHttpClient.get(
          commentsApiUrl,
          SPHttpClient.configurations.v1
        );

        if (commentsResponse.ok) {
          const commentsData = await commentsResponse.json();

          if (commentsData.value && commentsData.value.length > 0) {
            const lastComment = commentsData.value[0];
            
            // Format the comment information
            const createdDate = new Date(lastComment.createdDate);
            const localDateTime = createdDate.toLocaleString([], {
              year: 'numeric',
              month: '2-digit',
              day: '2-digit',
              hour: '2-digit',
              minute: '2-digit'
            });
            
            const firstName = lastComment.author?.firstName || '';
            const lastName = lastComment.author?.lastName || '';
            const fullName = `${firstName} ${lastName}`.trim();
            const email = lastComment.author?.email || '';
            
            // Check if the last commenter is an admin
            const isAdmin = (email && (email === admin1Email || email === admin2Email));
            const adminStatus = isAdmin ? 'admin: yes' : 'admin: no';
            
            console.log(`Comparing commenter email "${email}" with admin_1 "${admin1Email}" and admin_2 "${admin2Email}". Match: ${isAdmin}`);
            
            // Apply background color styling for non-admin commenters
            const bgColor = isAdmin ? 'transparent' : '#D4E7F6';
            return `<div style="padding: 4px; background-color: ${bgColor}; border-radius: 3px;">at: ${localDateTime}<br>by: ${fullName} ${email}<br>${adminStatus}</div>`;
          }
        }
      } catch (commentsError) {
        // Comments API not available
      }

      // No comments found - return empty string to show nothing
      return '';
    } catch (error) {
      console.error('Error in getLastCommenterEmail:', error);
      return '';
    }
  }
}
