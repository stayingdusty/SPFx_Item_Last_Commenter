# Deploy Last Commenter Field Customizer

## Step 1: Upload to SharePoint App Catalog
1. Go to your SharePoint Admin Center
2. Navigate to **More features** → **Apps** → **App Catalog**
3. Upload the `last-commenter-customizer.sppkg` file
4. Make it available to all sites

## Step 2: Associate with a Field (PowerShell)
Use this PowerShell script to associate the field customizer with a specific field:

```powershell
# Connect to SharePoint Online
Connect-PnPOnline -Url "https://yourtenant.sharepoint.com/sites/yoursite" -Interactive

# Field customizer ID from manifest
$fieldCustomizerId = "680d1d6e-610a-4a21-8d98-e5edccd066d7"

# Associate with a field in a list
Set-PnPField -List "Your List Name" -Identity "YourFieldName" -FieldCustomizer $fieldCustomizerId

# Or associate with a site column
Set-PnPField -Identity "YourSiteColumnName" -FieldCustomizer $fieldCustomizerId
```

## Step 3: View in Action
1. Navigate to your SharePoint list/library
2. The field will show the email of the last commenter on each item
3. Initially shows "Loading..." while fetching data
4. Displays the commenter's email once loaded

## Alternative: REST API Association
If you prefer REST API:

```javascript
// Associate field customizer via REST API
const fieldCustomizerId = "680d1d6e-610a-4a21-8d98-e5edccd066d7";

fetch(`${siteUrl}/_api/web/lists/getbytitle('YourList')/fields/getbyinternalnameortitle('YourField')`, {
  method: 'POST',
  headers: {
    'X-HTTP-Method': 'MERGE',
    'Content-Type': 'application/json'
  },
  body: JSON.stringify({
    ClientSideComponentId: fieldCustomizerId
  })
});
```

## What It Does
- **For Document Libraries**: Shows the email address of the person who last commented on each document (if comments are enabled)
- **For Regular Lists**: Shows the email address of the person who last modified each item
- Caches results for better performance
- Displays "Loading..." while fetching data
- Shows "N/A" if no data exists or errors occur
- Gracefully falls back from comments to last modified user if comments aren't available