# Last Commenter Field Customizer

## Summary

This SharePoint Framework field customizer displays the email address of the last person who commented on documents (in document libraries) or the last person who modified items (in regular lists). It provides a quick way to see who last interacted with each item in your SharePoint lists and libraries.

**Key Features:**
- Shows last commenter email for document libraries with comments enabled
- Falls back to last modified user for regular lists
- Caches results for improved performance
- Displays loading states and error handling
- Works with any field type (text, number, etc.)

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.22.0-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

- SharePoint Online or SharePoint Server 2019+
- SharePoint Framework development environment
- PnP PowerShell (for deployment)

## Solution

| Solution               | Author(s) |
| ---------------------- | --------- |
| last-commenter-customizer | SPFx Community |

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.0     | December 12, 2025 | Initial release |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - `npm install`
  - `npm run build`
  - Upload the generated `.sppkg` file to your SharePoint App Catalog
  - Associate the field customizer with a field using PowerShell or REST API

> Include any additional steps as needed.

Other build commands can be listed using `heft --help`.

## Features

### Attach via Browser Console (Current List)
Field customizer binding is stored on a specific field in a specific list.
Run this from the target list page to bind without hardcoding list title.

```javascript
(async () => {
  const listIdRaw = _spPageContextInfo.pageListId; // current list, often wrapped in {}
  const fieldInternalName = "ID"; // change per target field
  const componentId = "680d1d6e-610a-4a21-8d98-e5edccd066d7";

  if (!listIdRaw) {
    throw new Error("No list context found. Open a list view page and try again.");
  }

  // SharePoint often exposes pageListId as "{guid}". REST expects guid without braces.
  const listId = String(listIdRaw).replace(/[{}]/g, "");

  const webRel = (_spPageContextInfo.webServerRelativeUrl || "").replace(/\/$/, "");
  const apiBase = webRel + "/_api";

  const digestRes = await fetch(apiBase + "/contextinfo", {
    method: "POST",
    headers: { Accept: "application/json;odata=nometadata" }
  });
  const digest = (await digestRes.json()).FormDigestValue;

  const fieldUrl =
    apiBase +
    "/web/lists(guid'" +
    listId +
    "')/fields/getbyinternalnameortitle('" +
    encodeURIComponent(fieldInternalName) +
    "')";

  const res = await fetch(fieldUrl, {
    method: "POST",
    headers: {
      Accept: "application/json;odata=nometadata",
      "Content-Type": "application/json;odata=nometadata",
      "X-RequestDigest": digest,
      "IF-MATCH": "*",
      "X-HTTP-Method": "MERGE"
    },
    body: JSON.stringify({
      ClientSideComponentId: componentId,
      ClientSideComponentProperties: JSON.stringify({})
    })
  });

  console.log(res.ok ? "Attached" : `Failed ${res.status}: ${await res.text()}`);
})();
```

### Unbind via Browser Console (Current List)
Run this from the target list page to remove the field customizer binding.

```javascript
(async () => {
  const listIdRaw = _spPageContextInfo.pageListId;
  const fieldInternalName = "ID"; // change per target field

  if (!listIdRaw) {
    throw new Error("No list context found. Open a list view page and try again.");
  }

  const listId = String(listIdRaw).replace(/[{}]/g, "");
  const webRel = (_spPageContextInfo.webServerRelativeUrl || "").replace(/\/$/, "");
  const apiBase = webRel + "/_api";

  const digestRes = await fetch(apiBase + "/contextinfo", {
    method: "POST",
    headers: { Accept: "application/json;odata=nometadata" }
  });
  const digest = (await digestRes.json()).FormDigestValue;

  const fieldUrl =
    apiBase +
    "/web/lists(guid'" +
    listId +
    "')/fields/getbyinternalnameortitle('" +
    encodeURIComponent(fieldInternalName) +
    "')";

  const res = await fetch(fieldUrl, {
    method: "POST",
    headers: {
      Accept: "application/json;odata=nometadata",
      "Content-Type": "application/json;odata=nometadata",
      "X-RequestDigest": digest,
      "IF-MATCH": "*",
      "X-HTTP-Method": "MERGE"
    },
    body: JSON.stringify({
      ClientSideComponentId: null,
      ClientSideComponentProperties: null
    })
  });

  if (!res.ok) {
    const retry = await fetch(fieldUrl, {
      method: "POST",
      headers: {
        Accept: "application/json;odata=nometadata",
        "Content-Type": "application/json;odata=nometadata",
        "X-RequestDigest": digest,
        "IF-MATCH": "*",
        "X-HTTP-Method": "MERGE"
      },
      body: JSON.stringify({
        ClientSideComponentId: "00000000-0000-0000-0000-000000000000",
        ClientSideComponentProperties: null
      })
    });

    if (!retry.ok) {
      console.error(`Failed ${retry.status}: ${await retry.text()}`);
      return;
    }
  }

  console.log("Unbound");
})();
```

Tip for many lists: keep one script and only change `fieldInternalName`, then run it from each list where you want to bind or unbind.

### Last Commenter Display
The field customizer displays the most recent comment information for items in your list or library. The display includes:

- **Timestamp**: Date and time of the last comment (formatted as MM/DD/YYYY HH:MM)
- **Commenter**: Full name and email address of the person who made the last comment
- **Admin Status**: Indicates whether the commenter is an administrator

### Admin Fields (admin_1 and admin_2)
The customizer supports two optional people picker fields named `admin_1` and `admin_2` that define who is considered an administrator:

- **Field Type**: People Picker columns
- **Field Names**: Must be named exactly `admin_1` and `admin_2` (case-sensitive)
- **Purpose**: These fields identify authorized administrators for your list items
- **Impact**: The admin status in the display shows "admin: yes" if the last commenter matches either admin field, or "admin: no" otherwise

### Display Formatting
- **For Non-Admins**: The comment information is displayed with a light blue background (#D4E7F6) with rounded corners and 4px padding
- **For Admins**: The comment information is displayed with a transparent background
- **Format**: 
  ```
  at: [Date and Time]
  by: [Full Name] [Email]
  admin: [yes/no]
  ```

### Caching and Performance
- Comment data is cached in memory to reduce API calls and improve performance
- Admin field values are also cached per list item

### Error Handling
- Gracefully handles missing comments (returns empty display)
- Handles cases where admin fields don't exist on the list
- Provides console logging for debugging purposes

> Notice that better pictures and documentation will increase the sample usage and the value you are providing for others. Thanks for your submissions advance.

> Share your web part with others through Microsoft 365 Patterns and Practices program to get visibility and exposure. More details on the community, open-source projects and other activities from http://aka.ms/m365pnp.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
- [Heft Documentation](https://heft.rushstack.io/)