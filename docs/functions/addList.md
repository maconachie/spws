[**spws**](../README.md) • **Docs**

***

[spws](../globals.md) / addList

# Function: addList()

> **addList**(`listName`, `options`): `Promise`\<`Operation`\>

Creates a new list.

## Parameters

• **listName**: `string`

• **options**: `Params` = `{}`

## Returns

`Promise`\<`Operation`\>

## Link

https://learn.microsoft.com/en-us/previous-versions/office/developer/sharepoint-services/ms772560(v=office.12)

## Example

```
// Get list collection for current site
const res = await deleteList("Announcements", { webURL: "/sites/other" })
```

## Defined in

[services/lists/addList.ts:39](https://github.com/rlking1985/spws/blob/96ed2330ff15e8f8eb88949aa126d8a29c8f97dc/src/services/lists/addList.ts#L39)