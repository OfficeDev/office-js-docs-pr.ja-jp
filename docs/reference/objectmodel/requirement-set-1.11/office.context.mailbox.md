---
title: Office.context.mailbox - 要件セット 1.11
description: Outlook API 要件セット 1.11 バージョンのメールボックス オブジェクト モデル。
ms.date: 11/01/2021
ms.localizationpriority: medium
---

# <a name="mailbox-requirement-set-111"></a>メールボックス (要件セット 1.11)

### <a name="officecontextmailbox"></a>[Office](office.md)[.context](office.context.md).mailbox

Microsoft Outlook の Outlook アドイン オブジェクト モデルへのアクセスを提供します。

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[最小限のアクセス許可レベル](../../../outlook/understanding-outlook-add-in-permissions.md)| 制限あり|
|[適用可能な Outlook のモード](../../../outlook/outlook-add-ins-overview.md#extension-points)| 新規作成または閲覧|

## <a name="properties"></a>プロパティ

| プロパティ | 最小値<br>アクセス許可レベル | モード | 戻り値の種類 | 最小値<br>要件セット |
|---|---|---|---|:---:|
| [診断](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#outlook-office-mailbox-diagnostics-member) | ReadItem | 作成<br>読み取り | [Diagnostics](/javascript/api/outlook/office.diagnostics?view=outlook-js-1.11&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ewsUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#outlook-office-mailbox-ewsurl-member) | ReadItem | 作成<br>読み取り | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [item](office.context.mailbox.item.md) | Restricted | 作成<br>読み取り | [Item](/javascript/api/outlook/office.item?view=outlook-js-1.11&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [masterCategories](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#outlook-office-mailbox-mastercategories-member) | ReadWriteMailbox | 作成<br>読み取り | [MasterCategories](/javascript/api/outlook/office.mastercategories?view=outlook-js-1.11&preserve-view=true) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| [restUrl](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#outlook-office-mailbox-resturl-member) | ReadItem | 作成<br>読み取り | String | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [userProfile](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#outlook-office-mailbox-userprofile-member) | ReadItem | 作成<br>読み取り | [UserProfile](/javascript/api/outlook/office.userprofile?view=outlook-js-1.11&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>メソッド

| Method | 最小値<br>アクセス許可レベル | モード | 最小値<br>要件セット |
|---|---|---|:---:|
| [addHandlerAsync(eventType, handler, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#outlook-office-mailbox-addhandlerasync-member(1)) | ReadItem | 作成<br>読み取り | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [convertToEwsId(itemId, restVersion)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#outlook-office-mailbox-converttoewsid-member(1)) | Restricted | 作成<br>読み取り | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToLocalClientTime(timeValue)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#outlook-office-mailbox-converttolocalclienttime-member(1)) | ReadItem | 作成<br>読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [convertToRestId(itemId, restVersion)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#outlook-office-mailbox-converttorestid-member(1)) | Restricted | 作成<br>読み取り | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| [convertToUtcClientTime(input)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#outlook-office-mailbox-converttoutcclienttime-member(1)) | ReadItem | 作成<br>読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#outlook-office-mailbox-displayappointmentform-member(1)) | ReadItem | 作成<br>読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayAppointmentFormAsync(itemId, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#outlook-office-mailbox-displayappointmentformasync-member(1)) | ReadItem | 作成<br>読み取り | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| [displayMessageForm(itemId)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#outlook-office-mailbox-displaymessageform-member(1)) | ReadItem | 作成<br>読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayMessageFormAsync(itemId, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#outlook-office-mailbox-displaymessageformasync-member(1)) | ReadItem | 作成<br>読み取り | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| [displayNewAppointmentForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#outlook-office-mailbox-displaynewappointmentform-member(1)) | ReadItem | 読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayNewAppointmentFormAsync(parameters, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#outlook-office-mailbox-displaynewappointmentformasync-member(1)) | ReadItem | 読み取り | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| [displayNewMessageForm(parameters)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#outlook-office-mailbox-displaynewmessageform-member(1)) | ReadItem | 読み取り | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| [displayNewMessageFormAsync(parameters, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#outlook-office-mailbox-displaynewmessageformasync-member(1)) | ReadItem | 読み取り | [1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md) |
| [getCallbackTokenAsync([options], callback)](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#outlook-office-mailbox-getcallbacktokenasync-member(1)) | ReadItem | 作成<br>読み取り | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [getCallbackTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#outlook-office-mailbox-getcallbacktokenasync-member(2)) | ReadItem | 作成<br>読み取り | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md)<br>[1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [getUserIdentityTokenAsync(callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#outlook-office-mailbox-getuseridentitytokenasync-member(1)) | ReadItem | 作成<br>読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [makeEwsRequestAsync(data, callback, [userContext])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#outlook-office-mailbox-makeewsrequestasync-member(1)) | ReadWriteMailbox | 作成<br>読み取り | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [removeHandlerAsync(eventType, [options], [callback])](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#outlook-office-mailbox-removehandlerasync-member(1)) | ReadItem | 作成<br>読み取り | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |

## <a name="events"></a>イベント

[addHandlerAsync と removeHandlerAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#outlook-office-mailbox-addhandlerasync-member(1)) をそれぞれ使用して、次のイベントをサブスクライブおよび[サブスクライブ](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true#outlook-office-mailbox-removehandlerasync-member(1))解除します。

> [!IMPORTANT]
> イベントは、作業ウィンドウの実装でのみ使用できます。

| [イベント](/javascript/api/office/office.eventtype?view=outlook-js-1.11&preserve-view=true) | 説明 | 最小値<br>要件セット |
|---|---|:---:|
|`ItemChanged`| 作業ウィンドウOutlook表示する場合は、別のアイテムが選択されています。 | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
