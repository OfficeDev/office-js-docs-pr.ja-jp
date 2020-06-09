---
title: Office. メールボックス-要件セット1.2
description: Outlook メールボックス API 要件セット1.2 バージョンのアイテムオブジェクトモデル。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 137edae53dfb9b23dc4e6a5697b9220d68b9484c
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44610811"
---
# <a name="item-mailbox-requirement-set-12"></a>アイテム (メールボックス要件セット 1.2)

### <a name="officecontextmailboxitem"></a>[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item

`item`は、現在選択されているメッセージ、会議出席依頼、または予定にアクセスするために使用されます。 プロパティを使用して、アイテムの種類を調べることができ `itemType` ます。

##### <a name="requirements"></a>Requirements

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../../requirement-sets/outlook-api-requirement-sets.md)|1.1|
|[最小限のアクセス許可レベル](../../../outlook/understanding-outlook-add-in-permissions.md)|制限あり|
|[適用可能な Outlook のモード](../../../outlook/outlook-add-ins-overview.md#extension-points)|予定の開催者、予定の出席者、<br>メッセージの作成、またはメッセージの読み取り|

## <a name="properties"></a>プロパティ

| プロパティ | 最小値<br>アクセス許可レベル | 詳細モード | 戻り値の種類 | 最小値<br>要件セット |
|---|---|---|---|:---:|
| attachments | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.2#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| bcc | ReadItem | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.2#bcc) | [受信者](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| body | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.2#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.2#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.2#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| cc | ReadItem | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.2#cc) | [受信者](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.2#cc) | <[Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| conversationId | ReadItem | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.2#conversationid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.2#conversationid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeCreated | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2#datetimecreated) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.2#datetimecreated) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeModified | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2#datetimemodified) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.2#datetimemodified) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| end | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.2#end) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2#end) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.2#end)<br>(会議出席依頼) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| from | ReadItem | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.2#from) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| internetMessageId | ReadItem | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.2#internetmessageid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemClass | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2#itemclass) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.2#itemclass) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemId | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2#itemid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.2#itemid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemType | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.2#itemtype) | [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2#itemtype) | [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.2#itemtype) | [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.2#itemtype) | [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| 場所 | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.2#location) | [場所](/javascript/api/outlook/office.location) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2#location) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.2#location)<br>(会議出席依頼) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| normalizedSubject | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2#normalizedsubject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.2#normalizedsubject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| optionalAttendees | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.2#optionalattendees) | [受信者](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2#optionalattendees) | <[Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| organizer | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2#organizer) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| requiredAttendees | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.2#requiredattendees) | [受信者](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2#requiredattendees) | <[Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| 送信者 | ReadItem | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.2#sender) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| 開始 | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.2#start) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2#start) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.2#start)<br>(会議出席依頼) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| subject | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.2#subject) | [件名](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2#subject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.2#subject) | [件名](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.2#subject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| to | ReadItem | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.2#to) | [受信者](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.2#to) | <[Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>メソッド

| Method | 最小値<br>アクセス許可レベル | 詳細モード | 最小値<br>要件セット |
|---|---|---|:---:|
| addFileAttachmentAsync(uri, attachmentName, [options], [callback]) | ReadWriteItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.2#addfileattachmentasync-uri--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.2#addfileattachmentasync-uri--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| addItemAttachmentAsync(itemId, attachmentName, [options], [callback]) | ReadWriteItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.2#additemattachmentasync-itemid--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.2#additemattachmentasync-itemid--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyAllForm(formData) | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2#displayreplyallform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.2#displayreplyallform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyForm(formData) | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2#displayreplyform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.2#displayreplyform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntities () | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2#getentities--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.2#getentities--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntitiesByType (entityType) | Restricted | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2#getentitiesbytype-entitytype-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.2#getentitiesbytype-entitytype-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getFilteredEntitiesByName (名前) | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2#getfilteredentitiesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.2#getfilteredentitiesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatches () | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2#getregexmatches--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.2#getregexmatches--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatchesByName (名前) | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2#getregexmatchesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.2#getregexmatchesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getSelectedDataAsync (coercionType、[options]、callback) | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.2#getselecteddataasync-coerciontype--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.2#getselecteddataasync-coerciontype--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| loadCustomPropertiesAsync(callback, [userContext]) | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.2#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.2#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.2#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeAttachmentAsync(attachmentId, [options], [callback]) | ReadWriteItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.2#removeattachmentasync-attachmentid--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
|  |  | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.2#removeattachmentasync-attachmentid--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| setSelectedDataAsync(data, [options], callback) | ReadWriteItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.2#setselecteddataasync-data--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.2#setselecteddataasync-data--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |

## <a name="example"></a>例

次の JavaScript のコード例は、Outlook の現在のアイテムの `subject` プロパティにアクセスする方法を示しています。

```js
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
  });
};
```
