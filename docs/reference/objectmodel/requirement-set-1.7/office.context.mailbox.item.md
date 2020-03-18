---
title: Office. メールボックス-要件セット1.7
description: Office のオブジェクトモデルです。アイテムのモデル (要件セット 1.7)
ms.date: 03/16/2020
localization_priority: Normal
ms.openlocfilehash: 403012df4ffcb3350f0b0a6f64add39c9d377f9a
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717622"
---
# <a name="item"></a>item

### <a name="officecontextmailboxitem"></a>[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item

`item`は、現在選択されているメッセージ、会議出席依頼、または予定にアクセスするために使用されます。 `itemType`プロパティを使用して、アイテムの種類を調べることができます。

##### <a name="requirements"></a>Requirements

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../../requirement-sets/outlook-api-requirement-sets.md)|1.1|
|[最小限のアクセス許可レベル](../../../outlook/understanding-outlook-add-in-permissions.md)|制限あり|
|[適用可能な Outlook のモード](../../../outlook/outlook-add-ins-overview.md#extension-points)|予定の開催者、予定の出席者、<br>メッセージの作成、またはメッセージの読み取り|

## <a name="properties"></a>プロパティ

| プロパティ | 最小値<br>アクセス許可レベル | 詳細モード | 戻り値の種類 | 最小値<br>要件セット |
|---|---|---|---|:---:|
| attachments | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージの読み取り](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| Bcc | ReadItem | [メッセージの作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#bcc) | [受信者](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| body | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージの作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージの読み取り](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| Cc | ReadItem | [メッセージの作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#cc) | [受信者](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージの読み取り](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#cc) | <[Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| conversationId | ReadItem | [メッセージの作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#conversationid) | 文字列 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージの読み取り](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#conversationid) | 文字列 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeCreated | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#datetimecreated) | 日付 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージの読み取り](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#datetimecreated) | 日付 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeModified | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#datetimemodified) | 日付 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージの読み取り](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#datetimemodified) | 日付 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| end | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#end) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#end) | 日付 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージの読み取り](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#end)<br>(会議出席依頼) | 日付 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| from | ReadWriteItem | [メッセージの作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#from) | [From](/javascript/api/outlook/office.from) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | ReadItem | [メッセージの読み取り](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#from) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| internetMessageId | ReadItem | [メッセージの読み取り](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#internetmessageid) | 文字列 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemClass | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#itemclass) | 文字列 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージの読み取り](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#itemclass) | 文字列 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemId | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#itemid) | 文字列 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージの読み取り](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#itemid) | 文字列 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemType | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#itemtype) | [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#itemtype) | [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージの作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#itemtype) | [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージの読み取り](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#itemtype) | [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| location | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#location) | [Location](/javascript/api/outlook/office.location) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#location) | 文字列 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージの読み取り](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#location)<br>(会議出席依頼) | 文字列 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| normalizedSubject | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#normalizedsubject) | 文字列 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージの読み取り](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#normalizedsubject) | 文字列 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| notificationMessages | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [メッセージの作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [メッセージの読み取り](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| optionalAttendees | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#optionalattendees) | [受信者](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#optionalattendees) | <[Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| organizer | ReadWriteItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#organizer) | [Organizer](/javascript/api/outlook/office.organizer) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#organizer) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| recurrence | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#recurrence) | [繰り返さ](/javascript/api/outlook/office.recurrence) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#recurrence) | [繰り返さ](/javascript/api/outlook/office.recurrence) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [メッセージの読み取り](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#recurrence)<br>(会議出席依頼) | [繰り返さ](/javascript/api/outlook/office.recurrence) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| requiredAttendees | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#requiredattendees) | [受信者](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#requiredattendees) | <[Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| sender | ReadItem | [メッセージの読み取り](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#sender) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| 系列 Id | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#seriesid) | 文字列 | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#seriesid) | 文字列 | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [メッセージの作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#seriesid) | 文字列 | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [メッセージの読み取り](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#seriesid) | 文字列 | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| 開始 | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#start) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#start) | 日付 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージの読み取り](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#start)<br>(会議出席依頼) | 日付 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| subject | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#subject) | [Subject](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#subject) | 文字列 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージの作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#subject) | [Subject](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージの読み取り](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#subject) | 文字列 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| 宛先 | ReadItem | [メッセージの作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#to) | [受信者](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージの読み取り](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#to) | <[Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>メソッド

| メソッド | 最小値<br>アクセス許可レベル | 詳細モード | 最小値<br>要件セット |
|---|---|---|:---:|
| addFileAttachmentAsync(uri, attachmentName, [options], [callback]) | ReadWriteItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#addfileattachmentasync-uri--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージの作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#addfileattachmentasync-uri--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| addHandlerAsync(eventType, handler, [options], [callback]) | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#addhandlerasync-eventtype--handler--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#addhandlerasync-eventtype--handler--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [メッセージの作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#addhandlerasync-eventtype--handler--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [メッセージの読み取り](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#addhandlerasync-eventtype--handler--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| addItemAttachmentAsync(itemId, attachmentName, [options], [callback]) | ReadWriteItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#additemattachmentasync-itemid--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージの作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#additemattachmentasync-itemid--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| close() | Restricted | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#close--) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [メッセージの作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#close--) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| displayReplyAllForm(formData, [callback]) | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#displayreplyallform-formdata--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージの読み取り](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#displayreplyallform-formdata--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyForm(formData, [callback]) | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#displayreplyform-formdata--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージの読み取り](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#displayreplyform-formdata--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntities () | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#getentities--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージの読み取り](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#getentities--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntitiesByType (entityType) | Restricted | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#getentitiesbytype-entitytype-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージの読み取り](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#getentitiesbytype-entitytype-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getFilteredEntitiesByName (名前) | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#getfilteredentitiesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージの読み取り](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#getfilteredentitiesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatches () | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#getregexmatches--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージの読み取り](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#getregexmatches--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatchesByName (名前) | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#getregexmatchesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージの読み取り](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#getregexmatchesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getSelectedDataAsync (coercionType、[options]、callback) | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#getselecteddataasync-coerciontype--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [メッセージの作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#getselecteddataasync-coerciontype--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| Office.context.mailbox.item.getselectedentities () | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#getselectedentities--) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [メッセージの読み取り](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#getselectedentities--) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| Office.context.mailbox.item.getselectedregexmatches () | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#getselectedregexmatches--) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [メッセージの読み取り](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#getselectedregexmatches--) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| loadCustomPropertiesAsync(callback, [userContext]) | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージの作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージの読み取り](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeAttachmentAsync(attachmentId, [options], [callback]) | ReadWriteItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#removeattachmentasync-attachmentid--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
|  |  | [メッセージの作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#removeattachmentasync-attachmentid--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeHandlerAsync(eventType, [options], [callback]) | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#removehandlerasync-eventtype--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7#removehandlerasync-eventtype--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [メッセージの作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#removehandlerasync-eventtype--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [メッセージの読み取り](/javascript/api/outlook/office.messageread?view=outlook-js-1.7#removehandlerasync-eventtype--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| saveAsync([options], callback) | ReadWriteItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#saveasync-options--callback-) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [メッセージの作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#saveasync-options--callback-) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| setSelectedDataAsync(data, [options], callback) | ReadWriteItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7#setselecteddataasync-data--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [メッセージの作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7#setselecteddataasync-data--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |

## <a name="events"></a>イベント

および`addHandlerAsync` `removeHandlerAsync`を使用して、以下のイベントにサブスクライブし、サブスクライブを解除することができます。

| イベント | 説明 | 最小値<br>要件セット |
|---|---|:---:|
|`AppointmentTimeChanged`| 選択した予定またはデータ系列の日付または時刻が変更されました。 | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
|`RecipientsChanged`| 選択したアイテムまたは予定の場所の受信者の一覧が変更されました。 | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
|`RecurrenceChanged`| 選択したアイテムの定期的なパターンが変更されました。 | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |

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
