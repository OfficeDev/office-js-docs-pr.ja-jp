---
title: Office. アイテム-プレビュー要件セット
description: Outlook メールボックス API のプレビュー要件セットのバージョンのアイテムオブジェクトモデル。
ms.date: 03/27/2020
localization_priority: Normal
ms.openlocfilehash: c4c605b4d2a49e8fbd9dd9de1d7293738c5f1505
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612032"
---
# <a name="item-mailbox-preview-requirement-set"></a>アイテム (メールボックスプレビュー要件セット)

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
| attachments | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-preview#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| bcc | ReadItem | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#bcc) | [受信者](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| body | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-preview#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| categories | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#categories) | [カテゴリ](/javascript/api/outlook/office.categories) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#categories) | [カテゴリ](/javascript/api/outlook/office.categories) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#categories) | [カテゴリ](/javascript/api/outlook/office.categories) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-preview#categories) | [カテゴリ](/javascript/api/outlook/office.categories) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| cc | ReadItem | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#cc) | [受信者](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-preview#cc) | <[Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| conversationId | ReadItem | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#conversationid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-preview#conversationid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeCreated | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#datetimecreated) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-preview#datetimecreated) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeModified | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#datetimemodified) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-preview#datetimemodified) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| end | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#end) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#end) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-preview#end)<br>(会議出席依頼) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| enhancedLocation | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#enhancedlocation) | [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#enhancedlocation) | [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| from | ReadWriteItem | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#from) | [From](/javascript/api/outlook/office.from) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | ReadItem | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-preview#from) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| internetHeaders | ReadItem | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#internetheaders) | [InternetHeaders](/javascript/api/outlook/office.internetheaders) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| internetMessageId | ReadItem | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-preview#internetmessageid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| isAllDayEvent | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#isalldayevent) | [IsAllDayEvent](/javascript/api/outlook/office.isalldayevent) | [Preview](outlook-requirement-set-preview.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#isalldayevent) | ブール型 | [Preview](outlook-requirement-set-preview.md) |
| itemClass | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#itemclass) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-preview#itemclass) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemId | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#itemid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-preview#itemid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemType | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#itemtype) | [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#itemtype) | [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#itemtype) | [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-preview#itemtype) | [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| 場所 | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#location) | [場所](/javascript/api/outlook/office.location) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#location) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-preview#location)<br>(会議出席依頼) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| normalizedSubject | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#normalizedsubject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-preview#normalizedsubject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| notificationMessages | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-preview#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| optionalAttendees | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#optionalattendees) | [受信者](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#optionalattendees) | <[Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| organizer | ReadWriteItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#organizer) | [Organizer](/javascript/api/outlook/office.organizer) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#organizer) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| recurrence | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#recurrence) | [繰り返さ](/javascript/api/outlook/office.recurrence) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#recurrence) | [繰り返さ](/javascript/api/outlook/office.recurrence) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-preview#recurrence)<br>(会議出席依頼) | [繰り返さ](/javascript/api/outlook/office.recurrence) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| requiredAttendees | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#requiredattendees) | [受信者](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#requiredattendees) | <[Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| 送信者 | ReadItem | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-preview#sender) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| sensitivity | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#sensitivity) | [Sensitivity](/javascript/api/outlook/office.sensitivity) | [Preview](outlook-requirement-set-preview.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#sensitivity) | [MailboxEnums AppointmentSensitivityType](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype) | [Preview](outlook-requirement-set-preview.md) |
| 系列 Id | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#seriesid) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#seriesid) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#seriesid) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-preview#seriesid) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| 開始 | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#start) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#start) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-preview#start)<br>(会議出席依頼) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| subject | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#subject) | [件名](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#subject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#subject) | [件名](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-preview#subject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| to | ReadItem | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#to) | [受信者](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-preview#to) | <[Emailaddressdetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>メソッド

| Method | 最小値<br>アクセス許可レベル | 詳細モード | 最小値<br>要件セット |
|---|---|---|:---:|
| addFileAttachmentAsync(uri, attachmentName, [options], [callback]) | ReadWriteItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#addfileattachmentasync-uri--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#addfileattachmentasync-uri--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| addFileAttachmentFromBase64Async (base64File, attachmentName, [options], [callback]) | ReadWriteItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#addfileattachmentfrombase64async-base64file--attachmentname--options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#addfileattachmentfrombase64async-base64file--attachmentname--options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| addHandlerAsync(eventType, handler, [options], [callback]) | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#addhandlerasync-eventtype--handler--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#addhandlerasync-eventtype--handler--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#addhandlerasync-eventtype--handler--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-preview#addhandlerasync-eventtype--handler--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| addItemAttachmentAsync(itemId, attachmentName, [options], [callback]) | ReadWriteItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#additemattachmentasync-itemid--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#additemattachmentasync-itemid--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| close() | Restricted | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#close--) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#close--) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| disableClientSignatureAsync ([オプション], [callback]) | ReadWriteItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#disableclientsignatureasync-options--callback-) | [Preview](outlook-requirement-set-preview.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#disableclientsignatureasync-options--callback-) | [Preview](outlook-requirement-set-preview.md) |
| displayReplyAllForm(formData) | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#displayreplyallform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-preview#displayreplyallform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyForm(formData) | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#displayreplyform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-preview#displayreplyform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getAllInternetHeadersAsync ([オプション], [callback]) | ReadItem | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getallinternetheadersasync-options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getAttachmentContentAsync (attachmentId, [options], [callback]) | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#getattachmentcontentasync-attachmentid--options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getattachmentcontentasync-attachmentid--options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getattachmentcontentasync-attachmentid--options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getattachmentcontentasync-attachmentid--options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getAttachmentsAsync ([オプション], [callback]) | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#getattachmentsasync-options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getattachmentsasync-options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getComposeTypeAsync ([オプション], callback) | ReadItem | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getcomposetypeasync-options--callback-) | [Preview](outlook-requirement-set-preview.md) |
| getEntities () | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getentities--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getentities--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntitiesByType (entityType) | Restricted | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getentitiesbytype-entitytype-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getentitiesbytype-entitytype-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getFilteredEntitiesByName (名前) | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getfilteredentitiesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getfilteredentitiesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| 、Office.context.mailbox.item.getinitializationcontextasync ([オプション], [callback]) | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#getinitializationcontextasync-options--callback-) | [Preview](../preview-requirement-set/outlook-requirement-set-preview.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getinitializationcontextasync-options--callback-) | [Preview](../preview-requirement-set/outlook-requirement-set-preview.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getinitializationcontextasync-options--callback-) | [Preview](../preview-requirement-set/outlook-requirement-set-preview.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getinitializationcontextasync-options--callback-) | [Preview](../preview-requirement-set/outlook-requirement-set-preview.md) |
| getItemIdAsync ([オプション], callback) | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#getitemidasync-options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getitemidasync-options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getRegExMatches () | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getregexmatches--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getregexmatches--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatchesByName (名前) | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getregexmatchesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getregexmatchesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getSelectedDataAsync (coercionType、[options]、callback) | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#getselecteddataasync-coerciontype--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getselecteddataasync-coerciontype--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| Office.context.mailbox.item.getselectedentities () | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getselectedentities--) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getselectedentities--) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| Office.context.mailbox.item.getselectedregexmatches () | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getselectedregexmatches--) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getselectedregexmatches--) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| getSharedPropertiesAsync ([options], callback) | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#getsharedpropertiesasync-options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getsharedpropertiesasync-options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getsharedpropertiesasync-options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getsharedpropertiesasync-options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| isClientSignatureEnabledAsync ([オプション], callback) | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#isclientsignatureenabledasync-options--callback-) | [Preview](outlook-requirement-set-preview.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#isclientsignatureenabledasync-options--callback-) | [Preview](outlook-requirement-set-preview.md) |
| loadCustomPropertiesAsync(callback, [userContext]) | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-preview#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeAttachmentAsync(attachmentId, [options], [callback]) | ReadWriteItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#removeattachmentasync-attachmentid--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
|  |  | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#removeattachmentasync-attachmentid--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeHandlerAsync(eventType, [options], [callback]) | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#removehandlerasync-eventtype--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#removehandlerasync-eventtype--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#removehandlerasync-eventtype--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-preview#removehandlerasync-eventtype--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| saveAsync([options], callback) | ReadWriteItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#saveasync-options--callback-) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#saveasync-options--callback-) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| setSelectedDataAsync(data, [options], callback) | ReadWriteItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#setselecteddataasync-data--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#setselecteddataasync-data--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |

## <a name="events"></a>イベント

およびを使用して、以下のイベントにサブスクライブし、サブスクライブを解除することができ `addHandlerAsync` `removeHandlerAsync` ます。

| イベント | 説明 | 最小値<br>要件セット |
|---|---|:---:|
|`AppointmentTimeChanged`| 選択した予定またはデータ系列の日付または時刻が変更されました。 | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
|`AttachmentsChanged`| 添付ファイルがアイテムに追加またはアイテムから削除されています。 | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
|`EnhancedLocationsChanged`| 選択した予定の場所が変更されました。 | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
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
