---
title: Office.context.mailbox.item - 要件セット 1.9
description: Outlookメールボックス API 要件セット 1.9 バージョンの Item オブジェクト モデル。
ms.date: 07/16/2021
localization_priority: Normal
ms.openlocfilehash: f2cdbedc91906ec69363d8f61374c67d498fecb068db6d49c8401d53d8f01101
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57087689"
---
# <a name="item-mailbox-requirement-set-19"></a>item (メールボックス要件セット 1.9)

### <a name="officecontextmailboxitem"></a>[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item

`item` は、現在選択されているメッセージ、会議出席依頼、または予定にアクセスするために使用されます。 プロパティを使用して、アイテムの種類を確認 `itemType` できます。

##### <a name="requirements"></a>要件

|要件|値|
|---|---|
|[メールボックスの最小要件セットのバージョン](../../requirement-sets/outlook-api-requirement-sets.md)|1.1|
|[最小限のアクセス許可レベル](../../../outlook/understanding-outlook-add-in-permissions.md)|制限あり|
|[適用可能な Outlook のモード](../../../outlook/outlook-add-ins-overview.md#extension-points)|予定の開催者、予定の出席者、<br>メッセージの作成、またはメッセージの読み取り|

> [!IMPORTANT]
> Android と iOS: アドインがアクティブ化される時間と使用可能な API には制限があります。 詳細については、「[Outlook アドインにモバイル サポートを追加する](../../../outlook/add-mobile-support.md#compose-mode-and-appointments)」を参照してください。

## <a name="properties"></a>プロパティ

| プロパティ | 最小値<br>アクセス許可レベル | モード別の詳細 | 戻り値の種類 | 最小値<br>要件セット |
|---|---|---|---|:---:|
| attachments | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| bcc | ReadItem | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#bcc) | [受信者](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| body | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| categories | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#categories) | [Categories](/javascript/api/outlook/office.categories) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#categories) | [Categories](/javascript/api/outlook/office.categories) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#categories) | [Categories](/javascript/api/outlook/office.categories) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#categories) | [Categories](/javascript/api/outlook/office.categories) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| cc | ReadItem | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#cc) | [受信者](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#cc) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| conversationId | ReadItem | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#conversationId) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#conversationId) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeCreated | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#dateTimeCreated) | 日付 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#dateTimeCreated) | 日付 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeModified | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#dateTimeModified) | 日付 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#dateTimeModified) | 日付 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| end | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#end) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#end) | 日付 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#end)<br>(会議出席依頼) | 日付 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| enhancedLocation | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#enhancedLocation) | [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#enhancedLocation) | [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| from | ReadWriteItem | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#from) | [From](/javascript/api/outlook/office.from) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | ReadItem | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#from) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| internetHeaders | ReadItem | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#internetHeaders) | [InternetHeaders](/javascript/api/outlook/office.internetheaders) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| internetMessageId | ReadItem | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#internetMessageId) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemClass | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#itemClass) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#itemClass) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemId | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#itemId) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#itemId) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemType | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#itemType) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#itemType) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#itemType) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#itemType) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| 場所 | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#location) | [場所](/javascript/api/outlook/office.location) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#location) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#location)<br>(会議出席依頼) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| normalizedSubject | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#normalizedSubject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#normalizedSubject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| notificationMessages | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#notificationMessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#notificationMessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#notificationMessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#notificationMessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| optionalAttendees | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#optionalAttendees) | [受信者](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#optionalAttendees) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| 構成内容変更 | ReadWriteItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#organizer) | [Organizer](/javascript/api/outlook/office.organizer) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#organizer) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| recurrence | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#recurrence) | [Recurrence](/javascript/api/outlook/office.recurrence) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#recurrence) | [Recurrence](/javascript/api/outlook/office.recurrence) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#recurrence)<br>(会議出席依頼) | [Recurrence](/javascript/api/outlook/office.recurrence) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| requiredAttendees | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#requiredAttendees) | [受信者](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#requiredAttendees) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| sender | ReadItem | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#sender) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| seriesId | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#seriesId) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#seriesId) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#seriesId) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#seriesId) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| 開始 | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#start) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#start) | 日付 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#start)<br>(会議出席依頼) | 日付 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| subject | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#subject) | [[件名]](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#subject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#subject) | [[件名]](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#subject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| へ | ReadItem | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#to) | [受信者](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#to) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>メソッド

| メソッド | 最小値<br>アクセス許可レベル | モード別の詳細 | 最小値<br>要件セット |
|---|---|---|:---:|
| addFileAttachmentAsync(uri, attachmentName, [options], [callback]) | ReadWriteItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#addFileAttachmentAsync_uri__attachmentName__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#addFileAttachmentAsync_uri__attachmentName__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback]) | ReadWriteItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#addFileAttachmentFromBase64Async_base64File__attachmentName__options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#addFileAttachmentFromBase64Async_base64File__attachmentName__options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| addHandlerAsync(eventType, handler, [options], [callback]) | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#addHandlerAsync_eventType__handler__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#addHandlerAsync_eventType__handler__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#addHandlerAsync_eventType__handler__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#addHandlerAsync_eventType__handler__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| addItemAttachmentAsync(itemId, attachmentName, [options], [callback]) | ReadWriteItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#addItemAttachmentAsync_itemId__attachmentName__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#addItemAttachmentAsync_itemId__attachmentName__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| close() | Restricted | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#close__) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#close__) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| displayReplyAllForm(formData) | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#displayReplyAllForm_formData_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#displayReplyAllForm_formData_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyAllFormAsync(formData, [options], [callback]) | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#displayReplyAllFormAsync_formData__options__callback_) | [1.9](outlook-requirement-set-1.9.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#displayReplyAllFormAsync_formData__options__callback_) | [1.9](outlook-requirement-set-1.9.md) |
| displayReplyForm(formData) | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#displayReplyForm_formData_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#displayReplyForm_formData_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyFormAsync(formData, [options], [callback]) | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#displayReplyFormAsync_formData__options__callback_) | [1.9](outlook-requirement-set-1.9.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#displayReplyFormAsync_formData__options__callback_) | [1.9](outlook-requirement-set-1.9.md) |
| getAllInternetHeadersAsync([options], [callback]) | ReadItem | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#getAllInternetHeadersAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getAttachmentContentAsync(attachmentId, [options], [callback]) | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#getAttachmentContentAsync_attachmentId__options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#getAttachmentContentAsync_attachmentId__options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#getAttachmentContentAsync_attachmentId__options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#getAttachmentContentAsync_attachmentId__options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getAttachmentsAsync([options], [callback]) | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#getAttachmentsAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#getAttachmentsAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getEntities() | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#getEntities__) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#getEntities__) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntitiesByType(entityType) | Restricted | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#getEntitiesByType_entityType_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#getEntitiesByType_entityType_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getFilteredEntitiesByName(name) | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#getFilteredEntitiesByName_name_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#getFilteredEntitiesByName_name_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getItemIdAsync([options], callback) | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#getItemIdAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#getItemIdAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getRegExMatches() | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#getRegExMatches__) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#getRegExMatches__) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatchesByName(name) | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#getRegExMatchesByName_name_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#getRegExMatchesByName_name_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getSelectedDataAsync(coercionType, [options], callback) | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#getSelectedDataAsync_coercionType__options__callback_) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#getSelectedDataAsync_coercionType__options__callback_) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| getSelectedEntities() | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#getSelectedEntities__) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#getSelectedEntities__) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| getSelectedRegExMatches() | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#getSelectedRegExMatches__) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#getSelectedRegExMatches__) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| getSharedPropertiesAsync([options], callback) | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#getSharedPropertiesAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#getSharedPropertiesAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#getSharedPropertiesAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#getSharedPropertiesAsync_options__callback_) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| loadCustomPropertiesAsync(callback, [userContext]) | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#loadCustomPropertiesAsync_callback__userContext_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#loadCustomPropertiesAsync_callback__userContext_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#loadCustomPropertiesAsync_callback__userContext_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#loadCustomPropertiesAsync_callback__userContext_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeAttachmentAsync(attachmentId, [options], [callback]) | ReadWriteItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#removeAttachmentAsync_attachmentId__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
|  |  | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#removeAttachmentAsync_attachmentId__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeHandlerAsync(eventType, [options], [callback]) | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#removeAttachmentAsync_attachmentId__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.9&preserve-view=true#removeAttachmentAsync_attachmentId__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#removeAttachmentAsync_attachmentId__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.9&preserve-view=true#removeAttachmentAsync_attachmentId__options__callback_) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| saveAsync([options], callback) | ReadWriteItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#saveAsync_options__callback_) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#saveAsync_options__callback_) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| setSelectedDataAsync(data, [options], callback) | ReadWriteItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.9&preserve-view=true#setSelectedDataAsync_data__options__callback_) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.9&preserve-view=true#setSelectedDataAsync_data__options__callback_) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |

## <a name="events"></a>イベント

それぞれを使用して、以下のイベントを購読および購読 `addHandlerAsync` 解除 `removeHandlerAsync` できます。

> [!IMPORTANT]
> イベントは、作業ウィンドウの実装でのみ使用できます。

| [Event](/javascript/api/office/office.eventtype) | 説明 | 最小値<br>要件セット |
|---|---|:---:|
|`AppointmentTimeChanged`| 選択した予定または系列の日付または時刻が変更されました。 | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
|`AttachmentsChanged`| アイテムに添付ファイルが追加または削除されました。 | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
|`EnhancedLocationsChanged`| 選択した予定の場所が変更されました。 | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
|`RecipientsChanged`| 選択したアイテムまたは予定の場所の受信者リストが変更されました。 | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
|`RecurrenceChanged`| 選択した系列の定期的なパターンが変更されました。 | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |

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
