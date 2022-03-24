---
title: Office.context.mailbox.item - 要件セット 1.6
description: Outlookメールボックス API 要件セット 1.6 バージョンの Item オブジェクト モデル。
ms.date: 07/16/2021
ms.localizationpriority: medium
ms.openlocfilehash: 4ffd12da785f590f11efe566c3cd3f5a009c3ecd
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63747186"
---
# <a name="item-mailbox-requirement-set-16"></a>item (メールボックス要件セット 1.6)

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
| attachments | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-attachments-member) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-attachments-member) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| bcc | ReadItem | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6&preserve-view=true#outlook-office-messagecompose-bcc-member) | [受信者](/javascript/api/outlook/office.recipients?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| body | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentcompose-body-member) | [Body](/javascript/api/outlook/office.body?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-body-member) | [Body](/javascript/api/outlook/office.body?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6&preserve-view=true#outlook-office-messagecompose-body-member) | [Body](/javascript/api/outlook/office.body?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-body-member) | [Body](/javascript/api/outlook/office.body?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| cc | ReadItem | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6&preserve-view=true#outlook-office-messagecompose-cc-member) | [受信者](/javascript/api/outlook/office.recipients?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-cc-member) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| conversationId | ReadItem | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6&preserve-view=true#outlook-office-messagecompose-conversationid-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-conversationid-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeCreated | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-datetimecreated-member) | 日付 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-datetimecreated-member) | 日付 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeModified | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-datetimemodified-member) | 日付 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-datetimemodified-member) | 日付 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| end | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentcompose-end-member) | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-end-member) | 日付 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-end-member)<br>(会議出席依頼) | 日付 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| from | ReadItem | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-from-member) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| internetMessageId | ReadItem | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-internetmessageid-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemClass | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-itemclass-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-itemclass-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemId | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-itemid-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-itemid-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemType | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentcompose-itemtype-member) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-itemtype-member) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6&preserve-view=true#outlook-office-messagecompose-itemtype-member) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-itemtype-member) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| location | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentcompose-location-member) | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-location-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-location-member)<br>(会議出席依頼) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| normalizedSubject | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-normalizedsubject-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-normalizedsubject-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| notificationMessages | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentcompose-notificationmessages-member) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6&preserve-view=true) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-notificationmessages-member) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6&preserve-view=true) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6&preserve-view=true#outlook-office-messagecompose-notificationmessages-member) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6&preserve-view=true) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-notificationmessages-member) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6&preserve-view=true) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| optionalAttendees | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentcompose-optionalattendees-member) | [受信者](/javascript/api/outlook/office.recipients?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-optionalattendees-member) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| organizer | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-organizer-member) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| requiredAttendees | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentcompose-requiredattendees-member) | [受信者](/javascript/api/outlook/office.recipients?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-requiredattendees-member) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| sender | ReadItem | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-sender-member) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| 開始 | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentcompose-start-member) | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-start-member) | 日付 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-start-member)<br>(会議出席依頼) | 日付 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| subject | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentcompose-subject-member) | [件名](/javascript/api/outlook/office.subject?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-subject-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6&preserve-view=true#outlook-office-messagecompose-subject-member) | [件名](/javascript/api/outlook/office.subject?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-subject-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| へ | ReadItem | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6&preserve-view=true#outlook-office-messagecompose-to-member) | [受信者](/javascript/api/outlook/office.recipients?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-to-member) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>メソッド

| メソッド | 最小値<br>アクセス許可レベル | モード別の詳細 | 最小値<br>要件セット |
|---|---|---|:---:|
| addFileAttachmentAsync(uri, attachmentName, [options], [callback]) | ReadWriteItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentcompose-addfileattachmentasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6&preserve-view=true#outlook-office-messagecompose-addfileattachmentasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| addItemAttachmentAsync(itemId, attachmentName, [options], [callback]) | ReadWriteItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentcompose-additemattachmentasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6&preserve-view=true#outlook-office-messagecompose-additemattachmentasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| close() | Restricted | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentcompose-close-member(1)) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6&preserve-view=true#outlook-office-messagecompose-close-member(1)) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| displayReplyAllForm(formData) | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-displayreplyallform-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-displayreplyallform-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyForm(formData) | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-displayreplyform-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-displayreplyform-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntities() | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-getentities-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-getentities-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntitiesByType(entityType) | Restricted | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-getentitiesbytype-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-getentitiesbytype-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getFilteredEntitiesByName(name) | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-getfilteredentitiesbyname-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-getfilteredentitiesbyname-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatches() | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-getregexmatches-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-getregexmatches-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatchesByName(name) | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-getregexmatchesbyname-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-getregexmatchesbyname-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getSelectedDataAsync(coercionType, [options], callback) | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentcompose-getselecteddataasync-member(1)) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6&preserve-view=true#outlook-office-messagecompose-getselecteddataasync-member(1)) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| getSelectedEntities() | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-getselectedentities-member(1)) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-getselectedentities-member(1)) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| getSelectedRegExMatches() | ReadItem | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-getselectedregexmatches-member(1)) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-getselectedregexmatches-member(1)) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| loadCustomPropertiesAsync(callback, [userContext]) | ReadItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentcompose-loadcustompropertiesasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [予定の出席者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-loadcustompropertiesasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6&preserve-view=true#outlook-office-messagecompose-loadcustompropertiesasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [既読メッセージ](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-loadcustompropertiesasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeAttachmentAsync(attachmentId, [options], [callback]) | ReadWriteItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentcompose-removeattachmentasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
|  |  | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6&preserve-view=true#outlook-office-messagecompose-removeattachmentasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| saveAsync([options], callback) | ReadWriteItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentcompose-saveasync-member(1)) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6&preserve-view=true#outlook-office-messagecompose-saveasync-member(1)) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| setSelectedDataAsync(data, [options], callback) | ReadWriteItem | [予定の開催者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentcompose-setselecteddataasync-member(1)) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [メッセージ作成](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6&preserve-view=true#outlook-office-messagecompose-setselecteddataasync-member(1)) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |

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
