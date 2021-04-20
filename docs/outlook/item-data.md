---
title: Outlook アドインでアイテム データを取得または設定する
description: アドインが閲覧フォームまたは新規作成フォームのどちらでアクティブ化されるかによって、アイテムでアドインが使用できるプロパティも異なります。
ms.date: 12/10/2019
localization_priority: Normal
ms.openlocfilehash: 6ef0b868ba83a6003c401e15b7d6dd9ee0f4f94b
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609077"
---
# <a name="get-and-set-outlook-item-data-in-read-or-compose-forms"></a>閲覧または新規作成フォームの Outlook アイテム データを取得および設定する

Office アドイン マニフェスト スキーマのバージョン 1.1 以降、Outlook は、アイテムの表示または作成時にアドインをアクティブ化できます。 アドインが閲覧フォームまたは新規作成フォームのどちらでアクティブ化されるかによって、アイテムでアドインが使用できるプロパティも異なります。

たとえば、[dateTimeCreated](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) プロパティと [dateTimeModified](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) プロパティは送信済みのアイテム (アイテムは、その後閲覧フォームで表示されます) のみで定義され、(新規作成フォームで) メッセージの作成時にはこれらのプロパティは定義されません。 もう 1 つの例は [bcc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) プロパティです。このプロパティは、(新規作成フォームで) メッセージを作成する場合にのみ使用でき、閲覧フォームでは使用できません。

## <a name="item-properties-available-in-compose-and-read-forms"></a>新規作成フォームと閲覧フォームで使用できるアイテムのプロパティ

表1は、メールアドインの各モード (読み取りおよび作成) で利用できる Office JavaScript API のアイテムレベルのプロパティを示しています。通常、閲覧フォームで使用可能なプロパティは読み取り専用であり、新規作成フォームで使用できるプロパティは読み取り/書き込み可能です。 [itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)、 [ConversationId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)、および[itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)プロパティは例外です。ただし、読み取り専用になります。

新規作成フォームで使用可能な残りのアイテムレベルのプロパティは、アドインとユーザーが同時に同じプロパティの読み取りまたは書き込みを行う可能性があるため、新規作成モードでこれらのプロパティの取得や設定を行うメソッドは非同期です。このため、これらのプロパティが返すオブジェクトの種類も、新規作成フォームと閲覧フォームとで異なることがあります。 新規作成フォームで非同期のメソッドを使用してアイテムレベルのプロパティを取得または設定することについて詳しくは、「[Outlook で新規作成フォームのアイテム データを取得および設定する](get-and-set-item-data-in-a-compose-form.md)」をご覧ください。


**表 1. 新規作成フォームと閲覧フォームで使用できるアイテムのプロパティ**

<br/>

|**アイテムの種類**|**プロパティ**|**閲覧フォームにおけるプロパティのタイプ**|**新規作成フォームにおけるプロパティのタイプ**|
|:-----|:-----|:-----|:-----|
|予定とメッセージ|[dateTimeCreated](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|JavaScript **Date** オブジェクト|このプロパティは使用できません|
|予定とメッセージ|[dateTimeModified](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|JavaScript **Date** オブジェクト|このプロパティは使用できません|
|予定とメッセージ|[itemClass](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|String|このプロパティは使用できません|
|予定とメッセージ|[itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|String|このプロパティは使用できません|
|予定とメッセージ|[itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) 列挙型の文字列|[ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) 列挙型の文字列 (読み取り専用)|
|予定とメッセージ|[attachments](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)|このプロパティは使用できません|
|予定とメッセージ|[body](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[本文](/javascript/api/outlook/office.body)|[Body](/javascript/api/outlook/office.body)|
|予定とメッセージ|[normalizedSubject](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|String|このプロパティは使用できません|
|予定とメッセージ|[subject](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|String|[Subject](/javascript/api/outlook/office.subject)|
|予定|[end](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|JavaScript **Date** オブジェクト|[Time](/javascript/api/outlook/office.time)|
|予定|[location](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|String|[Location](/javascript/api/outlook/office.location)|
|予定|[optionalAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[受信者](/javascript/api/outlook/office.recipients)|
|予定|[organizer](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)|
|予定|[requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[受信者](/javascript/api/outlook/office.recipients)|
|予定|[start](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|JavaScript **Date** オブジェクト|[Time](/javascript/api/outlook/office.time)|
|メッセージ|[bcc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|このプロパティは使用できません|[受信者](/javascript/api/outlook/office.recipients)|
|メッセージ|[cc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[受信者](/javascript/api/outlook/office.recipients)|
|メッセージ|[conversationId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|String|文字列 (読み取り専用)|
|メッセージ|[from](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)|
|メッセージ|[internetMessageId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|整数|このプロパティは使用できません|
|メッセージ|[sender](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|このプロパティは使用できません|
|メッセージ|[to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[受信者](/javascript/api/outlook/office.recipients)|

## <a name="use-exchange-server-callback-tokens-from-a-read-add-in"></a>Exchange Server コールバック トークンを閲覧アドインから使用する

Outlook アドインが閲覧フォームでアクティブ化されると、Exchange コールバック トークンを取得できます。 このトークンをサーバー側のコードで使用して、Exchange Web Services (EWS) を介してすべてのアイテムにアクセスできます。

アドイン マニフェストで **ReadItem** のアクセス許可を指定すると、[mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) メソッドを使用した Exchange コールバック トークンの取得、[mailbox.ewsUrl](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties) プロパティを使用したユーザーのメールボックスの EWS エンドポイントの URL の取得、[item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) による選択したアイテムの EWS ID の取得が可能です。 その後、コールバック トークン、EWS エンドポイントの URL、EWS アイテム ID をサーバー側のコードに渡して [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) の操作にアクセスし、アイテムのその他のプロパティを取得することができます。


## <a name="access-ews-from-a-read-or-compose-add-in"></a>閲覧アドインまたは新規作成アドインから EWS にアクセスする

[mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) メソッドを使用すると、Exchange Web Services (EWS) の操作である [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) および [UpdateItem](/exchange/client-developer/web-service-reference/updateitem-operation) にアドインから直接アクセスすることもできます。 これらの操作を使用して、指定したアイテムの多数のプロパティを取得および設定できます。 このメソッドは、アドイン マニフェストで **ReadWriteMailbox** のアクセス許可が指定されている限り、アドインが閲覧フォームまたは新規作成フォームのどちらでアクティブ化されたかに関係なく、Outlook アドインで使用できます。

**makeEwsRequestAsync** を使用した EWS の操作へのアクセスの詳細については、「[Outlook アドインから Web サービスを呼び出す](web-services.md)」を参照してください。


## <a name="see-also"></a>関連項目

- [Outlook で新規作成フォームのアイテム データを取得および設定する](get-and-set-item-data-in-a-compose-form.md)
- [Outlook アドインから Web サービスを呼び出す](web-services.md)
