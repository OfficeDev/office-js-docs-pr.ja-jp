---
title: Outlook アドインでアイテム データを取得または設定する
description: アドインが閲覧フォームまたは新規作成フォームのどちらでアクティブ化されるかによって、アイテムでアドインが使用できるプロパティも異なります。
ms.date: 10/03/2022
ms.localizationpriority: medium
ms.openlocfilehash: 8349d81b376aa55d239a88a5d4598381fd8bfc4d
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467273"
---
# <a name="get-and-set-outlook-item-data-in-read-or-compose-forms"></a>閲覧または新規作成フォームの Outlook アイテム データを取得および設定する

Office アドイン マニフェスト スキーマのバージョン 1.1 以降、Outlook は、アイテムの表示または作成時にアドインをアクティブ化できます。 アドインが閲覧フォームまたは新規作成フォームのどちらでアクティブ化されるかによって、アイテムでアドインが使用できるプロパティも異なります。

たとえば、[dateTimeCreated](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) プロパティと [dateTimeModified](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) プロパティは送信済みのアイテム (アイテムは、その後閲覧フォームで表示されます) のみで定義され、(新規作成フォームで) メッセージの作成時にはこれらのプロパティは定義されません。 もう 1 つの例は [bcc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) プロパティです。このプロパティは、(新規作成フォームで) メッセージを作成する場合にのみ使用でき、閲覧フォームでは使用できません。

## <a name="item-properties-available-in-compose-and-read-forms"></a>新規作成フォームと閲覧フォームで使用できるアイテムのプロパティ

表 1 に、メール アドインの各モード (読み取りと作成) で使用できる Office JavaScript API のアイテム レベルのプロパティを示します。通常、読み取りフォームで使用できるプロパティは読み取り専用であり、作成フォームで使用できるプロパティは読み取り/書き込みであり、 [itemId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)、 [conversationId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)、 [および itemType](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) プロパティは例外であり、常に読み取り専用です。

新規作成フォームで使用可能な残りのアイテムレベルのプロパティは、アドインとユーザーが同時に同じプロパティの読み取りまたは書き込みを行う可能性があるため、新規作成モードでこれらのプロパティの取得や設定を行うメソッドは非同期です。このため、これらのプロパティが返すオブジェクトの種類も、新規作成フォームと閲覧フォームとで異なることがあります。 新規作成フォームで非同期のメソッドを使用してアイテムレベルのプロパティを取得または設定することについて詳しくは、「[Outlook で新規作成フォームのアイテム データを取得および設定する](get-and-set-item-data-in-a-compose-form.md)」をご覧ください。


**表 1. 新規作成フォームと閲覧フォームで使用できるアイテムのプロパティ**

<br/>

|**アイテムの種類**|**プロパティ**|**閲覧フォームにおけるプロパティのタイプ**|**新規作成フォームにおけるプロパティのタイプ**|
|:-----|:-----|:-----|:-----|
|予定とメッセージ|[dateTimeCreated](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|JavaScript **Date** オブジェクト|このプロパティは使用できません|
|予定とメッセージ|[dateTimeModified](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|JavaScript **Date** オブジェクト|このプロパティは使用できません|
|予定とメッセージ|[itemClass](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|String|このプロパティは使用できません|
|予定とメッセージ|[itemId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|String|このプロパティは使用できません|
|予定とメッセージ|[itemType](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) 列挙型の文字列|[ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) 列挙の文字列 (読み取り専用)|
|予定とメッセージ|[attachments](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)|このプロパティは使用できません|
|予定とメッセージ|[body](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[本文](/javascript/api/outlook/office.body)|[Body](/javascript/api/outlook/office.body)|
|予定とメッセージ|[normalizedSubject](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|String|このプロパティは使用できません|
|予定とメッセージ|[subject](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|String|[Subject](/javascript/api/outlook/office.subject)|
|予定|[end](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|JavaScript **Date** オブジェクト|[Time](/javascript/api/outlook/office.time)|
|予定|[location](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|String|[Location](/javascript/api/outlook/office.location)|
|予定|[optionalAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[受信者](/javascript/api/outlook/office.recipients)|
|予定|[organizer](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)|
|予定|[requiredAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[受信者](/javascript/api/outlook/office.recipients)|
|予定|[start](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|JavaScript **Date** オブジェクト|[Time](/javascript/api/outlook/office.time)|
|メッセージ|[bcc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|このプロパティは使用できません|[受信者](/javascript/api/outlook/office.recipients)|
|メッセージ|[cc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[受信者](/javascript/api/outlook/office.recipients)|
|メッセージ|[conversationId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|String|文字列 (読み取り専用)|
|メッセージ|[from](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)|
|メッセージ|[internetMessageId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|整数|このプロパティは使用できません|
|メッセージ|[sender](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|このプロパティは使用できません|
|メッセージ|[to](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[受信者](/javascript/api/outlook/office.recipients)|

## <a name="use-exchange-server-callback-tokens-from-a-read-add-in"></a>Exchange Server コールバック トークンを閲覧アドインから使用する

Outlook アドインが閲覧フォームでアクティブ化されると、Exchange コールバック トークンを取得できます。 このトークンをサーバー側のコードで使用して、Exchange Web Services (EWS) を介してすべてのアイテムにアクセスできます。

アドイン マニフェストで [アイテムの読み取りアクセス許可](understanding-outlook-add-in-permissions.md#read-item-permission) を指定することで、 [mailbox.getCallbackTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) メソッドを使用して Exchange コールバック トークンを取得し、 [mailbox.ewsUrl](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#properties) プロパティを使用してユーザーのメールボックスの EWS エンドポイントの URL を取得し、 [item.itemId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) を使用して選択したアイテムの EWS ID を取得できます。 その後、コールバック トークン、EWS エンドポイントの URL、EWS アイテム ID をサーバー側のコードに渡して [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) の操作にアクセスし、アイテムのその他のプロパティを取得することができます。

## <a name="access-ews-from-a-read-or-compose-add-in"></a>閲覧アドインまたは新規作成アドインから EWS にアクセスする

[mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) メソッドを使用すると、Exchange Web Services (EWS) の操作である [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) および [UpdateItem](/exchange/client-developer/web-service-reference/updateitem-operation) にアドインから直接アクセスすることもできます。 これらの操作を使用して、指定したアイテムの多数のプロパティを取得および設定できます。 このメソッドは、アドインマニフェストで読み取り **/書き込みメールボックス** のアクセス許可を指定する限り、読み取りフォームまたは作成フォームでアドインがアクティブ化されているかどうかに関係なく、Outlook アドインで使用できます。 **読み取り/書き込みメールボックス** のアクセス許可の詳細については、「[Outlook アドインのアクセス許可について](understanding-outlook-add-in-permissions.md)」を参照してください。

**makeEwsRequestAsync** を使用した EWS の操作へのアクセスの詳細については、「[Outlook アドインから Web サービスを呼び出す](web-services.md)」を参照してください。


## <a name="see-also"></a>関連項目

- [Outlook で新規作成フォームのアイテム データを取得および設定する](get-and-set-item-data-in-a-compose-form.md)
- [Outlook アドインから Web サービスを呼び出す](web-services.md)
