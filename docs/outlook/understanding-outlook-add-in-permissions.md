---
title: Outlook アドインのアクセス許可を理解する
description: Outlook アドインでは、必要なアクセス許可のレベルをマニフェストで指定します。使用可能なレベルは Restricted、ReadItem、ReadWriteItem、ReadWriteMailbox です。
ms.date: 02/19/2020
ms.localizationpriority: medium
ms.openlocfilehash: 6350e0d3aed499d831c13e440945fda1f60742ca
ms.sourcegitcommit: b66ba72aee8ccb2916cd6012e66316df2130f640
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484185"
---
# <a name="understanding-outlook-add-in-permissions"></a>Outlook アドインのアクセス許可を理解する

Outlook アドインでは、必要なアクセス許可のレベルをマニフェストで指定します。使用可能なレベルは **Restricted**、**ReadItem**、**ReadWriteItem**、**ReadWriteMailbox** です。これらのレベルのアクセス許可は累積されます。**Restricted** は最低レベルであり、それぞれの上位レベルには、下位レベルのアクセス許可がすべて含まれます。**ReadWriteMailbox** にはサポートされるアクセス許可がすべて含まれます。

メール アドインが要求するアクセス許可を、[AppSource](https://appsource.microsoft.com) からメール アドインをインストールする前に表示できます。Exchange 管理センターで、インストールしたアドインに必要なアクセス許可を表示することもできます。

## <a name="restricted-permission"></a>制限付きアクセス許可


  **Restricted** アクセス許可は、最も基本的なアクセス許可レベルです。このアクセス許可を要求するには、マニフェストの [Permissions](/javascript/api/manifest/permissions) 要素内で **Restricted** を指定します。メール アドインがマニフェストで特定のアクセス許可を要求していない場合、Outlook は既定でこのアクセス許可をそのアドインに割り当てます。

### <a name="can-do"></a>できること

- アイテムの件名または本文から [特定のエンティティのみを取得](match-strings-in-an-item-as-well-known-entities.md) (電話番号、アドレス、URL)。

- 閲覧フォームまたは新規作成フォームの現在のアイテムが特定のアイテムの種類であることを要求する [ItemIs アクティブ化ルール](activation-rules.md#itemis-rule)を指定、または、選択したアイテムでサポートされる既知のエンティティ (電話番号、アドレス、URL) の小さなサブセットに一致する [ItemHasKnownEntity ルール](match-strings-in-an-item-as-well-known-entities.md)を指定。

- ユーザーまたはアイテムに関する特定の情報に関連 **しない** プロパティとメソッドへのアクセス (これを実行するメンバーのリストは、次のセクションを参照)。

### <a name="cant-do"></a>できないこと

- 連絡先、電子メール アドレス、会議の提案、またはタスク提案エンティティで [ItemHasKnownEntity](/javascript/api/manifest/rule#itemhasknownentity-rule) ルールを使用します。

- [ItemHasAttachment](/javascript/api/manifest/rule#itemhasattachment-rule) ルールまたは [ItemHasRegularExpressionMatch](/javascript/api/manifest/rule#itemhasregularexpressionmatch-rule) ルールを使用。

- ユーザーまたはアイテムの情報に関連する次のリストに示すメンバーへのアクセス。このリストのメンバーにアクセスしようとすると、**null** が返され、Outlook がメール アドインにアクセス許可の引き上げを要求していることを伝えるエラー メッセージが表示されます。

  - [item.addFileAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.addItemAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.attachments](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.bcc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.body](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.cc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.from](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.getRegExMatches](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.getRegExMatchesByName](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.optionalAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.organizer](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.removeAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.requiredAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.sender](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.to](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [mailbox.getCallbackTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
  - [mailbox.getUserIdentityTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
  - [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
  - [mailbox.userProfile](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#properties)
  - [Body](/javascript/api/outlook/office.body) およびその子メンバーすべて
  - [Location](/javascript/api/outlook/office.location) およびその子メンバーすべて
  - [Recipients](/javascript/api/outlook/office.recipients) およびその子メンバーすべて
  - [Subject](/javascript/api/outlook/office.subject) およびその子メンバーすべて
  - [Time](/javascript/api/outlook/office.time) およびその子メンバーすべて

## <a name="readitem-permission"></a>ReadItem アクセス許可

**ReadItem** アクセス許可は、アクセス許可モデルの中でその次に位置するアクセス許可レベルです。このアクセス許可を要求するには、マニフェストの **Permissions** 要素内で **ReadItem** を指定します。

### <a name="can-do"></a>できること

- 閲覧フォームまたは [新規作成フォーム](item-data.md)の現在のアイテムの [すべてのプロパティの読み取り](get-and-set-item-data-in-a-compose-form.md)。たとえば、閲覧フォームの [item.to](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) および新規作成フォームの [item.to.getAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-getasync-member(1))。

- Exchange Web Services (EWS) または [Outlook REST API](use-rest-api.md) で[アイテムの添付ファイルを取得する](get-attachments-of-an-outlook-item.md)か、アイテム全体を取得するためのコールバック トークンを取得。

- そのアイテムのアドインが設定する[カスタム プロパティの書き込み](/javascript/api/outlook/office.customproperties)。

- アイテムの件名または本文から、サブセットだけでなく、[存在する既知のエンティティをすべて取得する](match-strings-in-an-item-as-well-known-entities.md)。

- [ItemHasKnownEntity](activation-rules.md#itemhasknownentity-rule) ルールの [既知のエンティティ](/javascript/api/manifest/rule#itemhasknownentity-rule)、または [ItemHasRegularExpressionMatch](activation-rules.md#itemhasregularexpressionmatch-rule) ルールの [正規表現](/javascript/api/manifest/rule#itemhasregularexpressionmatch-rule)をすべて使用します。 次の例は、スキーマ v1.1 に従っています。 選択したメッセージの件名または本文に既知のエンティティが 1 つ以上見つかった場合にアドインをアクティブ化するルールが表示されます。

  ```XML
    <Permissions>ReadItem</Permissions>
        <Rule xsi:type="RuleCollection" Mode="And">
        <Rule xsi:type="ItemIs" FormType = "Read" ItemType="Message" />
        <Rule xsi:type="RuleCollection" Mode="Or">
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="PhoneNumber" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Url" />
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="MeetingSuggestion" />
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="TaskSuggestion" />
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="EmailAddress" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Contact" />
    </Rule>
  ```

### <a name="cant-do"></a>できないこと

- **mailbox.getCallbackTokenAsync** によって提供されるトークンを次の目的に使用すること。
  - Outlook REST API を使用した現在のアイテムの更新または削除、またはユーザーのメールボックスにあるその他アイテムへのアクセス。
  - Outlook REST API を使用した現在の予定表イベント アイテムの取得。

- 次の API を使用します。
  - [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
  - [item.addFileAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.addItemAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.bcc.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))
  - [item.bcc.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))
  - [item.body.prependAsync](/javascript/api/outlook/office.body#outlook-office-body-prependasync-member(1))
  - [item.body.setAsync](/javascript/api/outlook/office.body#outlook-office-body-setasync-member(1))
  - [item.body.setSelectedDataAsync](/javascript/api/outlook/office.body#outlook-office-body-setselecteddataasync-member(1))
  - [item.cc.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))
  - [item.cc.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))
  - [item.end.setAsync](/javascript/api/outlook/office.time#outlook-office-time-setasync-member(1))
  - [item.location.setAsync](/javascript/api/outlook/office.location#outlook-office-location-setasync-member(1))
  - [item.optionalAttendees.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))
  - [item.optionalAttendees.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))
  - [item.removeAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.requiredAttendees.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))
  - [item.requiredAttendees.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))
  - [item.start.setAsync](/javascript/api/outlook/office.time#outlook-office-time-setasync-member(1))
  - [item.subject.setAsync](/javascript/api/outlook/office.subject#outlook-office-subject-setasync-member(1))
  - [item.to.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))
  - [item.to.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))

## <a name="readwriteitem-permission"></a>ReadWriteItem アクセス許可

マニフェストの **Permissions** 要素に **ReadWriteItem** を指定すると、このアクセス許可を要求できます。作成フォームでアクティブになり、書き込みメソッド (**Message.to.addAsync** または **Message.to.setAsync**) を使用するメール アドインは、このレベル以上のアクセス許可を使用する必要があります。

### <a name="can-do"></a>できること

- Outlook で閲覧または新規作成されているアイテムの[すべてのアイテム レベルのプロパティを読み書き](item-data.md)。

- そのアイテムで[添付ファイルを追加または削除](add-and-remove-attachments-to-an-item-in-a-compose-form.md)。

- **Mailbox.makeEWSRequestAsync** を除Officeメール アドインに適用可能な JavaScript API の他のすべてのメンバーを使用します。

### <a name="cant-do"></a>できないこと

- **mailbox.getCallbackTokenAsync** によって提供されるトークンを次の目的に使用すること。
  - Outlook REST API を使用した現在のアイテムの更新または削除、またはユーザーのメールボックスにあるその他アイテムへのアクセス。
  - Outlook REST API を使用した現在の予定表イベント アイテムの取得。

- **Mailbox.makeEWSRequestAsync** の使用。

## <a name="readwritemailbox-permission"></a>ReadWriteMailbox アクセス許可

**ReadWriteMailbox** アクセス許可は、最高のアクセス許可レベルです。このアクセス許可を要求するには、マニフェストの **Permissions** 要素内で **ReadWriteMailbox** を指定します。

**ReadWriteItem** アクセス許可がサポートする内容に加え、**mailbox.getCallbackTokenAsync** が提供するトークンでは、Exchange Web Services (EWS) 操作または Outlook REST API を使用して以下を行うためのアクセス権が提供されます。

- ユーザーのメール ボックスのアイテムのすべてのプロパティの読み取りと書き込み。
- そのメール ボックスのフォルダーまたはアイテムの作成、読み取り、書き込み。
- そのメール ボックスからのアイテムの送信。

**mailbox.makeEWSRequestAsync を使用** すると、次の EWS 操作にアクセスできます。

- [CopyItem](/exchange/client-developer/web-service-reference/copyitem-operation)
- [CreateFolder](/exchange/client-developer/web-service-reference/createfolder-operation)
- [CreateItem](/exchange/client-developer/web-service-reference/createitem-operation)
- [FindConversation](/exchange/client-developer/web-service-reference/findconversation-operation)
- [FindFolder](/exchange/client-developer/web-service-reference/findfolder-operation)
- [FindItem](/exchange/client-developer/web-service-reference/finditem-operation)
- [GetConversationItems](/exchange/client-developer/web-service-reference/getconversationitems-operation)
- [GetFolder](/exchange/client-developer/web-service-reference/getfolder-operation)
- [GetItem](/exchange/client-developer/web-service-reference/getitem-operation)
- [MarkAsJunk](/exchange/client-developer/web-service-reference/markasjunk-operation)
- [MoveItem](/exchange/client-developer/web-service-reference/moveitem-operation)
- [SendItem](/exchange/client-developer/web-service-reference/senditem-operation)
- [UpdateFolder](/exchange/client-developer/web-service-reference/updatefolder-operation)
- [UpdateItem](/exchange/client-developer/web-service-reference/updateitem-operation)

サポートされていない操作を使用すると、エラーが返されます。

## <a name="see-also"></a>関連項目

- [Outlook アドインに関するプライバシー、アクセス許可、セキュリティ](../concepts/privacy-and-security.md)
- [Outlook アイテム内の文字列を既知のエンティティとして照合する](match-strings-in-an-item-as-well-known-entities.md)
