---
title: Outlook アドインのアクセス許可を理解する
description: Outlook アドインでは、必要なアクセス許可のレベルをマニフェストで指定します。使用可能なレベルは Restricted、ReadItem、ReadWriteItem、ReadWriteMailbox です。
ms.date: 02/19/2020
localization_priority: Normal
ms.openlocfilehash: 7b0b481edc77170bb395d86f77688bc976f8e6e4
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292518"
---
# <a name="understanding-outlook-add-in-permissions"></a>Outlook アドインのアクセス許可を理解する

Outlook アドインでは、必要なアクセス許可のレベルをマニフェストで指定します。使用可能なレベルは **Restricted**、**ReadItem**、**ReadWriteItem**、**ReadWriteMailbox** です。これらのレベルのアクセス許可は累積されます。**Restricted** は最低レベルであり、それぞれの上位レベルには、下位レベルのアクセス許可がすべて含まれます。**ReadWriteMailbox** にはサポートされるアクセス許可がすべて含まれます。

メール アドインが要求するアクセス許可を、[AppSource](https://appsource.microsoft.com) からメール アドインをインストールする前に表示できます。Exchange 管理センターで、インストールしたアドインに必要なアクセス許可を表示することもできます。

## <a name="restricted-permission"></a>制限付きアクセス許可


  **Restricted** アクセス許可は、最も基本的なアクセス許可レベルです。このアクセス許可を要求するには、マニフェストの [Permissions](../reference/manifest/permissions.md) 要素内で **Restricted** を指定します。メール アドインがマニフェストで特定のアクセス許可を要求していない場合、Outlook は既定でこのアクセス許可をそのアドインに割り当てます。

### <a name="can-do"></a>できること

- アイテムの件名または本文から [特定のエンティティのみを取得](match-strings-in-an-item-as-well-known-entities.md) (電話番号、アドレス、URL)。

- 閲覧フォームまたは新規作成フォームの現在のアイテムが特定のアイテムの種類であることを要求する [ItemIs アクティブ化ルール](activation-rules.md#itemis-rule)を指定、または、選択したアイテムでサポートされる既知のエンティティ (電話番号、アドレス、URL) の小さなサブセットに一致する [ItemHasKnownEntity ルール](match-strings-in-an-item-as-well-known-entities.md)を指定。

- ユーザーまたはアイテムに関する特定の情報に関連**しない**プロパティとメソッドへのアクセス (これを実行するメンバーのリストは、次のセクションを参照)。

### <a name="cant-do"></a>できないこと

- 連絡先、電子メールアドレス、会議提案、またはタスク提案エンティティで [Itemhasknownentity](../reference/manifest/rule.md#itemhasknownentity-rule) ルールを使用します。

- [ItemHasAttachment](../reference/manifest/rule.md#itemhasattachment-rule) ルールまたは [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) ルールを使用。

- ユーザーまたはアイテムの情報に関連する次のリストに示すメンバーへのアクセス。このリストのメンバーにアクセスしようとすると、**null** が返され、Outlook がメール アドインにアクセス許可の引き上げを要求していることを伝えるエラー メッセージが表示されます。

    - [item.addFileAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [item.addItemAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [item.attachments](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [item.bcc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [item.body](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [item.cc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [item.from](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [item.getRegExMatches](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [item.getRegExMatchesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [item.optionalAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [item.organizer](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [item.removeAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [item.requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [item.sender](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [item.to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)
    - [mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
    - [mailbox.getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
    - [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
    - [mailbox.userProfile](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties)
    - [Body](/javascript/api/outlook/office.body) およびその子メンバーすべて
    - [Location](/javascript/api/outlook/office.location) およびその子メンバーすべて
    - [Recipients](/javascript/api/outlook/office.recipients) およびその子メンバーすべて
    - [Subject](/javascript/api/outlook/office.subject) およびその子メンバーすべて
    - [Time](/javascript/api/outlook/office.time) およびその子メンバーすべて

## <a name="readitem-permission"></a>ReadItem アクセス許可

**ReadItem** アクセス許可は、アクセス許可モデルの中でその次に位置するアクセス許可レベルです。このアクセス許可を要求するには、マニフェストの **Permissions** 要素内で **ReadItem** を指定します。

### <a name="can-do"></a>できること

- 閲覧フォームまたは [新規作成フォーム](item-data.md)の現在のアイテムの [すべてのプロパティの読み取り](get-and-set-item-data-in-a-compose-form.md)。たとえば、閲覧フォームの [item.to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) および新規作成フォームの [item.to.getAsync](/javascript/api/outlook/office.Recipients#getasync-options--callback-)。

- Exchange Web Services (EWS) または [Outlook REST API](use-rest-api.md) で[アイテムの添付ファイルを取得する](get-attachments-of-an-outlook-item.md)か、アイテム全体を取得するためのコールバック トークンを取得。

- そのアイテムのアドインが設定する[カスタム プロパティの書き込み](/javascript/api/outlook/office.CustomProperties)。

- アイテムの件名または本文から、サブセットだけでなく、[存在する既知のエンティティをすべて取得する](match-strings-in-an-item-as-well-known-entities.md)。

- 
  [ItemHasKnownEntity](activation-rules.md#itemhasknownentity-rule) ルールの [既知のエンティティ](../reference/manifest/rule.md#itemhasknownentity-rule)、または [ItemHasRegularExpressionMatch](activation-rules.md#itemhasregularexpressionmatch-rule) ルールの [正規表現](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule)をすべて使用します。次の例は、スキーマ v1.1 に従っています。選択されたメッセージの件名または本文に既知のエンティティが 1 つ以上ある場合にアクティブ化されるルールを示しています。

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

- 次のいずれかの API を使用すること。
    - [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)
    - [item.addFileAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [item.addItemAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [item.bcc.addAsync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [item.bcc.setAsync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)
    - [item.body.prependAsync](/javascript/api/outlook/office.Body#prependasync-data--options--callback-)
    - [item.body.setAsync](/javascript/api/outlook/office.Body#setasync-data--options--callback-)
    - [item.body.setSelectedDataAsync](/javascript/api/outlook/office.Body#setselecteddataasync-data--options--callback-)
    - [item.cc.addAsync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [item.cc.setAsync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)
    - [item.end.setAsync](/javascript/api/outlook/office.Time#setasync-datetime--options--callback-)
    - [item.location.setAsync](/javascript/api/outlook/office.Location#setasync-location--options--callback-)
    - [item.optionalAttendees.addAsync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [item.optionalAttendees.setAsync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)
    - [item.removeAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
    - [item.requiredAttendees.addAsync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [item.requiredAttendees.setAsync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)
    - [item.start.setAsync](/javascript/api/outlook/office.Time#setasync-datetime--options--callback-)
    - [item.subject.setAsync](/javascript/api/outlook/office.Subject#setasync-subject--options--callback-)
    - [item.to.addAsync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)
    - [item.to.setAsync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)

## <a name="readwriteitem-permission"></a>ReadWriteItem アクセス許可

マニフェストの **Permissions** 要素に **ReadWriteItem** を指定すると、このアクセス許可を要求できます。作成フォームでアクティブになり、書き込みメソッド (**Message.to.addAsync** または **Message.to.setAsync**) を使用するメール アドインは、このレベル以上のアクセス許可を使用する必要があります。

### <a name="can-do"></a>できること

- Outlook で閲覧または新規作成されているアイテムの[すべてのアイテム レベルのプロパティを読み書き](item-data.md)。

- そのアイテムで[添付ファイルを追加または削除](add-and-remove-attachments-to-an-item-in-a-compose-form.md)。

- メールアドインに適用される Office JavaScript API の他のすべてのメンバー ( **makeEWSRequestAsync**を除く) を使用します。

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

**mailbox.makeEWSRequestAsync** を使用して、次の EWS の操作にアクセスできます。

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
