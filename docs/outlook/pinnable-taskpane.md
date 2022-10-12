---
title: Outlook アドインにピン留め可能な作業ウィンドウを実装する
description: アドイン コマンド用の作業ウィンドウ UX シェイプは、開いたメッセージまたは会議出席依頼の右側に縦方向の作業ウィンドウを開きます。アドインは、このウィンドウを使用することで、より詳細な対話式操作に対応した UI を提供できようになります。
ms.date: 10/13/2022
ms.localizationpriority: medium
ms.openlocfilehash: 834d43a6046ddaa63a7c8899cfd5b07d0ea80ef6
ms.sourcegitcommit: a2df9538b3deb32ae3060ecb09da15f5a3d6cb8d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/12/2022
ms.locfileid: "68541123"
---
# <a name="implement-a-pinnable-task-pane-in-outlook"></a>Outlook にピン留め可能な作業ウィンドウを実装する

The [task pane](add-in-commands-for-outlook.md#launch-a-task-pane) UX shape for add-in commands opens a vertical task pane to the right of an open message or meeting request, allowing the add-in to provide UI for more detailed interactions (filling in multiple fields, etc.). This task pane can be shown in the Reading Pane when viewing a list of messages, allowing for quick processing of a message.

However, by default, if a user has an add-in task pane open for a message in the Reading Pane, and then selects a new message, the task pane is automatically closed. For a heavily-used add-in, the user may prefer to keep that pane open, eliminating the need to reactivate the add-in on each message. With pinnable task panes, your add-in can give the user that option.

> [!NOTE]
> ピン留め可能な作業ウィンドウ機能は [要件セット 1.5](/javascript/api/requirement-sets/outlook/requirement-set-1.5/outlook-requirement-set-1.5) で導入されましたが、現時点では、次を使用して Microsoft 365 サブスクライバーのみが使用できます。
>
> - Windows 版 Outlook 2016 以降 (Current Channel または Office Insider Channel のユーザーはビルド 7668.2000 以降、Deferred Channel のユーザーはビルド 7900.xxxx 以降) および Outlook Online
> - Mac Outlook 2016 以降 (バージョン 16.13.503 以降)
> - モダン Outlook on the web

> [!IMPORTANT]
> ピン留め可能な作業ウィンドウは、次の作業ウィンドウでは使用できません。
>
> - 予定および会議
> - Outlook.com

> [!TIP]
> Outlook アドインを [AppSource](https://appsource.microsoft.com) に[発行](../publish/publish.md)する予定で、ピン留め可能な作業ウィンドウ用に構成されている場合、[AppSource の検証](/legal/marketplace/certification-policies)に合格するためにアドイン コンテンツは静的ではなく、メールボックスで開いているか選択されているメッセージに関連するデータを明確に表示する必要があります。

## <a name="support-task-pane-pinning"></a>作業ウィンドウのピン留めをサポートする

ピン留めのサポートを追加する際の最初の手順は、アドインのマニフェストで実行します。 マークアップは、マニフェストの種類によって異なります。

# <a name="xml-manifest"></a>[XML マニフェスト](#tab/xmlmanifest)

作業ウィンドウ ボタンを記述する要素に **\<Action\>** [SupportsPinning](/javascript/api/manifest/action#supportspinning) 要素を追加します。 次に例を示します。

```xml
<!-- Task pane button -->
<Control xsi:type="Button" id="msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Supertip>
    <Title resid="paneReadSuperTipTitle" />
    <Description resid="paneReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="green-icon-16" />
    <bt:Image size="32" resid="green-icon-32" />
    <bt:Image size="80" resid="green-icon-80" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
    <SupportsPinning>true</SupportsPinning>
  </Action>
</Control>
```

この要素は **\<SupportsPinning\>** VersionOverrides v1.1 スキーマで定義されているため、v1.0 と v1.1 の両方に [VersionOverrides](/javascript/api/manifest/versionoverrides) 要素を含める必要があります。

# <a name="teams-manifest-developer-preview"></a>[Teams マニフェスト (開発者プレビュー)](#tab/jsonmanifest)

作業ウィンドウを開くボタンまたはメニュー項目を定義する `true`"actions" 配列内のオブジェクトに 、"ピン留め可能" プロパティを追加します。 次に例を示します。

```json
"actions": [
    {
        "id": "OpenTaskPane",
        "type": "openPage",
        "view": "TaskPaneView",
        "displayName": "OpenTaskPane",
        "pinnable": true
    }
]
```

---

完全な例については、[command-demo のサンプル マニフェスト](https://github.com/OfficeDev/outlook-add-in-command-demo/blob/master/command-demo-manifest.xml)の `msgReadOpenPaneButton` コントロールをご覧ください。

## <a name="handling-ui-updates-based-on-currently-selected-message"></a>現在選択されているメッセージに基づいた UI の更新を処理する

現在のアイテムに基づいて作業ウィンドウの UI または内部変数を更新するには、変更の通知を受け取るイベント ハンドラの登録が必要になります。

### <a name="implement-the-event-handler"></a>イベント ハンドラを実装する

The event handler should accept a single parameter, which is an object literal. The `type` property of this object will be set to `Office.EventType.ItemChanged`. When the event is called, the `Office.context.mailbox.item` object is already updated to reflect the currently selected item.

```js
function itemChanged(eventArgs) {
  // Update UI based on the new current item
  UpdateTaskPaneUI(Office.context.mailbox.item);
}
```

> [!IMPORTANT]
> ItemChanged イベントのイベント ハンドラーの実装では、Office.content.mailbox.item が null かどうかを確認する必要があります。
>
> ```js
> // Example implementation
> function UpdateTaskPaneUI(item)
> {
>   // Assuming that item is always a read item (instead of a compose item).
>   if (item != null) console.log(item.subject);
> }
> ```

### <a name="register-the-event-handler"></a>イベント ハンドラーを登録する

Use the [Office.context.mailbox.addHandlerAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) method to register your event handler for the `Office.EventType.ItemChanged` event. This should be done in the `Office.initialize` function for your task pane.

```js
Office.initialize = function (reason) {
  $(document).ready(function () {

    // Set up ItemChanged event
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);

    UpdateTaskPaneUI(Office.context.mailbox.item);
  });
};
```

## <a name="see-also"></a>関連項目

ピン留め可能な作業ウィンドウを実装するサンプル アドインについては、GitHub の [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo) をご覧ください。
