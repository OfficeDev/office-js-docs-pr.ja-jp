---
title: Outlook アドインにピン留め可能な作業ウィンドウを実装する
description: アドイン コマンド用の作業ウィンドウ UX シェイプは、開いたメッセージまたは会議出席依頼の右側に縦方向の作業ウィンドウを開きます。アドインは、このウィンドウを使用することで、より詳細な対話式操作に対応した UI を提供できようになります。
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: 39af3a532d553835b02709301c998a78dc9958bb
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093869"
---
# <a name="implement-a-pinnable-task-pane-in-outlook"></a>Outlook にピン留め可能な作業ウィンドウを実装する

The [task pane](add-in-commands-for-outlook.md#launching-a-task-pane) UX shape for add-in commands opens a vertical task pane to the right of an open message or meeting request, allowing the add-in to provide UI for more detailed interactions (filling in multiple fields, etc.). This task pane can be shown in the Reading Pane when viewing a list of messages, allowing for quick processing of a message.

However, by default, if a user has an add-in task pane open for a message in the Reading Pane, and then selects a new message, the task pane is automatically closed. For a heavily-used add-in, the user may prefer to keep that pane open, eliminating the need to reactivate the add-in on each message. With pinnable task panes, your add-in can give the user that option.

> [!NOTE]
> Pinnable 作業ウィンドウ機能は[要件セット 1.5](../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)で導入されていますが、現時点では、次のものを使用して Microsoft 365 サブスクライバーのみが利用できます。
> - Outlook 2016 以降 (現在のまたは Office Insider チャネル内のユーザーのためにビルド7668.2000 以降) (段階的提供チャネルのユーザー用に7900以降をビルドする)
> - Outlook 2016 以降 (バージョン16.13.503 以降)
> - モダン Outlook on the web

> [!IMPORTANT]
> 次の場合、ピン留め可能な作業ウィンドウは使用できません。
> - 予定および会議
> - Outlook.com

## <a name="support-task-pane-pinning"></a>作業ウィンドウのピン留めをサポートする

The first step is to add pinning support, which is done in the add-in [manifest](manifests.md). This is done by adding the [SupportsPinning](../reference/manifest/action.md#supportspinning) element to the `Action` element that describes the task pane button.

`SupportsPinning` 要素は、VersionOverrides v1.1 スキーマで定義されているため、v1.0 と v1.1 のどちらの場合も [VersionOverrides](../reference/manifest/versionoverrides.md) 要素を含める必要があります。

> [!NOTE]
> Outlook アドインを [AppSource](https://appsource.microsoft.com) に[発行](../publish/publish.md)する予定であれば、**SupportsPinning** 要素を使う場合、[AppSource 検証](/legal/marketplace/certification-policies)に合格するためには、アドインのコンテンツを静的にすることはできません。また、メールボックスで開かれているか選択されているメッセージに関連するデータを、そのコンテンツで明確に表示する必要があります。

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

Use the [Office.context.mailbox.addHandlerAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method to register your event handler for the `Office.EventType.ItemChanged` event. This should be done in the `Office.initialize` function for your task pane.

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
