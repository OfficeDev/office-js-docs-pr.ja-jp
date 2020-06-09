---
title: Outlook アドインにピン留め可能な作業ウィンドウを実装する
description: アドイン コマンド用の作業ウィンドウ UX シェイプは、開いたメッセージまたは会議出席依頼の右側に縦方向の作業ウィンドウを開きます。アドインは、このウィンドウを使用することで、より詳細な対話式操作に対応した UI を提供できようになります。
ms.date: 02/28/2020
localization_priority: Normal
ms.openlocfilehash: ea9dc255bfb3b689a05d880007282da011edef3e
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44605319"
---
# <a name="implement-a-pinnable-task-pane-in-outlook"></a>Outlook にピン留め可能な作業ウィンドウを実装する

アドイン コマンド用の[作業ウィンドウ](add-in-commands-for-outlook.md#launching-a-task-pane) UX シェイプは、開いたメッセージまたは会議出席依頼の右側に縦方向の作業ウィンドウを開きます。アドインは、このウィンドウを使用することで、より詳細な対話式操作 (複数フィールドの入力など) に対応した UI を提供できようになります。この作業ウィンドウは、メッセージの一覧を表示しているときに、閲覧ウィンドウに表示できます。これにより、メッセージのすばやい処理が可能になります。

ただし、既定では、ユーザーが新しいメッセージを選択すると、閲覧ウィンドウ内で開いていたメッセージのアドイン作業ウィンドウは自動的に閉じられます。頻繁に使用されるアドインの場合、ユーザーはそのウィンドウを開いたままにして、メッセージごとにアドインを有効化する手間がなくなることを望むでしょう。ピン留め可能な作業ウィンドウでは、これに該当するオプションをユーザーに提供できます。

> [!NOTE]
> Pinnable 作業ウィンドウ機能は[要件セット 1.5](../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)で導入されていますが、現時点では、次のものを使用して Office 365 サブスクライバーのみが利用できます。
> - Outlook 2016 以降 (現在のまたは Office Insider チャネル内のユーザーのためにビルド7668.2000 以降) (段階的提供チャネルのユーザー用に7900以降をビルドする)
> - Outlook 2016 以降 (バージョン16.13.503 以降)
> - モダン Outlook on the web

> [!IMPORTANT]
> 次の場合、ピン留め可能な作業ウィンドウは使用できません。
> - 予定および会議
> - Outlook.com

## <a name="support-task-pane-pinning"></a>作業ウィンドウのピン留めをサポートする

ピン留めのサポートを追加する際の最初の手順は、アドインの[マニフェスト](manifests.md)で実行します。この手順は、作業ウィンドウのボタンについて記述する [SupportsPinning](../reference/manifest/action.md#supportspinning) 要素を `Action` 要素に追加することで実行します。

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

イベント ハンドラは、オブジェクト リテラルの単一パラメーターを受け入れる必要があります。このオブジェクトの `type` プロパティは、`Office.EventType.ItemChanged` に設定されます。イベントが呼び出されたときには、既に、`Office.context.mailbox.item` オブジェクトは現在選択されているアイテムを反映するように更新されています。

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

[Office.context.mailbox.addHandlerAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) メソッドを使用して、`Office.EventType.ItemChanged` イベント用のイベント ハンドラを登録します。これは、作業ウィンドウの `Office.initialize` 関数で実行する必要があります。

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
