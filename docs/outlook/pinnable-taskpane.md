---
title: Outlook アドインにピン留め可能な作業ウィンドウを実装する
description: アドイン コマンド用の作業ウィンドウ UX シェイプは、開いたメッセージまたは会議出席依頼の右側に縦方向の作業ウィンドウを開きます。アドインは、このウィンドウを使用することで、より詳細な対話式操作に対応した UI を提供できようになります。
ms.date: 11/18/2019
localization_priority: Normal
ms.openlocfilehash: 94c136a74dfddac1af663aea06c3c6ca27f22dcd
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166462"
---
# <a name="implement-a-pinnable-task-pane-in-outlook"></a><span data-ttu-id="60261-103">Outlook にピン留め可能な作業ウィンドウを実装する</span><span class="sxs-lookup"><span data-stu-id="60261-103">Implement a pinnable task pane in Outlook</span></span>

<span data-ttu-id="60261-p101">アドイン コマンド用の[作業ウィンドウ](add-in-commands-for-outlook.md#launching-a-task-pane) UX シェイプは、開いたメッセージまたは会議出席依頼の右側に縦方向の作業ウィンドウを開きます。アドインは、このウィンドウを使用することで、より詳細な対話式操作 (複数フィールドの入力など) に対応した UI を提供できようになります。この作業ウィンドウは、メッセージの一覧を表示しているときに、閲覧ウィンドウに表示できます。これにより、メッセージのすばやい処理が可能になります。</span><span class="sxs-lookup"><span data-stu-id="60261-p101">The [task pane](add-in-commands-for-outlook.md#launching-a-task-pane) UX shape for add-in commands opens a vertical task pane to the right of an open message or meeting request, allowing the add-in to provide UI for more detailed interactions (filling in multiple fields, etc.). This task pane can be shown in the Reading Pane when viewing a list of messages, allowing for quick processing of a message.</span></span>

<span data-ttu-id="60261-p102">ただし、既定では、ユーザーが新しいメッセージを選択すると、閲覧ウィンドウ内で開いていたメッセージのアドイン作業ウィンドウは自動的に閉じられます。頻繁に使用されるアドインの場合、ユーザーはそのウィンドウを開いたままにして、メッセージごとにアドインを有効化する手間がなくなることを望むでしょう。ピン留め可能な作業ウィンドウでは、これに該当するオプションをユーザーに提供できます。</span><span class="sxs-lookup"><span data-stu-id="60261-p102">However, by default, if a user has an add-in task pane open for a message in the Reading Pane, and then selects a new message, the task pane is automatically closed. For a heavily-used add-in, the user may prefer to keep that pane open, eliminating the need to reactivate the add-in on each message. With pinnable task panes, your add-in can give the user that option.</span></span>

> [!NOTE]
> <span data-ttu-id="60261-109">現在、ピン留め可能な作業ウィンドウは、Windows 版 Outlook 2016 以降 (Current Channel または Office Insider Channel のユーザーはビルド 7668.2000 以降、Deferred Channel のユーザーはビルド 7900.xxxx 以降)、Mac 用 Outlook 2016 以降 (バージョン 16.13.503 以降)、および Outlook on the web を使用している Office 365 サブスクライバーがご利用になれます。</span><span class="sxs-lookup"><span data-stu-id="60261-109">Pinnable task panes are currently available to Office 365 subscribers using Outlook 2016 or later on Windows (build 7668.2000 or later for users in the Current or Office Insider Channels, build 7900.xxxx or later for users in Deferred channels), Outlook 2016 or later on Mac (version 16.13.503 or later), and Outlook on the web.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="60261-110">次の場合、ピン留め可能な作業ウィンドウは使用できません。</span><span class="sxs-lookup"><span data-stu-id="60261-110">Pinnable task panes are not available for the following.</span></span>
> - <span data-ttu-id="60261-111">予定および会議</span><span class="sxs-lookup"><span data-stu-id="60261-111">Appointments/Meetings</span></span>
> - <span data-ttu-id="60261-112">Outlook.com</span><span class="sxs-lookup"><span data-stu-id="60261-112">Outlook.com</span></span>

## <a name="support-task-pane-pinning"></a><span data-ttu-id="60261-113">作業ウィンドウのピン留めをサポートする</span><span class="sxs-lookup"><span data-stu-id="60261-113">Support task pane pinning</span></span>

<span data-ttu-id="60261-p103">ピン留めのサポートを追加する際の最初の手順は、アドインの[マニフェスト](manifests.md)で実行します。この手順は、作業ウィンドウのボタンについて記述する [SupportsPinning](../reference/manifest/action.md#supportspinning) 要素を `Action` 要素に追加することで実行します。</span><span class="sxs-lookup"><span data-stu-id="60261-p103">The first step is to add pinning support, which is done in the add-in [manifest](manifests.md). This is done by adding the [SupportsPinning](../reference/manifest/action.md#supportspinning) element to the `Action` element that describes the task pane button.</span></span>

<span data-ttu-id="60261-116">`SupportsPinning` 要素は、VersionOverrides v1.1 スキーマで定義されているため、v1.0 と v1.1 のどちらの場合も [VersionOverrides](../reference/manifest/versionoverrides.md) 要素を含める必要があります。</span><span class="sxs-lookup"><span data-stu-id="60261-116">The `SupportsPinning` element is defined in the VersionOverrides v1.1 schema, so you will need to include a [VersionOverrides](../reference/manifest/versionoverrides.md) element both for v1.0 and v1.1.</span></span>

> [!NOTE]
> <span data-ttu-id="60261-117">Outlook アドインを [AppSource](https://appsource.microsoft.com) に[発行](../publish/publish.md)する予定であれば、**SupportsPinning** 要素を使う場合、[AppSource 検証](/office/dev/store/validation-policies)に合格するためには、アドインのコンテンツを静的にすることはできません。また、メールボックスで開かれているか選択されているメッセージに関連するデータを、そのコンテンツで明確に表示する必要があります。</span><span class="sxs-lookup"><span data-stu-id="60261-117">If you plan to [publish](../publish/publish.md) your Outlook add-in to [AppSource](https://appsource.microsoft.com), when you use the **SupportsPinning** element, in order to pass [AppSource validation](/office/dev/store/validation-policies), your add-in content must not be static and it must clearly display data related to the message that is open or selected in the mailbox.</span></span>

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

<span data-ttu-id="60261-118">完全な例については、[command-demo のサンプル マニフェスト](https://github.com/OfficeDev/outlook-add-in-command-demo/blob/master/command-demo-manifest.xml)の `msgReadOpenPaneButton` コントロールをご覧ください。</span><span class="sxs-lookup"><span data-stu-id="60261-118">For a full example, see the `msgReadOpenPaneButton` control in the [command-demo sample manifest](https://github.com/OfficeDev/outlook-add-in-command-demo/blob/master/command-demo-manifest.xml).</span></span>

## <a name="handling-ui-updates-based-on-currently-selected-message"></a><span data-ttu-id="60261-119">現在選択されているメッセージに基づいた UI の更新を処理する</span><span class="sxs-lookup"><span data-stu-id="60261-119">Handling UI updates based on currently selected message</span></span>

<span data-ttu-id="60261-120">現在のアイテムに基づいて作業ウィンドウの UI または内部変数を更新するには、変更の通知を受け取るイベント ハンドラの登録が必要になります。</span><span class="sxs-lookup"><span data-stu-id="60261-120">To update your task pane's UI or internal variables based on the current item, you'll need to register an event handler to get notified of the change.</span></span>

### <a name="implement-the-event-handler"></a><span data-ttu-id="60261-121">イベント ハンドラを実装する</span><span class="sxs-lookup"><span data-stu-id="60261-121">Implement the event handler</span></span>

<span data-ttu-id="60261-p104">イベント ハンドラは、オブジェクト リテラルの単一パラメーターを受け入れる必要があります。このオブジェクトの `type` プロパティは、`Office.EventType.ItemChanged` に設定されます。イベントが呼び出されたときには、既に、`Office.context.mailbox.item` オブジェクトは現在選択されているアイテムを反映するように更新されています。</span><span class="sxs-lookup"><span data-stu-id="60261-p104">The event handler should accept a single parameter, which is an object literal. The `type` property of this object will be set to `Office.EventType.ItemChanged`. When the event is called, the `Office.context.mailbox.item` object is already updated to reflect the currently selected item.</span></span>

```js
function itemChanged(eventArgs) {
  // Update UI based on the new current item
  UpdateTaskPaneUI(Office.context.mailbox.item);
}
```

> [!IMPORTANT]
> <span data-ttu-id="60261-125">ItemChanged イベントのイベント ハンドラーの実装では、Office.content.mailbox.item が null かどうかを確認する必要があります。</span><span class="sxs-lookup"><span data-stu-id="60261-125">The implementation of event handlers for an ItemChanged event should check whether or not the Office.content.mailbox.item is null.</span></span>
>
> ```js
> // Example implementation
> function UpdateTaskPaneUI(item)
> {
>   // Assuming that item is always a read item (instead of a compose item).
>   if (item != null) console.log(item.subject);
> }
> ```

### <a name="register-the-event-handler"></a><span data-ttu-id="60261-126">イベント ハンドラーを登録する</span><span class="sxs-lookup"><span data-stu-id="60261-126">Register the event handler</span></span>

<span data-ttu-id="60261-p105">[Office.context.mailbox.addHandlerAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) メソッドを使用して、`Office.EventType.ItemChanged` イベント用のイベント ハンドラを登録します。これは、作業ウィンドウの `Office.initialize` 関数で実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="60261-p105">Use the [Office.context.mailbox.addHandlerAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method to register your event handler for the `Office.EventType.ItemChanged` event. This should be done in the `Office.initialize` function for your task pane.</span></span>

```js
Office.initialize = function (reason) {
  $(document).ready(function () {

    // Set up ItemChanged event
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);

    UpdateTaskPaneUI(Office.context.mailbox.item);
  });
};
```

## <a name="see-also"></a><span data-ttu-id="60261-129">関連項目</span><span class="sxs-lookup"><span data-stu-id="60261-129">See also</span></span>

<span data-ttu-id="60261-130">ピン留め可能な作業ウィンドウを実装するサンプル アドインについては、GitHub の [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo) をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="60261-130">For an example add-in that implements a pinnable task pane, see [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo) on GitHub.</span></span>
