---
title: イベント ベースOutlook用にアドインを構成する (プレビュー)
description: イベント ベースのアクティブ化Outlookアドインを構成する方法について学習します。
ms.topic: article
ms.date: 04/29/2021
localization_priority: Normal
ms.openlocfilehash: 45f9ff16b3aca0a1fb8f3a8ee3d9ffa8e0f33ea2
ms.sourcegitcommit: 6057afc1776e1667b231d2e9809d261d372151f6
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/30/2021
ms.locfileid: "52100300"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a><span data-ttu-id="a056b-103">イベント ベースOutlook用にアドインを構成する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="a056b-103">Configure your Outlook add-in for event-based activation (preview)</span></span>

<span data-ttu-id="a056b-104">イベント ベースのアクティブ化機能がない場合、ユーザーはタスクを完了するためにアドインを明示的に起動する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a056b-104">Without the event-based activation feature, a user has to explicitly launch an add-in to complete their tasks.</span></span> <span data-ttu-id="a056b-105">この機能を使用すると、特定のイベントに基づいて、特にすべてのアイテムに適用される操作に基づいてタスクを実行できます。</span><span class="sxs-lookup"><span data-stu-id="a056b-105">This feature enables your add-in to run tasks based on certain events, particularly for operations that apply to every item.</span></span> <span data-ttu-id="a056b-106">作業ウィンドウや UI レス機能と統合することもできます。</span><span class="sxs-lookup"><span data-stu-id="a056b-106">You can also integrate with the task pane and UI-less functionality.</span></span> <span data-ttu-id="a056b-107">現時点では、次のイベントがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="a056b-107">At present, the following events are supported.</span></span>

|<span data-ttu-id="a056b-108">イベント</span><span class="sxs-lookup"><span data-stu-id="a056b-108">Event</span></span>|<span data-ttu-id="a056b-109">説明</span><span class="sxs-lookup"><span data-stu-id="a056b-109">Description</span></span>|
|---|---|
|`OnNewMessageCompose`|<span data-ttu-id="a056b-110">新しいメッセージを作成する場合 (返信、すべて返信、転送を含む) が、下書きなど編集時には作成されません。</span><span class="sxs-lookup"><span data-stu-id="a056b-110">On composing a new message (includes reply, reply all, and forward) but not on editing, for example, a draft.</span></span>|
|`OnNewAppointmentOrganizer`|<span data-ttu-id="a056b-111">既存の予定の編集ではなく、新しい予定を作成する場合。</span><span class="sxs-lookup"><span data-stu-id="a056b-111">On creating a new appointment but not on editing an existing one.</span></span>|
|`OnMessageAttachmentsChanged`|<span data-ttu-id="a056b-112">メッセージの作成中に添付ファイルを追加または削除する場合。</span><span class="sxs-lookup"><span data-stu-id="a056b-112">On adding or removing attachments while composing a message.</span></span>|
|`OnAppointmentAttachmentsChanged`|<span data-ttu-id="a056b-113">予定の作成中に添付ファイルを追加または削除する場合。</span><span class="sxs-lookup"><span data-stu-id="a056b-113">On adding or removing attachments while composing an appointment.</span></span>|
|`OnMessageRecipientsChanged`|<span data-ttu-id="a056b-114">メッセージの作成中に受信者を追加または削除する場合。</span><span class="sxs-lookup"><span data-stu-id="a056b-114">On adding or removing recipients while composing a message.</span></span>|
|`OnAppointmentAttendeesChanged`|<span data-ttu-id="a056b-115">予定の作成中に出席者を追加または削除する場合。</span><span class="sxs-lookup"><span data-stu-id="a056b-115">On adding or removing attendees while composing an appointment.</span></span>|
|`OnAppointmentTimeChanged`|<span data-ttu-id="a056b-116">予定の作成中に日付/時刻を変更する場合。</span><span class="sxs-lookup"><span data-stu-id="a056b-116">On changing date/time while composing an appointment.</span></span>|
|`OnAppointmentRecurrenceChanged`|<span data-ttu-id="a056b-117">予定の作成中に定期的な詳細を追加、変更、または削除する場合。</span><span class="sxs-lookup"><span data-stu-id="a056b-117">On adding, changing, or removing the recurrence details while composing an appointment.</span></span> <span data-ttu-id="a056b-118">日付/時刻が変更された場合、 `OnAppointmentTimeChanged` イベントも発生します。</span><span class="sxs-lookup"><span data-stu-id="a056b-118">If the date/time is changed, the `OnAppointmentTimeChanged` event will also be fired.</span></span>|
|`OnInfoBarDismissClicked`|<span data-ttu-id="a056b-119">メッセージまたは予定アイテムの作成中に通知を却下する場合。</span><span class="sxs-lookup"><span data-stu-id="a056b-119">On dismissing a notification while composing a message or appointment item.</span></span> <span data-ttu-id="a056b-120">通知を追加したアドインだけが通知されます。</span><span class="sxs-lookup"><span data-stu-id="a056b-120">Only the add-in that added the notification will be notified.</span></span>|

<span data-ttu-id="a056b-121">このチュートリアルの最後には、新しいアイテムが作成され、件名が設定されるたびに実行されるアドインがあります。</span><span class="sxs-lookup"><span data-stu-id="a056b-121">By the end of this walkthrough, you'll have an add-in that runs whenever a new item is created and sets the subject.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="a056b-122">この機能は、Web 上[](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)および Outlookサブスクリプションでのプレビュー WindowsでのみMicrosoft 365されます。</span><span class="sxs-lookup"><span data-stu-id="a056b-122">This feature is only supported for [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web and on Windows with a Microsoft 365 subscription.</span></span> <span data-ttu-id="a056b-123">詳細 [については、この記事の「イベント](#how-to-preview-the-event-based-activation-feature) ベースのライセンス認証機能をプレビューする方法」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a056b-123">See [How to preview the event-based activation feature](#how-to-preview-the-event-based-activation-feature) in this article for more details.</span></span>
>
> <span data-ttu-id="a056b-124">プレビュー機能は予告なしに変更される可能性があるため、実稼働アドインでは使用できません。</span><span class="sxs-lookup"><span data-stu-id="a056b-124">Because preview features are subject to change without notice, they shouldn't be used in production add-ins.</span></span>

## <a name="how-to-preview-the-event-based-activation-feature"></a><span data-ttu-id="a056b-125">イベント ベースのアクティブ化機能をプレビューする方法</span><span class="sxs-lookup"><span data-stu-id="a056b-125">How to preview the event-based activation feature</span></span>

<span data-ttu-id="a056b-126">イベント ベースのアクティブ化機能を試してみてください。</span><span class="sxs-lookup"><span data-stu-id="a056b-126">We invite you to try out the event-based activation feature!</span></span> <span data-ttu-id="a056b-127">このページの最後にある「フィードバック」セクションをGitHubフィードバックを提供することで、お客様のシナリオと改善方法をお知らせします。</span><span class="sxs-lookup"><span data-stu-id="a056b-127">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="a056b-128">この機能をプレビューするには、次の方法を使用します。</span><span class="sxs-lookup"><span data-stu-id="a056b-128">To preview this feature:</span></span>

- <span data-ttu-id="a056b-129">Web Outlookの詳細については、次の情報を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a056b-129">For Outlook on the web:</span></span>
  - <span data-ttu-id="a056b-130">[ターゲット リリースをテナントにMicrosoft 365します](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)。</span><span class="sxs-lookup"><span data-stu-id="a056b-130">[Configure targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span>
  - <span data-ttu-id="a056b-131">()**の** ベータ ライブラリを参照 https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) CDN。</span><span class="sxs-lookup"><span data-stu-id="a056b-131">Reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="a056b-132">TypeScript[のコンパイルと](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)定義の種類定義ファイルは、IntelliSenseと[DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts)にあるCDNです。</span><span class="sxs-lookup"><span data-stu-id="a056b-132">The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span></span> <span data-ttu-id="a056b-133">これらの種類は、 を使用してインストールできます `npm install --save-dev @types/office-js-preview` 。</span><span class="sxs-lookup"><span data-stu-id="a056b-133">You can install these types with `npm install --save-dev @types/office-js-preview`.</span></span>
- <span data-ttu-id="a056b-134">[OutlookのWindows: 必要な最小ビルドは 16.0.13729.20000 です。</span><span class="sxs-lookup"><span data-stu-id="a056b-134">For Outlook on Windows: The minimum required build is 16.0.13729.20000.</span></span> <span data-ttu-id="a056b-135">ベータ版[ビルドOfficeアクセスするには、Insider](https://insider.office.com)プログラムOffice参加します。</span><span class="sxs-lookup"><span data-stu-id="a056b-135">Join the [Office Insider program](https://insider.office.com) for access to Office beta builds.</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="a056b-136">環境を設定する</span><span class="sxs-lookup"><span data-stu-id="a056b-136">Set up your environment</span></span>

<span data-ttu-id="a056b-137">クイック スタート[Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator)完了し、Yeoman ジェネレーターを使用してアドイン プロジェクトを作成し、Office作成します。</span><span class="sxs-lookup"><span data-stu-id="a056b-137">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="a056b-138">マニフェストを構成する</span><span class="sxs-lookup"><span data-stu-id="a056b-138">Configure the manifest</span></span>

<span data-ttu-id="a056b-139">アドインのイベント ベースのアクティブ化を有効にするには、マニフェストのノードで [Runtimes](../reference/manifest/runtimes.md) 要素と [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) 拡張ポイント `VersionOverridesV1_1` を構成する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a056b-139">To enable event-based activation of your add-in, you must configure the [Runtimes](../reference/manifest/runtimes.md) element and [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) extension point in the `VersionOverridesV1_1` node of the manifest.</span></span> <span data-ttu-id="a056b-140">今のところ、 `DesktopFormFactor` サポートされている唯一のフォーム ファクターです。</span><span class="sxs-lookup"><span data-stu-id="a056b-140">For now, `DesktopFormFactor` is the only supported form factor.</span></span>

1. <span data-ttu-id="a056b-141">コード エディターで、クイック スタート プロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="a056b-141">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="a056b-142">プロジェクトの **manifest.xml** にあるファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="a056b-142">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="a056b-143">ノード全体 (開くタグと閉じるタグを含む) を選択し、次の XML に置き換え `<VersionOverrides>` 、変更を保存します。</span><span class="sxs-lookup"><span data-stu-id="a056b-143">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML, then save your changes.</span></span>

```XML
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    <Requirements>
      <bt:Sets DefaultMinVersion="1.3">
        <bt:Set Name="Mailbox" />
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- Event-based activation happens in a lightweight runtime.-->
        <Runtimes>
          <!-- HTML file including reference to or inline JavaScript event handlers.
               This is used by Outlook on the web. -->
          <Runtime resid="WebViewRuntime.Url">
            <!-- JavaScript file containing event handlers. This is used by Outlook Desktop. -->
            <Override type="javascript" resid="JSRuntime.Url"/>
          </Runtime>
        </Runtimes>
        <DesktopFormFactor>
          <FunctionFile resid="Commands.Url" />
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="msgReadGroup">
                <Label resid="GroupLabel" />
                <Control xsi:type="Button" id="msgReadOpenPaneButton">
                  <Label resid="TaskpaneButton.Label" />
                  <Supertip>
                    <Title resid="TaskpaneButton.Label" />
                    <Description resid="TaskpaneButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16" />
                    <bt:Image size="32" resid="Icon.32x32" />
                    <bt:Image size="80" resid="Icon.80x80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <SourceLocation resid="Taskpane.Url" />
                  </Action>
                </Control>
                <Control xsi:type="Button" id="ActionButton">
                  <Label resid="ActionButton.Label"/>
                  <Supertip>
                    <Title resid="ActionButton.Label"/>
                    <Description resid="ActionButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>action</FunctionName>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>

          <!-- Can configure other command surface extension points for add-in command support. -->

          <!-- Enable launching the add-in on the included events. -->
          <ExtensionPoint xsi:type="LaunchEvent">
            <LaunchEvents>
              <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
              <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onAppointmentComposeHandler"/>
            </LaunchEvents>
            <!-- Identifies the runtime to be used (also referenced by the Runtime element). -->
            <SourceLocation resid="WebViewRuntime.Url"/>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="https://localhost:3000/assets/icon-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="https://localhost:3000/assets/icon-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="https://localhost:3000/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Commands.Url" DefaultValue="https://localhost:3000/commands.html" />
        <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/taskpane.html" />
        <bt:Url id="WebViewRuntime.Url" DefaultValue="https://localhost:3000/commands.html" />
        <!-- Entry needed for Outlook Desktop. -->
        <bt:Url id="JSRuntime.Url" DefaultValue="https://localhost:3000/src/commands/commands.js" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GroupLabel" DefaultValue="Contoso Add-in"/>
        <bt:String id="TaskpaneButton.Label" DefaultValue="Show Taskpane"/>
        <bt:String id="ActionButton.Label" DefaultValue="Perform an action"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Opens a pane displaying all available properties."/>
        <bt:String id="ActionButton.Tooltip" DefaultValue="Perform an action when clicked."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</VersionOverrides>
```

<span data-ttu-id="a056b-144">OutlookはWindows JavaScript ファイルを使用しますが、web 上Outlookは同じ JavaScript ファイルを参照できる HTML ファイルを使用します。</span><span class="sxs-lookup"><span data-stu-id="a056b-144">Outlook on Windows uses a JavaScript file, while Outlook on the web uses an HTML file that can reference the same JavaScript file.</span></span> <span data-ttu-id="a056b-145">Outlook プラットフォームは最終的に、Outlook クライアントに基づいて HTML または JavaScript を使用するかどうかを決定するために、マニフェストのノードでこれらの両方のファイル `Resources` への参照を提供する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a056b-145">You must provide references to both these files in the `Resources` node of the manifest as the Outlook platform ultimately determines whether to use HTML or JavaScript based on the Outlook client.</span></span> <span data-ttu-id="a056b-146">そのため、イベント処理を構成するには、要素内の HTML の場所を指定し、その子要素で HTML によってインライン化または参照される JavaScript ファイルの場所 `Runtime` `Override` を指定します。</span><span class="sxs-lookup"><span data-stu-id="a056b-146">As such, to configure event handling, provide the location of the HTML in the `Runtime` element, then in its `Override` child element provide the location of the JavaScript file inlined or referenced by the HTML.</span></span>

> [!TIP]
> <span data-ttu-id="a056b-147">アドインのマニフェストのOutlook詳細については、「Outlook[マニフェスト」を参照してください](manifests.md)。</span><span class="sxs-lookup"><span data-stu-id="a056b-147">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-event-handling"></a><span data-ttu-id="a056b-148">イベント処理の実装</span><span class="sxs-lookup"><span data-stu-id="a056b-148">Implement event handling</span></span>

<span data-ttu-id="a056b-149">選択したイベントの処理を実装する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a056b-149">You have to implement handling for your selected events.</span></span>

<span data-ttu-id="a056b-150">このシナリオでは、新しいアイテムを作成する処理を追加します。</span><span class="sxs-lookup"><span data-stu-id="a056b-150">In this scenario, you'll add handling for composing new items.</span></span>

1. <span data-ttu-id="a056b-151">同じクイック スタート プロジェクトで、コード エディター **で ./src/commands/commands.js** ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="a056b-151">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="a056b-152">関数の `action` 後に、次の JavaScript 関数を挿入します。</span><span class="sxs-lookup"><span data-stu-id="a056b-152">After the `action` function, insert the following JavaScript functions.</span></span>

    ```js
    function onMessageComposeHandler(event) {
      setSubject(event);
    }
    function onAppointmentComposeHandler(event) {
      setSubject(event);
    }
    function setSubject(event) {
      Office.context.mailbox.item.subject.setAsync(
        "Set by an event-based add-in!",
        {
          "asyncContext" : event
        },
        function (asyncResult) {
          // Handle success or error.
          if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
            console.error("Failed to set subject: " + JSON.stringify(asyncResult.error));
          }
    
          // Call event.completed() after all work is done.
          asyncResult.asyncContext.completed();
        });
    }
    ```

1. <span data-ttu-id="a056b-153">Office アドインの **Yeoman** ジェネレーターによって生成されたこのプロジェクトを使用して、web 上の Outlook で機能する関数については、ファイルの末尾に次のステートメントを追加します。</span><span class="sxs-lookup"><span data-stu-id="a056b-153">For the functions to work in **Outlook on the web** with this project generated by the Yeoman generator for Office Add-ins, add the following statements at the end of the file.</span></span>

    ```js
    g.onMessageComposeHandler = onMessageComposeHandler;
    g.onAppointmentComposeHandler = onAppointmentComposeHandler;
    ```

1. <span data-ttu-id="a056b-154">関数がファイル上でOutlook **動作** Windows、ファイルの末尾に次の JavaScript コードを追加します。</span><span class="sxs-lookup"><span data-stu-id="a056b-154">For the functions to work in **Outlook on Windows**, add the following JavaScript code at the end of the file.</span></span>

    ```js
    if (Office.actions) {
      // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
      Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
      Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
    }
    ```

    <span data-ttu-id="a056b-155">**注**: チェックを `Office.actions` 実行すると、web 上Outlookこれらのステートメントが無視されます。</span><span class="sxs-lookup"><span data-stu-id="a056b-155">**Note**: Checking for `Office.actions` ensures that Outlook on the web ignores these statements.</span></span>

1. <span data-ttu-id="a056b-156">変更内容を保存します。</span><span class="sxs-lookup"><span data-stu-id="a056b-156">Save your changes.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="a056b-157">試してみる</span><span class="sxs-lookup"><span data-stu-id="a056b-157">Try it out</span></span>

1. <span data-ttu-id="a056b-158">プロジェクトのルート ディレクトリから次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="a056b-158">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="a056b-159">このコマンドを実行すると、ローカル Web サーバーが (まだ実行されていない場合) 起動し、アドインがサイドロードされます。</span><span class="sxs-lookup"><span data-stu-id="a056b-159">When you run this command, the local web server will start (if it's not already running) and your add-in will be sideloaded.</span></span>

    ```command&nbsp;line
    npm start
    ```

1. <span data-ttu-id="a056b-160">Outlook on the web で新しいメッセージを作成します。</span><span class="sxs-lookup"><span data-stu-id="a056b-160">In Outlook on the web, create a new message.</span></span>

    ![作成時に件名が設定Outlook Web 上のメッセージ ウィンドウのスクリーンショット](../images/outlook-web-autolaunch-1.png)

1. <span data-ttu-id="a056b-162">[Outlook] でWindows新しいメッセージを作成します。</span><span class="sxs-lookup"><span data-stu-id="a056b-162">In Outlook on Windows, create a new message.</span></span>

    ![作成時に件名が設定されているOutlookのWindowsウィンドウのスクリーンショット](../images/outlook-win-autolaunch.png)

    > [!NOTE]
    > <span data-ttu-id="a056b-164">"localhost からこのアドインを開くことができません" というエラーが表示される場合は、ループバックの除外を有効にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="a056b-164">If you see the error "We can't open this add-in from localhost," you'll need to enable a loopback exemption.</span></span>
    >
    > 1. <span data-ttu-id="a056b-165">Outlook を終了します。</span><span class="sxs-lookup"><span data-stu-id="a056b-165">Close Outlook.</span></span>
    > 2. <span data-ttu-id="a056b-166">タスク マネージャー **を開** き、タスク **msoadfs.exeが** 実行されていないか確認します。</span><span class="sxs-lookup"><span data-stu-id="a056b-166">Open the **Task Manager** and ensure that the **msoadfs.exe** process is not running.</span></span>
    > 3. <span data-ttu-id="a056b-167">次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="a056b-167">Run the following command.</span></span>
    >
    >     ```command&nbsp;line
    >     call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_http___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
    >     ```
    >
    > 4. <span data-ttu-id="a056b-168">Outlook を再起動します。</span><span class="sxs-lookup"><span data-stu-id="a056b-168">Restart Outlook.</span></span>

## <a name="debug"></a><span data-ttu-id="a056b-169">Debug</span><span class="sxs-lookup"><span data-stu-id="a056b-169">Debug</span></span>

<span data-ttu-id="a056b-170">独自の機能を実装する場合は、コードのデバッグが必要な場合があります。</span><span class="sxs-lookup"><span data-stu-id="a056b-170">As you implement your own functionality, you may need to debug your code.</span></span> <span data-ttu-id="a056b-171">イベント ベースのアドインのアクティブ化をデバッグする方法のガイダンスについては、「Debug [your event-based Outlook アドイン」を参照してください](debug-autolaunch.md)。</span><span class="sxs-lookup"><span data-stu-id="a056b-171">For guidance on how to debug event-based add-in activation, see [Debug your event-based Outlook add-in](debug-autolaunch.md).</span></span>

## <a name="event-based-activation-behavior-and-limitations"></a><span data-ttu-id="a056b-172">イベント ベースのアクティブ化の動作と制限</span><span class="sxs-lookup"><span data-stu-id="a056b-172">Event-based activation behavior and limitations</span></span>

<span data-ttu-id="a056b-173">イベントに基づいてアクティブ化するアドインは、実行時間が短く、軽量で、可能な限り非侵襲的である必要があります。</span><span class="sxs-lookup"><span data-stu-id="a056b-173">Add-ins that activate based on events are expected to be short-running, lightweight, and as non-invasive as possible.</span></span> <span data-ttu-id="a056b-174">アドインが起動イベントの処理を完了したと知らされる場合は、アドインでメソッドを呼び出す必要 `event.completed` があります。</span><span class="sxs-lookup"><span data-stu-id="a056b-174">To signal that your add-in has completed processing the launch event, we recommend you have your add-in call the `event.completed` method.</span></span> <span data-ttu-id="a056b-175">この呼び出しが行われた場合、アドインはイベント ベースのアドインを実行できる最大時間である約 300 秒以内にタイム アウトします。ユーザーが作成ウィンドウを閉じると、アドインも終了します。</span><span class="sxs-lookup"><span data-stu-id="a056b-175">If that call is not made, the add-in will time out within approximately 300 seconds, the maximum length of time allowed for running event-based add-ins. The add-in also ends when the user closes the compose window.</span></span>

<span data-ttu-id="a056b-176">ユーザーが同じイベントにサブスクライブしている複数のアドインがある場合、Outlook プラットフォームは特定の順序でアドインを起動します。</span><span class="sxs-lookup"><span data-stu-id="a056b-176">If the user has multiple add-ins that subscribed to the same event, the Outlook platform launches the add-ins in no particular order.</span></span> <span data-ttu-id="a056b-177">現在、アクティブに実行できるイベント ベースのアドインは 5 つのみです。</span><span class="sxs-lookup"><span data-stu-id="a056b-177">Currently, only five event-based add-ins can be actively running.</span></span> <span data-ttu-id="a056b-178">追加のアドインはキューにプッシュされ、以前にアクティブだったアドインが完了または非アクティブ化されると実行されます。</span><span class="sxs-lookup"><span data-stu-id="a056b-178">Any additional add-ins are pushed to a queue then run as previously active add-ins are completed or deactivated.</span></span>

<span data-ttu-id="a056b-179">ユーザーは、アドインの実行を開始した現在のメール アイテムから切り替えまたは移動できます。</span><span class="sxs-lookup"><span data-stu-id="a056b-179">The user can switch or navigate away from the current mail item where the add-in started running.</span></span> <span data-ttu-id="a056b-180">起動されたアドインは、バックグラウンドで操作を終了します。</span><span class="sxs-lookup"><span data-stu-id="a056b-180">The add-in that was launched will finish its operation in the background.</span></span>

<span data-ttu-id="a056b-181">UI Office.js変更する API の一部は、イベント ベースのアドインでは使用できない場合があります。ブロックされている API を次に示します。</span><span class="sxs-lookup"><span data-stu-id="a056b-181">Some Office.js APIs that change or alter the UI are not allowed from event-based add-ins. The following are the blocked APIs:</span></span>

- <span data-ttu-id="a056b-182">[ `Office.context.auth` : ] の下</span><span class="sxs-lookup"><span data-stu-id="a056b-182">Under `Office.context.auth`:</span></span>
  - `getAccessToken`
  - `getAccessTokenAsync`
- <span data-ttu-id="a056b-183">[ `Office.context.mailbox` : ] の下</span><span class="sxs-lookup"><span data-stu-id="a056b-183">Under `Office.context.mailbox`:</span></span>
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- <span data-ttu-id="a056b-184">[ `Office.context.mailbox.item` : ] の下</span><span class="sxs-lookup"><span data-stu-id="a056b-184">Under `Office.context.mailbox.item`:</span></span>
  - `close`
- <span data-ttu-id="a056b-185">[ `Office.context.ui` : ] の下</span><span class="sxs-lookup"><span data-stu-id="a056b-185">Under `Office.context.ui`:</span></span>
  - `displayDialogAsync`
  - `messageParent`

## <a name="see-also"></a><span data-ttu-id="a056b-186">関連項目</span><span class="sxs-lookup"><span data-stu-id="a056b-186">See also</span></span>

<span data-ttu-id="a056b-187">[Outlook アドイン マニフェスト](manifests.md) 
[イベント ベースのアドインをデバッグする方法](debug-autolaunch.md)</span><span class="sxs-lookup"><span data-stu-id="a056b-187">[Outlook add-in manifests](manifests.md)
[How to debug event-based add-ins](debug-autolaunch.md)</span></span>
