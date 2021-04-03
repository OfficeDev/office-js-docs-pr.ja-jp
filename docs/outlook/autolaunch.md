---
title: イベント ベースのライセンス認証用に Outlook アドインを構成する (プレビュー)
description: イベント ベースのライセンス認証用に Outlook アドインを構成する方法について学習します。
ms.topic: article
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: b9a4460b05af14f57eecb1bf4181c706843537b2
ms.sourcegitcommit: 074526a6dca8381dbdabf2705474c5ae6753b829
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/02/2021
ms.locfileid: "51506148"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a><span data-ttu-id="e495e-103">イベント ベースのライセンス認証用に Outlook アドインを構成する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="e495e-103">Configure your Outlook add-in for event-based activation (preview)</span></span>

<span data-ttu-id="e495e-104">イベント ベースのアクティブ化機能がない場合、ユーザーはタスクを完了するためにアドインを明示的に起動する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e495e-104">Without the event-based activation feature, a user has to explicitly launch an add-in to complete their tasks.</span></span> <span data-ttu-id="e495e-105">この機能を使用すると、特定のイベントに基づいて、特にすべてのアイテムに適用される操作に基づいてタスクを実行できます。</span><span class="sxs-lookup"><span data-stu-id="e495e-105">This feature enables your add-in to run tasks based on certain events, particularly for operations that apply to every item.</span></span> <span data-ttu-id="e495e-106">作業ウィンドウや UI レス機能と統合することもできます。</span><span class="sxs-lookup"><span data-stu-id="e495e-106">You can also integrate with the task pane and UI-less functionality.</span></span> <span data-ttu-id="e495e-107">現時点では、次のイベントがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="e495e-107">At present, the following events are supported.</span></span>

- <span data-ttu-id="e495e-108">`OnNewMessageCompose`: 新しいメッセージを作成する場合 (返信、全員に返信、転送を含む)</span><span class="sxs-lookup"><span data-stu-id="e495e-108">`OnNewMessageCompose`: On composing a new message (includes reply, reply all, and forward)</span></span>
- <span data-ttu-id="e495e-109">`OnNewAppointmentOrganizer`: 新しい予定を作成する場合</span><span class="sxs-lookup"><span data-stu-id="e495e-109">`OnNewAppointmentOrganizer`: On creating a new appointment</span></span>

  > [!IMPORTANT]
  > <span data-ttu-id="e495e-110">この機能は **、下** 書きや既存の予定など、アイテムの編集ではアクティブ化されない。</span><span class="sxs-lookup"><span data-stu-id="e495e-110">This feature does **not** activate on editing an item, for example, a draft or an existing appointment.</span></span>

<span data-ttu-id="e495e-111">このチュートリアルの最後には、新しいメッセージが作成されるたびに実行されるアドインがあります。</span><span class="sxs-lookup"><span data-stu-id="e495e-111">By the end of this walkthrough, you'll have an add-in that runs whenever a new message is created.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="e495e-112">この機能は、Outlook on [](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) the web および Microsoft 365 サブスクリプションを使用した Windows でのプレビューでのみサポートされます。</span><span class="sxs-lookup"><span data-stu-id="e495e-112">This feature is only supported for [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web and on Windows with a Microsoft 365 subscription.</span></span> <span data-ttu-id="e495e-113">詳細 [については、この記事の「イベント](#how-to-preview-the-event-based-activation-feature) ベースのライセンス認証機能をプレビューする方法」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e495e-113">See [How to preview the event-based activation feature](#how-to-preview-the-event-based-activation-feature) in this article for more details.</span></span>
>
> <span data-ttu-id="e495e-114">プレビュー機能は予告なしに変更される可能性があるため、実稼働アドインでは使用できません。</span><span class="sxs-lookup"><span data-stu-id="e495e-114">Because preview features are subject to change without notice, they shouldn't be used in production add-ins.</span></span>

## <a name="how-to-preview-the-event-based-activation-feature"></a><span data-ttu-id="e495e-115">イベント ベースのアクティブ化機能をプレビューする方法</span><span class="sxs-lookup"><span data-stu-id="e495e-115">How to preview the event-based activation feature</span></span>

<span data-ttu-id="e495e-116">イベント ベースのアクティブ化機能を試してみてください。</span><span class="sxs-lookup"><span data-stu-id="e495e-116">We invite you to try out the event-based activation feature!</span></span> <span data-ttu-id="e495e-117">GitHub を通じてフィードバックを提供することで、シナリオと改善方法についてお知らせします(このページの最後にある「フィードバック」セクションを参照してください)。</span><span class="sxs-lookup"><span data-stu-id="e495e-117">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="e495e-118">この機能をプレビューするには、次の方法を使用します。</span><span class="sxs-lookup"><span data-stu-id="e495e-118">To preview this feature:</span></span>

- <span data-ttu-id="e495e-119">Outlook on the web の場合:</span><span class="sxs-lookup"><span data-stu-id="e495e-119">For Outlook on the web:</span></span>
  - <span data-ttu-id="e495e-120">[Microsoft 365 テナントで対象となるリリースを構成します](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)。</span><span class="sxs-lookup"><span data-stu-id="e495e-120">[Configure targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span>
  - <span data-ttu-id="e495e-121">CDN の **ベータ** ライブラリを参照する ( https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) .</span><span class="sxs-lookup"><span data-stu-id="e495e-121">Reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="e495e-122">TypeScript [のコンパイルと](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) 定義の種類定義ファイルIntelliSense CDN と [DefinitelyTyped で確認できます](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts)。</span><span class="sxs-lookup"><span data-stu-id="e495e-122">The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span></span> <span data-ttu-id="e495e-123">これらの種類は、 を使用してインストールできます `npm install --save-dev @types/office-js-preview` 。</span><span class="sxs-lookup"><span data-stu-id="e495e-123">You can install these types with `npm install --save-dev @types/office-js-preview`.</span></span>
- <span data-ttu-id="e495e-124">Outlook on Windows の場合: 最小必要なビルドは 16.0.13729.20000 です。</span><span class="sxs-lookup"><span data-stu-id="e495e-124">For Outlook on Windows: The minimum required build is 16.0.13729.20000.</span></span> <span data-ttu-id="e495e-125">ベータ版 [ビルドOfficeアクセスするには、Insider](https://insider.office.com) プログラムOffice参加します。</span><span class="sxs-lookup"><span data-stu-id="e495e-125">Join the [Office Insider program](https://insider.office.com) for access to Office beta builds.</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="e495e-126">環境を設定する</span><span class="sxs-lookup"><span data-stu-id="e495e-126">Set up your environment</span></span>

<span data-ttu-id="e495e-127">Outlook の [クイック スタートを](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) 完了し、新しいアドイン用の Yeoman ジェネレーターを使用してアドイン プロジェクトOffice作成します。</span><span class="sxs-lookup"><span data-stu-id="e495e-127">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="e495e-128">マニフェストを構成する</span><span class="sxs-lookup"><span data-stu-id="e495e-128">Configure the manifest</span></span>

<span data-ttu-id="e495e-129">アドインのイベント ベースのアクティブ化を有効にするには、マニフェストのノードで [Runtimes](../reference/manifest/runtimes.md) 要素と [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) 拡張ポイント `VersionOverridesV1_1` を構成する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e495e-129">To enable event-based activation of your add-in, you must configure the [Runtimes](../reference/manifest/runtimes.md) element and [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) extension point in the `VersionOverridesV1_1` node of the manifest.</span></span> <span data-ttu-id="e495e-130">今のところ、 `DesktopFormFactor` サポートされている唯一のフォーム ファクターです。</span><span class="sxs-lookup"><span data-stu-id="e495e-130">For now, `DesktopFormFactor` is the only supported form factor.</span></span>

1. <span data-ttu-id="e495e-131">コード エディターで、クイック スタート プロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="e495e-131">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="e495e-132">プロジェクトの **manifest.xml** にあるファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="e495e-132">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="e495e-133">ノード全体 (開くタグと閉じるタグを含む) を選択し、次の XML に置き換え `<VersionOverrides>` 、変更を保存します。</span><span class="sxs-lookup"><span data-stu-id="e495e-133">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML, then save your changes.</span></span>

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

<span data-ttu-id="e495e-134">Outlook on Windows では JavaScript ファイルを使用し、Outlook on the web では同じ JavaScript ファイルを参照できる HTML ファイルを使用します。</span><span class="sxs-lookup"><span data-stu-id="e495e-134">Outlook on Windows uses a JavaScript file, while Outlook on the web uses an HTML file that can reference the same JavaScript file.</span></span> <span data-ttu-id="e495e-135">Outlook プラットフォームが最終的に Outlook クライアントに基づいて HTML または JavaScript を使用するかどうかを決定する場合は、マニフェストのノードでこれらの両方のファイルへの参照を指定する `Resources` 必要があります。</span><span class="sxs-lookup"><span data-stu-id="e495e-135">You must provide references to both these files in the `Resources` node of the manifest as the Outlook platform ultimately determines whether to use HTML or JavaScript based on the Outlook client.</span></span> <span data-ttu-id="e495e-136">そのため、イベント処理を構成するには、要素内の HTML の場所を指定し、その子要素で HTML によってインライン化または参照される JavaScript ファイルの場所 `Runtime` `Override` を指定します。</span><span class="sxs-lookup"><span data-stu-id="e495e-136">As such, to configure event handling, provide the location of the HTML in the `Runtime` element, then in its `Override` child element provide the location of the JavaScript file inlined or referenced by the HTML.</span></span>

> [!TIP]
> <span data-ttu-id="e495e-137">Outlook アドインのマニフェストの詳細については、「Outlook アドイン マニフェスト [」を参照してください](manifests.md)。</span><span class="sxs-lookup"><span data-stu-id="e495e-137">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-event-handling"></a><span data-ttu-id="e495e-138">イベント処理の実装</span><span class="sxs-lookup"><span data-stu-id="e495e-138">Implement event handling</span></span>

<span data-ttu-id="e495e-139">選択したイベントの処理を実装する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e495e-139">You have to implement handling for your selected events.</span></span>

<span data-ttu-id="e495e-140">このシナリオでは、新しいアイテムを作成する処理を追加します。</span><span class="sxs-lookup"><span data-stu-id="e495e-140">In this scenario, you'll add handling for composing new items.</span></span>

1. <span data-ttu-id="e495e-141">同じクイック スタート プロジェクトで、コード エディター **で ./src/commands/commands.js** ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="e495e-141">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="e495e-142">関数の `action` 後に、次の JavaScript 関数を挿入します。</span><span class="sxs-lookup"><span data-stu-id="e495e-142">After the `action` function, insert the following JavaScript functions.</span></span>

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

1. <span data-ttu-id="e495e-143">Office アドインの Yeoman ジェネレーターによって生成されたこのプロジェクトで Outlook **on the web** で機能する関数については、ファイルの末尾に次のステートメントを追加します。</span><span class="sxs-lookup"><span data-stu-id="e495e-143">For the functions to work in **Outlook on the web** with this project generated by the Yeoman generator for Office Add-ins, add the following statements at the end of the file.</span></span>

    ```js
    g.onMessageComposeHandler = onMessageComposeHandler;
    g.onAppointmentComposeHandler = onAppointmentComposeHandler;
    ```

1. <span data-ttu-id="e495e-144">Outlook on Windows で機能する関数 **については**、ファイルの末尾に次の JavaScript コードを追加します。</span><span class="sxs-lookup"><span data-stu-id="e495e-144">For the functions to work in **Outlook on Windows**, add the following JavaScript code at the end of the file.</span></span>

    ```js
    if (Office.actions) {
      // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
      Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
      Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
    }
    ```

    <span data-ttu-id="e495e-145">**注**: 確認すると `Office.actions` 、Outlook on the web ではこれらのステートメントが無視されます。</span><span class="sxs-lookup"><span data-stu-id="e495e-145">**Note**: Checking for `Office.actions` ensures that Outlook on the web ignores these statements.</span></span>

1. <span data-ttu-id="e495e-146">変更内容を保存します。</span><span class="sxs-lookup"><span data-stu-id="e495e-146">Save your changes.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="e495e-147">試してみる</span><span class="sxs-lookup"><span data-stu-id="e495e-147">Try it out</span></span>

1. <span data-ttu-id="e495e-148">プロジェクトのルート ディレクトリから次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="e495e-148">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="e495e-149">このコマンドを実行すると、ローカル Web サーバーが (まだ実行されていない場合) 起動し、アドインがサイドロードされます。</span><span class="sxs-lookup"><span data-stu-id="e495e-149">When you run this command, the local web server will start (if it's not already running) and your add-in will be sideloaded.</span></span>

    ```command&nbsp;line
    npm start
    ```

1. <span data-ttu-id="e495e-150">Outlook on the web で新しいメッセージを作成します。</span><span class="sxs-lookup"><span data-stu-id="e495e-150">In Outlook on the web, create a new message.</span></span>

    ![構成時に件名が設定されている Outlook on the web のメッセージ ウィンドウのスクリーンショット](../images/outlook-web-autolaunch-1.png)

1. <span data-ttu-id="e495e-152">Outlook on Windows で、新しいメッセージを作成します。</span><span class="sxs-lookup"><span data-stu-id="e495e-152">In Outlook on Windows, create a new message.</span></span>

    ![構成時に件名が設定されている Outlook on Windows のメッセージ ウィンドウのスクリーンショット](../images/outlook-win-autolaunch.png)

    > [!NOTE]
    > <span data-ttu-id="e495e-154">"localhost からこのアドインを開くことができません" というエラーが表示される場合は、ループバックの除外を有効にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="e495e-154">If you see the error "We can't open this add-in from localhost," you'll need to enable a loopback exemption.</span></span>
    >
    > 1. <span data-ttu-id="e495e-155">Outlook を終了します。</span><span class="sxs-lookup"><span data-stu-id="e495e-155">Close Outlook.</span></span>
    > 2. <span data-ttu-id="e495e-156">タスク マネージャー **を開** き、タスク **msoadfs.exeが** 実行されていないか確認します。</span><span class="sxs-lookup"><span data-stu-id="e495e-156">Open the **Task Manager** and ensure that the **msoadfs.exe** process is not running.</span></span>
    > 3. <span data-ttu-id="e495e-157">次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="e495e-157">Run the following command.</span></span>
    >
    >     ```command&nbsp;line
    >     call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_http___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
    >     ```
    >
    > 4. <span data-ttu-id="e495e-158">Outlook を再起動します。</span><span class="sxs-lookup"><span data-stu-id="e495e-158">Restart Outlook.</span></span>

## <a name="debug"></a><span data-ttu-id="e495e-159">Debug</span><span class="sxs-lookup"><span data-stu-id="e495e-159">Debug</span></span>

<span data-ttu-id="e495e-160">独自の機能を実装する場合は、コードのデバッグが必要な場合があります。</span><span class="sxs-lookup"><span data-stu-id="e495e-160">As you implement your own functionality, you may need to debug your code.</span></span> <span data-ttu-id="e495e-161">イベント ベースのアドインのアクティブ化をデバッグする方法のガイダンスについては、「イベント ベースの Outlook アドインをデバッグする」 [を参照してください](debug-autolaunch.md)。</span><span class="sxs-lookup"><span data-stu-id="e495e-161">For guidance on how to debug event-based add-in activation, see [Debug your event-based Outlook add-in](debug-autolaunch.md).</span></span>

## <a name="event-based-activation-behavior-and-limitations"></a><span data-ttu-id="e495e-162">イベント ベースのアクティブ化の動作と制限</span><span class="sxs-lookup"><span data-stu-id="e495e-162">Event-based activation behavior and limitations</span></span>

<span data-ttu-id="e495e-163">イベントに基づいてアクティブ化するアドインは、実行時間が短く、軽量で、可能な限り非侵襲的である必要があります。</span><span class="sxs-lookup"><span data-stu-id="e495e-163">Add-ins that activate based on events are expected to be short-running, lightweight, and as non-invasive as possible.</span></span> <span data-ttu-id="e495e-164">アドインが起動イベントの処理を完了したと知らされる場合は、アドインでメソッドを呼び出す必要 `event.completed` があります。</span><span class="sxs-lookup"><span data-stu-id="e495e-164">To signal that your add-in has completed processing the launch event, we recommend you have your add-in call the `event.completed` method.</span></span> <span data-ttu-id="e495e-165">この呼び出しが行われた場合、アドインはイベント ベースのアドインを実行できる最大時間である約 300 秒以内にタイム アウトします。ユーザーが作成ウィンドウを閉じると、アドインも終了します。</span><span class="sxs-lookup"><span data-stu-id="e495e-165">If that call is not made, the add-in will time out within approximately 300 seconds, the maximum length of time allowed for running event-based add-ins. The add-in also ends when the user closes the compose window.</span></span>

<span data-ttu-id="e495e-166">ユーザーが同じイベントにサブスクライブしている複数のアドインがある場合、Outlook プラットフォームはアドインを特定の順序で起動します。</span><span class="sxs-lookup"><span data-stu-id="e495e-166">If the user has multiple add-ins that subscribed to the same event, the Outlook platform launches the add-ins in no particular order.</span></span> <span data-ttu-id="e495e-167">現在、アクティブに実行できるイベント ベースのアドインは 5 つのみです。</span><span class="sxs-lookup"><span data-stu-id="e495e-167">Currently, only five event-based add-ins can be actively running.</span></span> <span data-ttu-id="e495e-168">追加のアドインはキューにプッシュされ、以前にアクティブだったアドインが完了または非アクティブ化されると実行されます。</span><span class="sxs-lookup"><span data-stu-id="e495e-168">Any additional add-ins are pushed to a queue then run as previously active add-ins are completed or deactivated.</span></span>

<span data-ttu-id="e495e-169">ユーザーは、アドインの実行を開始した現在のメール アイテムから切り替えまたは移動できます。</span><span class="sxs-lookup"><span data-stu-id="e495e-169">The user can switch or navigate away from the current mail item where the add-in started running.</span></span> <span data-ttu-id="e495e-170">起動されたアドインは、バックグラウンドで操作を終了します。</span><span class="sxs-lookup"><span data-stu-id="e495e-170">The add-in that was launched will finish its operation in the background.</span></span>

<span data-ttu-id="e495e-171">UI Office.js変更する API の一部は、イベント ベースのアドインでは使用できない場合があります。ブロックされている API を次に示します。</span><span class="sxs-lookup"><span data-stu-id="e495e-171">Some Office.js APIs that change or alter the UI are not allowed from event-based add-ins. The following are the blocked APIs:</span></span>

- <span data-ttu-id="e495e-172">[ `Office.context.auth` : ] の下</span><span class="sxs-lookup"><span data-stu-id="e495e-172">Under `Office.context.auth`:</span></span>
  - `getAccessToken`
  - `getAccessTokenAsync`
- <span data-ttu-id="e495e-173">[ `Office.context.mailbox` : ] の下</span><span class="sxs-lookup"><span data-stu-id="e495e-173">Under `Office.context.mailbox`:</span></span>
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- <span data-ttu-id="e495e-174">[ `Office.context.mailbox.item` : ] の下</span><span class="sxs-lookup"><span data-stu-id="e495e-174">Under `Office.context.mailbox.item`:</span></span>
  - `close`
- <span data-ttu-id="e495e-175">[ `Office.context.ui` : ] の下</span><span class="sxs-lookup"><span data-stu-id="e495e-175">Under `Office.context.ui`:</span></span>
  - `displayDialogAsync`
  - `messageParent`

## <a name="see-also"></a><span data-ttu-id="e495e-176">関連項目</span><span class="sxs-lookup"><span data-stu-id="e495e-176">See also</span></span>

<span data-ttu-id="e495e-177">[Outlook アドイン マニフェスト](manifests.md) 
[イベント ベースのアドインをデバッグする方法](debug-autolaunch.md)</span><span class="sxs-lookup"><span data-stu-id="e495e-177">[Outlook add-in manifests](manifests.md)
[How to debug event-based add-ins](debug-autolaunch.md)</span></span>
