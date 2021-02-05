---
title: イベント ベースのアクティブ化用に Outlook アドインを構成する (プレビュー)
description: イベント ベースのアクティブ化用に Outlook アドインを構成する方法について学習します。
ms.topic: article
ms.date: 02/03/2021
localization_priority: Normal
ms.openlocfilehash: a4fce335738d1bcff2be43e4e609998be89fca20
ms.sourcegitcommit: 8546889a759590c3798ce56e311d9e46f0171413
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/04/2021
ms.locfileid: "50104855"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a><span data-ttu-id="c05ea-103">イベント ベースのアクティブ化用に Outlook アドインを構成する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="c05ea-103">Configure your Outlook add-in for event-based activation (preview)</span></span>

<span data-ttu-id="c05ea-104">イベント ベースのアクティブ化機能がない場合、ユーザーは自分のタスクを完了するためにアドインを明示的に起動する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c05ea-104">Without the event-based activation feature, a user has to explicitly launch an add-in to complete their tasks.</span></span> <span data-ttu-id="c05ea-105">この機能を使用すると、特にすべてのアイテムに適用される操作に関して、特定のイベントに基づいてタスクを実行できます。</span><span class="sxs-lookup"><span data-stu-id="c05ea-105">This feature enables your add-in to run tasks based on certain events, particularly for operations that apply to every item.</span></span> <span data-ttu-id="c05ea-106">作業ウィンドウと UI を使用する機能を統合することもできます。</span><span class="sxs-lookup"><span data-stu-id="c05ea-106">You can also integrate with the task pane and UI-less functionality.</span></span> <span data-ttu-id="c05ea-107">現時点では、次のイベントがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="c05ea-107">At present, the following events are supported.</span></span>

- <span data-ttu-id="c05ea-108">`OnNewMessageCompose`: 新しいメッセージの作成時 (返信、全員に返信、転送を含む)</span><span class="sxs-lookup"><span data-stu-id="c05ea-108">`OnNewMessageCompose`: On composing a new message (includes reply, reply all, and forward)</span></span>
- <span data-ttu-id="c05ea-109">`OnNewAppointmentOrganizer`: 新しい予定の作成時</span><span class="sxs-lookup"><span data-stu-id="c05ea-109">`OnNewAppointmentOrganizer`: On creating a new appointment</span></span>

  > [!IMPORTANT]
  > <span data-ttu-id="c05ea-110">この機能は **、下** 書きや既存の予定など、アイテムの編集時にはアクティブ化されない。</span><span class="sxs-lookup"><span data-stu-id="c05ea-110">This feature does **not** activate on editing an item, for example, a draft or an existing appointment.</span></span>

<span data-ttu-id="c05ea-111">このチュートリアルの終わりまでに、新しいメッセージが作成されるたびに実行されるアドインが作成されます。</span><span class="sxs-lookup"><span data-stu-id="c05ea-111">By the end of this walkthrough, you'll have an add-in that runs whenever a new message is created.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c05ea-112">この機能は、Microsoft 365 サブスクリプションを使用する Outlook on the web および Windows でのプレビューでのみサポートされます。 [](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)</span><span class="sxs-lookup"><span data-stu-id="c05ea-112">This feature is only supported for [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web and Windows with a Microsoft 365 subscription.</span></span> <span data-ttu-id="c05ea-113">詳細 [については、この記事のイベント ベース](#how-to-preview-the-event-based-activation-feature) のアクティブ化機能をプレビューする方法を参照してください。</span><span class="sxs-lookup"><span data-stu-id="c05ea-113">See [How to preview the event-based activation feature](#how-to-preview-the-event-based-activation-feature) in this article for more details.</span></span>
>
> <span data-ttu-id="c05ea-114">プレビュー機能は予告なしに変更されることがありますので、実稼働アドインでは使用できません。</span><span class="sxs-lookup"><span data-stu-id="c05ea-114">Because preview features are subject to change without notice, they shouldn't be used in production add-ins.</span></span>

## <a name="how-to-preview-the-event-based-activation-feature"></a><span data-ttu-id="c05ea-115">イベント ベースのアクティブ化機能をプレビューする方法</span><span class="sxs-lookup"><span data-stu-id="c05ea-115">How to preview the event-based activation feature</span></span>

<span data-ttu-id="c05ea-116">イベント ベースのアクティブ化機能をお試しください。</span><span class="sxs-lookup"><span data-stu-id="c05ea-116">We invite you to try out the event-based activation feature!</span></span> <span data-ttu-id="c05ea-117">GitHub を通じてフィードバックを提供することで、お客様のシナリオと改善方法をお知らせします (このページの最後にあるフィードバックセクションをご覧ください)。</span><span class="sxs-lookup"><span data-stu-id="c05ea-117">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="c05ea-118">この機能をプレビューするには:</span><span class="sxs-lookup"><span data-stu-id="c05ea-118">To preview this feature:</span></span>

- <span data-ttu-id="c05ea-119">Outlook on the web の場合:</span><span class="sxs-lookup"><span data-stu-id="c05ea-119">For Outlook on the web:</span></span>
  - <span data-ttu-id="c05ea-120">[Microsoft 365 テナントで対象指定リリースを構成します](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)。</span><span class="sxs-lookup"><span data-stu-id="c05ea-120">[Configure targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span>
  - <span data-ttu-id="c05ea-121">CDN で **ベータ** ライブラリを参照します ( https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) .</span><span class="sxs-lookup"><span data-stu-id="c05ea-121">Reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="c05ea-122">TypeScript [のコンパイルと](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) 読み取りIntelliSense定義ファイルは CDN と [DefinitelyTyped で見つかりました](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts)。</span><span class="sxs-lookup"><span data-stu-id="c05ea-122">The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span></span> <span data-ttu-id="c05ea-123">次の種類を使用してインストールできます `npm install --save-dev @types/office-js-preview` 。</span><span class="sxs-lookup"><span data-stu-id="c05ea-123">You can install these types with `npm install --save-dev @types/office-js-preview`.</span></span>
- <span data-ttu-id="c05ea-124">Windows 上の Outlook の場合: 最低限必要なビルドは 16.0.13729.20000 です。</span><span class="sxs-lookup"><span data-stu-id="c05ea-124">For Outlook on Windows: The minimum required build is 16.0.13729.20000.</span></span> <span data-ttu-id="c05ea-125">ベータ ビルド [Officeアクセス](https://insider.office.com) するには、Insider Program にOffice参加します。</span><span class="sxs-lookup"><span data-stu-id="c05ea-125">Join the [Office Insider program](https://insider.office.com) for access to Office beta builds.</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="c05ea-126">環境を設定する</span><span class="sxs-lookup"><span data-stu-id="c05ea-126">Set up your environment</span></span>

<span data-ttu-id="c05ea-127">Outlook クイック [スタートを完了](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) します。このクイック スタートでは、アドイン用の Yeoman ジェネレーターを使用してアドイン Office作成します。</span><span class="sxs-lookup"><span data-stu-id="c05ea-127">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="c05ea-128">マニフェストを構成する</span><span class="sxs-lookup"><span data-stu-id="c05ea-128">Configure the manifest</span></span>

<span data-ttu-id="c05ea-129">アドインのイベント ベースのアクティブ化を有効にするには、マニフェストのノードで [Runtimes](../reference/manifest/runtimes.md) 要素と [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) 拡張ポイント `VersionOverridesV1_1` を構成する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c05ea-129">To enable event-based activation of your add-in, you must configure the [Runtimes](../reference/manifest/runtimes.md) element and [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) extension point in the `VersionOverridesV1_1` node of the manifest.</span></span> <span data-ttu-id="c05ea-130">今のところ、 `DesktopFormFactor` サポートされているフォーム ファクターは 1 つのみです。</span><span class="sxs-lookup"><span data-stu-id="c05ea-130">For now, `DesktopFormFactor` is the only supported form factor.</span></span>

1. <span data-ttu-id="c05ea-131">コード エディターで、クイック スタート プロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="c05ea-131">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="c05ea-132">プロジェクトの **manifest.xml** にある新しいファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="c05ea-132">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="c05ea-133">ノード全体 (開いているタグと閉じるタグを含む) を選択し `<VersionOverrides>` 、次の XML に置き換えてください。</span><span class="sxs-lookup"><span data-stu-id="c05ea-133">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML.</span></span>

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

<span data-ttu-id="c05ea-134">Outlook on Windows では JavaScript ファイルを使用し、Outlook on the web は同じ JavaScript ファイルを参照できる HTML ファイルを使用します。</span><span class="sxs-lookup"><span data-stu-id="c05ea-134">Outlook on Windows uses a JavaScript file, while Outlook on the web uses an HTML file that can reference the same JavaScript file.</span></span> <span data-ttu-id="c05ea-135">Outlook プラットフォームは最終的に Outlook クライアントに基づいて HTML または JavaScript を使用するかどうかを決定しますので、マニフェストのノードでこれらの両方のファイルへの参照を提供する `Resources` 必要があります。</span><span class="sxs-lookup"><span data-stu-id="c05ea-135">You must provide references to both these files in the `Resources` node of the manifest as the Outlook platform ultimately determines whether to use HTML or JavaScript based on the Outlook client.</span></span> <span data-ttu-id="c05ea-136">そのため、イベント処理を構成するには、要素内の HTML の場所を指定し、その子要素で、HTML によってインライン化または参照される JavaScript ファイルの場所を提供します `Runtime` `Override` 。</span><span class="sxs-lookup"><span data-stu-id="c05ea-136">As such, to configure event handling, provide the location of the HTML in the `Runtime` element, then in its `Override` child element provide the location of the JavaScript file inlined or referenced by the HTML.</span></span>

> [!TIP]
> <span data-ttu-id="c05ea-137">Outlook アドインのマニフェストの詳細については、Outlook アドインの [マニフェストを参照してください](manifests.md)。</span><span class="sxs-lookup"><span data-stu-id="c05ea-137">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-event-handling"></a><span data-ttu-id="c05ea-138">イベント処理を実装する</span><span class="sxs-lookup"><span data-stu-id="c05ea-138">Implement event handling</span></span>

<span data-ttu-id="c05ea-139">選択したイベントの処理を実装する必要があります。</span><span class="sxs-lookup"><span data-stu-id="c05ea-139">You have to implement handling for your selected events.</span></span>

<span data-ttu-id="c05ea-140">このシナリオでは、新しいアイテムを作成する処理を追加します。</span><span class="sxs-lookup"><span data-stu-id="c05ea-140">In this scenario, you'll add handling for composing new items.</span></span>

1. <span data-ttu-id="c05ea-141">同じクイック スタート プロジェクトから、コード エディターで **ファイル ./src/commands/commands.js** を開きます。</span><span class="sxs-lookup"><span data-stu-id="c05ea-141">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="c05ea-142">関数の `action` 後に、次の JavaScript 関数を挿入します。</span><span class="sxs-lookup"><span data-stu-id="c05ea-142">After the `action` function, insert the following JavaScript functions.</span></span>

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

1. <span data-ttu-id="c05ea-143">Office アドイン用の Yeoman ジェネレーターによって生成されたこのプロジェクトを使用して Outlook **on the web** で関数を動作するには、ファイルの末尾に次のステートメントを追加します。</span><span class="sxs-lookup"><span data-stu-id="c05ea-143">For the functions to work in **Outlook on the web** with this project generated by the Yeoman generator for Office Add-ins, add the following statements at the end of the file.</span></span>

    ```js
    g.onMessageComposeHandler = onMessageComposeHandler;
    g.onAppointmentComposeHandler = onAppointmentComposeHandler;
    ```

1. <span data-ttu-id="c05ea-144">Windows 上の **Outlook** で関数を動作するには、ファイルの末尾に次の JavaScript コードを追加します。</span><span class="sxs-lookup"><span data-stu-id="c05ea-144">For the functions to work in **Outlook on Windows**, add the following JavaScript code at the end of the file.</span></span>

    ```js
    if (Office.actions) {
      // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
      Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
      Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
    }
    ```

    <span data-ttu-id="c05ea-145">**注**: Outlook on the `Office.actions` web がこれらのステートメントを無視することを確認します。</span><span class="sxs-lookup"><span data-stu-id="c05ea-145">**Note**: Checking for `Office.actions` ensures that Outlook on the web ignores these statements.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="c05ea-146">試してみる</span><span class="sxs-lookup"><span data-stu-id="c05ea-146">Try it out</span></span>

1. <span data-ttu-id="c05ea-147">プロジェクトのルート ディレクトリから次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="c05ea-147">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="c05ea-148">このコマンドを実行すると、ローカル Web サーバーが起動します (まだ実行されていない場合)。</span><span class="sxs-lookup"><span data-stu-id="c05ea-148">When you run this command, the local web server will start (if it's not already running).</span></span>

    ```command&nbsp;line
    npm run dev-server
    ```

1. <span data-ttu-id="c05ea-149">「[テスト用に Outlook アドインをサイドロードする](sideload-outlook-add-ins-for-testing.md)」の手順に従って Outlook アドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="c05ea-149">Follow the instructions in [Sideload Outlook add-ins for testing](sideload-outlook-add-ins-for-testing.md) to sideload the add-in in Outlook.</span></span>

1. <span data-ttu-id="c05ea-150">Outlook on the web で新しいメッセージを作成します。</span><span class="sxs-lookup"><span data-stu-id="c05ea-150">In Outlook on the web, create a new message.</span></span>

    ![Outlook on the web のメッセージ ウィンドウのスクリーンショット(件名が新規作成時に設定されている場合)](../images/outlook-web-autolaunch-1.png)

1. <span data-ttu-id="c05ea-152">Windows 上の Outlook で、新しいメッセージを作成します。</span><span class="sxs-lookup"><span data-stu-id="c05ea-152">In Outlook on Windows, create a new message.</span></span>

    ![Outlook on Windows のメッセージ ウィンドウと新規作成時に件名が設定されているスクリーンショット](../images/outlook-win-autolaunch.png)

## <a name="event-based-activation-behavior-and-limitations"></a><span data-ttu-id="c05ea-154">イベント ベースのアクティブ化の動作と制限事項</span><span class="sxs-lookup"><span data-stu-id="c05ea-154">Event-based activation behavior and limitations</span></span>

<span data-ttu-id="c05ea-155">イベントに基づいてアクティブ化されるアドインは、実行時間が短く、軽量であり、可能な限り非非アクティブである必要があります。</span><span class="sxs-lookup"><span data-stu-id="c05ea-155">Add-ins that activate based on events are expected to be short-running, lightweight, and as non-invasive as possible.</span></span> <span data-ttu-id="c05ea-156">アドインが起動イベントの処理を完了したと示す場合は、アドインでメソッドを呼び出す方法をお勧 `event.completed` めします。</span><span class="sxs-lookup"><span data-stu-id="c05ea-156">To signal that your add-in has completed processing the launch event, we recommend you have your add-in call the `event.completed` method.</span></span> <span data-ttu-id="c05ea-157">この呼び出しが行われた場合、アドインは約 300 秒 (イベント ベースのアドインの実行に許容される最大時間) 内にタイム アウトします。ユーザーが作成ウィンドウを閉じると、アドインも終了します。</span><span class="sxs-lookup"><span data-stu-id="c05ea-157">If that call is not made, the add-in will time out within approximately 300 seconds, the maximum length of time allowed for running event-based add-ins. The add-in also ends when the user closes the compose window.</span></span>

<span data-ttu-id="c05ea-158">ユーザーが同じイベントにサブスクライブした複数のアドインを持っている場合、Outlook プラットフォームは特定の順序でアドインを起動します。</span><span class="sxs-lookup"><span data-stu-id="c05ea-158">If the user has multiple add-ins that subscribed to the same event, the Outlook platform launches the add-ins in no particular order.</span></span> <span data-ttu-id="c05ea-159">現在、アクティブに実行できるイベント ベースのアドインは 5 つのみです。</span><span class="sxs-lookup"><span data-stu-id="c05ea-159">Currently, only five event-based add-ins can be actively running.</span></span> <span data-ttu-id="c05ea-160">追加のアドインはキューにプッシュされ、以前にアクティブだったアドインが完了または非アクティブ化されると実行されます。</span><span class="sxs-lookup"><span data-stu-id="c05ea-160">Any additional add-ins are pushed to a queue then run as previously active add-ins are completed or deactivated.</span></span>

<span data-ttu-id="c05ea-161">ユーザーは、アドインの実行を開始した現在のメール アイテムから切り替えまたは移動できます。</span><span class="sxs-lookup"><span data-stu-id="c05ea-161">The user can switch or navigate away from the current mail item where the add-in started running.</span></span> <span data-ttu-id="c05ea-162">起動されたアドインは、バックグラウンドで操作を完了します。</span><span class="sxs-lookup"><span data-stu-id="c05ea-162">The add-in that was launched will finish its operation in the background.</span></span>

<span data-ttu-id="c05ea-163">UI Office.js変更する API の一部は、イベント ベースのアドインでは許可されません。ブロックされる API を次に示します。</span><span class="sxs-lookup"><span data-stu-id="c05ea-163">Some Office.js APIs that change or alter the UI are not allowed from event-based add-ins. The following are the blocked APIs:</span></span>

- <span data-ttu-id="c05ea-164">Under `Office.context.mailbox` :</span><span class="sxs-lookup"><span data-stu-id="c05ea-164">Under `Office.context.mailbox`:</span></span>
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- <span data-ttu-id="c05ea-165">Under `Office.context.ui` :</span><span class="sxs-lookup"><span data-stu-id="c05ea-165">Under `Office.context.ui`:</span></span>
  - `displayDialogAsync`
  - `messageParent`
- <span data-ttu-id="c05ea-166">Under `Office.context.auth` :</span><span class="sxs-lookup"><span data-stu-id="c05ea-166">Under `Office.context.auth`:</span></span>
  - `getAccessToken`
  - `getAccessTokenAsync`

## <a name="see-also"></a><span data-ttu-id="c05ea-167">関連項目</span><span class="sxs-lookup"><span data-stu-id="c05ea-167">See also</span></span>

[<span data-ttu-id="c05ea-168">Outlook アドインのマニフェスト</span><span class="sxs-lookup"><span data-stu-id="c05ea-168">Outlook add-in manifests</span></span>](manifests.md)