---
title: イベント ベースのアクティブ化用に Outlook アドインを構成する (プレビュー)
description: イベント ベースのアクティブ化用に Outlook アドインを構成する方法について学習します。
ms.topic: article
ms.date: 01/25/2021
localization_priority: Normal
ms.openlocfilehash: 4790de491b84cfba3b64bfb6c176e7bf1ff42ec7
ms.sourcegitcommit: adbc9d59ffa5efdff5afa9115e0990544f2246ab
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/27/2021
ms.locfileid: "49990506"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a><span data-ttu-id="89748-103">イベント ベースのアクティブ化用に Outlook アドインを構成する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="89748-103">Configure your Outlook add-in for event-based activation (preview)</span></span>

<span data-ttu-id="89748-104">イベント ベースのアクティブ化機能がない場合、ユーザーは自分のタスクを完了するためにアドインを明示的に起動する必要があります。</span><span class="sxs-lookup"><span data-stu-id="89748-104">Without the event-based activation feature, a user has to explicitly launch an add-in to complete their tasks.</span></span> <span data-ttu-id="89748-105">この機能を使用すると、特にすべてのアイテムに適用される操作に関して、特定のイベントに基づいてタスクを実行できます。</span><span class="sxs-lookup"><span data-stu-id="89748-105">This feature enables your add-in to run tasks based on certain events, particularly for operations that apply to every item.</span></span> <span data-ttu-id="89748-106">作業ウィンドウと UI を使用する機能を統合することもできます。</span><span class="sxs-lookup"><span data-stu-id="89748-106">You can also integrate with the task pane and UI-less functionality.</span></span> <span data-ttu-id="89748-107">現時点では、次のイベントがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="89748-107">At present, the following events are supported.</span></span>

- <span data-ttu-id="89748-108">`OnNewMessageCompose`: 新しいメッセージの作成時 (返信、全員に返信、転送を含む)</span><span class="sxs-lookup"><span data-stu-id="89748-108">`OnNewMessageCompose`: On composing a new message (includes reply, reply all, and forward)</span></span>
- <span data-ttu-id="89748-109">`OnNewAppointmentOrganizer`: 新しい予定の作成時</span><span class="sxs-lookup"><span data-stu-id="89748-109">`OnNewAppointmentOrganizer`: On creating a new appointment</span></span>

  > [!IMPORTANT]
  > <span data-ttu-id="89748-110">この機能は **、下** 書きや既存の予定など、アイテムの編集時にはアクティブ化されない。</span><span class="sxs-lookup"><span data-stu-id="89748-110">This feature does **not** activate on editing an item, for example, a draft or an existing appointment.</span></span>

<span data-ttu-id="89748-111">このチュートリアルの終わりまでに、新しいメッセージが作成されるたびに実行されるアドインが作成されます。</span><span class="sxs-lookup"><span data-stu-id="89748-111">By the end of this walkthrough, you'll have an add-in that runs whenever a new message is created.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="89748-112">この機能は、Microsoft 365 サブスクリプションを使用した Outlook on the web でのプレビューでのみサポートされます。 [](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)</span><span class="sxs-lookup"><span data-stu-id="89748-112">This feature is only supported for [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web with a Microsoft 365 subscription.</span></span> <span data-ttu-id="89748-113">詳細 [については、この記事のイベント ベースのアクティブ化](#how-to-preview-the-event-based-activation-feature) 機能をプレビューする方法を参照してください。</span><span class="sxs-lookup"><span data-stu-id="89748-113">See [How to preview the event-based activation feature](#how-to-preview-the-event-based-activation-feature) in this article for more details.</span></span>
>
> <span data-ttu-id="89748-114">プレビュー機能は予告なしに変更されることがありますので、実稼働アドインでは使用できません。</span><span class="sxs-lookup"><span data-stu-id="89748-114">Because preview features are subject to change without notice, they shouldn't be used in production add-ins.</span></span>

## <a name="how-to-preview-the-event-based-activation-feature"></a><span data-ttu-id="89748-115">イベント ベースのアクティブ化機能をプレビューする方法</span><span class="sxs-lookup"><span data-stu-id="89748-115">How to preview the event-based activation feature</span></span>

<span data-ttu-id="89748-116">イベント ベースのアクティブ化機能をお試しください。</span><span class="sxs-lookup"><span data-stu-id="89748-116">We invite you to try out the event-based activation feature!</span></span> <span data-ttu-id="89748-117">GitHub を通じてフィードバックを提供することで、お客様のシナリオと改善方法をお知らせします (このページの最後にあるフィードバックセクションをご覧ください)。</span><span class="sxs-lookup"><span data-stu-id="89748-117">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="89748-118">この機能をプレビューするには、次の方法を使用します。</span><span class="sxs-lookup"><span data-stu-id="89748-118">To preview this feature:</span></span>

- <span data-ttu-id="89748-119">CDN で **ベータ** ライブラリを参照する ( https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) .</span><span class="sxs-lookup"><span data-stu-id="89748-119">Reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="89748-120">TypeScript [のコンパイルと](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) 読み取りIntelliSense定義ファイルは CDN と [DefinitelyTyped で見つかりました](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts)。</span><span class="sxs-lookup"><span data-stu-id="89748-120">The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span></span> <span data-ttu-id="89748-121">次の種類を使用してインストールできます `npm install --save-dev @types/office-js-preview` 。</span><span class="sxs-lookup"><span data-stu-id="89748-121">You can install these types with `npm install --save-dev @types/office-js-preview`.</span></span>
- <span data-ttu-id="89748-122">[Microsoft 365 テナントで対象指定リリースを構成します](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)。</span><span class="sxs-lookup"><span data-stu-id="89748-122">[Configure targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="89748-123">環境を設定する</span><span class="sxs-lookup"><span data-stu-id="89748-123">Set up your environment</span></span>

<span data-ttu-id="89748-124">Outlook クイック [スタートを完了](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) します。このクイック スタートでは、アドイン用の Yeoman ジェネレーターを使用してアドイン Office作成します。</span><span class="sxs-lookup"><span data-stu-id="89748-124">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="89748-125">マニフェストを構成する</span><span class="sxs-lookup"><span data-stu-id="89748-125">Configure the manifest</span></span>

<span data-ttu-id="89748-126">アドインのイベント ベースのアクティブ化を有効にするには、マニフェストで [Runtimes](../reference/manifest/runtimes.md) 要素と [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) 拡張ポイントを構成する必要があります。</span><span class="sxs-lookup"><span data-stu-id="89748-126">To enable event-based activation of your add-in, you must configure the [Runtimes](../reference/manifest/runtimes.md) element and [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) extension point in the manifest.</span></span> <span data-ttu-id="89748-127">今のところ、 `DesktopFormFactor` サポートされているフォーム ファクターは 1 つのみです。</span><span class="sxs-lookup"><span data-stu-id="89748-127">For now, `DesktopFormFactor` is the only supported form factor.</span></span>

1. <span data-ttu-id="89748-128">コード エディターで、クイック スタート プロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="89748-128">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="89748-129">プロジェクトの **manifest.xml** にある新しいファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="89748-129">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="89748-130">ノード全体 `<VersionOverrides>` (開いているタグと閉じるタグを含む) を選択し、次の XML に置き換えてください。</span><span class="sxs-lookup"><span data-stu-id="89748-130">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML.</span></span>

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
        <bt:Url id="JSRuntime.Url" DefaultValue="https://localhost:3000/runtime.js" />
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

<span data-ttu-id="89748-131">Outlook on Windows では JavaScript ファイルを使用し、Outlook on the web は同じ JavaScript ファイルを参照する HTML ファイルを使用します。</span><span class="sxs-lookup"><span data-stu-id="89748-131">Outlook on Windows uses a JavaScript file, while Outlook on the web uses an HTML file that references the same JavaScript file.</span></span> <span data-ttu-id="89748-132">Outlook プラットフォームは最終的に Outlook クライアントに基づいて HTML または JavaScript を使用するかどうかを決定しますので、マニフェストでこれらの両方のファイルへの参照を提供する必要があります。</span><span class="sxs-lookup"><span data-stu-id="89748-132">You must provide references to both these files in the manifest as the Outlook platform ultimately determines whether to use HTML or JavaScript based on the Outlook client.</span></span> <span data-ttu-id="89748-133">そのため、イベント処理を構成するには、要素内の HTML の場所を指定し、その子要素で、HTML によってインライン化または参照される JavaScript ファイルの場所を提供します `Runtime` `Override` 。</span><span class="sxs-lookup"><span data-stu-id="89748-133">As such, to configure event handling, provide the location of the HTML in the `Runtime` element, then in its `Override` child element provide the location of the JavaScript file inlined or referenced by the HTML.</span></span>

> [!TIP]
> <span data-ttu-id="89748-134">Outlook アドインのマニフェストの詳細については、Outlook アドインの [マニフェストを参照してください](manifests.md)。</span><span class="sxs-lookup"><span data-stu-id="89748-134">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-event-handling"></a><span data-ttu-id="89748-135">イベント処理を実装する</span><span class="sxs-lookup"><span data-stu-id="89748-135">Implement event handling</span></span>

<span data-ttu-id="89748-136">選択したイベントの処理を実装する必要があります。</span><span class="sxs-lookup"><span data-stu-id="89748-136">You have to implement handling for your selected events.</span></span>

<span data-ttu-id="89748-137">このシナリオでは、新しいアイテムを作成する処理を追加します。</span><span class="sxs-lookup"><span data-stu-id="89748-137">In this scenario, you'll add handling for composing new items.</span></span>

1. <span data-ttu-id="89748-138">同じクイック スタート プロジェクトから、コード エディターで **ファイル ./src/commands/commands.js** を開きます。</span><span class="sxs-lookup"><span data-stu-id="89748-138">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="89748-139">関数の `action` 後に、次の JavaScript 関数を挿入します。</span><span class="sxs-lookup"><span data-stu-id="89748-139">After the `action` function, insert the following JavaScript functions.</span></span>

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

1. <span data-ttu-id="89748-140">ファイルの最後に、次のステートメントを追加します。</span><span class="sxs-lookup"><span data-stu-id="89748-140">At the end of the file, add the following statements.</span></span>

    ```js
    g.onMessageComposeHandler = onMessageComposeHandler;
    g.onAppointmentComposeHandler = onAppointmentComposeHandler;
    ```

## <a name="try-it-out"></a><span data-ttu-id="89748-141">試してみる</span><span class="sxs-lookup"><span data-stu-id="89748-141">Try it out</span></span>

1. <span data-ttu-id="89748-142">プロジェクトのルート ディレクトリから次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="89748-142">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="89748-143">このコマンドを実行すると、ローカル Web サーバーが起動します (まだ実行されていない場合)。</span><span class="sxs-lookup"><span data-stu-id="89748-143">When you run this command, the local web server will start (if it's not already running).</span></span>

    ```command&nbsp;line
    npm run dev-server
    ```

1. <span data-ttu-id="89748-144">「[テスト用に Outlook アドインをサイドロードする](sideload-outlook-add-ins-for-testing.md)」の手順に従って Outlook アドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="89748-144">Follow the instructions in [Sideload Outlook add-ins for testing](sideload-outlook-add-ins-for-testing.md) to sideload the add-in in Outlook.</span></span>

1. <span data-ttu-id="89748-145">Outlook on the web で新しいメッセージを作成します。</span><span class="sxs-lookup"><span data-stu-id="89748-145">In Outlook on the web, create a new message.</span></span>

    ![Outlook on the web のメッセージ ウィンドウのスクリーンショット。件名が新規作成に設定されています。](../images/outlook-web-autolaunch.png)

## <a name="event-based-activation-behavior-and-limitations"></a><span data-ttu-id="89748-147">イベント ベースのアクティブ化の動作と制限事項</span><span class="sxs-lookup"><span data-stu-id="89748-147">Event-based activation behavior and limitations</span></span>

<span data-ttu-id="89748-148">イベントに基づいてアクティブ化するアドインは、最大で約 300 秒、実行時間が短いアドインとして設計されています。</span><span class="sxs-lookup"><span data-stu-id="89748-148">Add-ins that activate based on events are designed to be short-running, up to approximately 300 seconds.</span></span> <span data-ttu-id="89748-149">起動イベントの処理が完了したメソッドを呼び出す `event.completed` アドインをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="89748-149">We recommend you have your add-in call the `event.completed` method to signal it has completed processing the launch event.</span></span> <span data-ttu-id="89748-150">ユーザーが作成ウィンドウを閉じると、アドインも終了します。</span><span class="sxs-lookup"><span data-stu-id="89748-150">The add-in also ends when the user closes the compose window.</span></span>

<span data-ttu-id="89748-151">ユーザーが同じイベントにサブスクライブした複数のアドインを持っている場合、Outlook プラットフォームは特定の順序でアドインを起動します。</span><span class="sxs-lookup"><span data-stu-id="89748-151">If the user has multiple add-ins that subscribed to the same event, the Outlook platform launches the add-ins in no particular order.</span></span> <span data-ttu-id="89748-152">現在、アクティブに実行できるイベント ベースのアドインは 5 つのみです。</span><span class="sxs-lookup"><span data-stu-id="89748-152">Currently, only five event-based add-ins can be actively running.</span></span> <span data-ttu-id="89748-153">追加のアドインはキューにプッシュされ、以前にアクティブだったアドインが完了または非アクティブ化されると実行されます。</span><span class="sxs-lookup"><span data-stu-id="89748-153">Any additional add-ins are pushed to a queue then run as previously active add-ins are completed or deactivated.</span></span>

<span data-ttu-id="89748-154">ユーザーは、アドインの実行を開始した現在のメール アイテムから切り替えまたは移動できます。</span><span class="sxs-lookup"><span data-stu-id="89748-154">The user can switch or navigate away from the current mail item where the add-in started running.</span></span> <span data-ttu-id="89748-155">起動されたアドインは、バックグラウンドで操作を完了します。</span><span class="sxs-lookup"><span data-stu-id="89748-155">The add-in that was launched will finish its operation in the background.</span></span>

<span data-ttu-id="89748-156">UI Office.js変更する API の一部は、イベント ベースのアドインでは許可されません。ブロックされる API を次に示します。</span><span class="sxs-lookup"><span data-stu-id="89748-156">Some Office.js APIs that change or alter the UI are not allowed from event-based add-ins. The following are the blocked APIs.</span></span>

- <span data-ttu-id="89748-157">Under `Office.context.mailbox` :</span><span class="sxs-lookup"><span data-stu-id="89748-157">Under `Office.context.mailbox`:</span></span>
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- <span data-ttu-id="89748-158">Under `Office.context.ui` :</span><span class="sxs-lookup"><span data-stu-id="89748-158">Under `Office.context.ui`:</span></span>
  - `displayDialogAsync`
  - `messageParent`
- <span data-ttu-id="89748-159">Under `Office.context.auth` :</span><span class="sxs-lookup"><span data-stu-id="89748-159">Under `Office.context.auth`:</span></span>
  - `getAccessToken`
  - `getAccessTokenAsync`

## <a name="see-also"></a><span data-ttu-id="89748-160">関連項目</span><span class="sxs-lookup"><span data-stu-id="89748-160">See also</span></span>

[<span data-ttu-id="89748-161">Outlook アドインのマニフェスト</span><span class="sxs-lookup"><span data-stu-id="89748-161">Outlook add-in manifests</span></span>](manifests.md)
