---
title: イベントベースのライセンス認証用に Outlook アドインを構成する (プレビュー)
description: イベントベースのライセンス認証用に Outlook アドインを構成する方法について説明します。
ms.topic: article
ms.date: 08/24/2020
localization_priority: Normal
ms.openlocfilehash: 0131cafa8315315d63b6319ecad4fd41b1168073
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293927"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a><span data-ttu-id="b728e-103">イベントベースのライセンス認証用に Outlook アドインを構成する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="b728e-103">Configure your Outlook add-in for event-based activation (preview)</span></span>

<span data-ttu-id="b728e-104">イベントベースのアクティブ化機能がない場合、ユーザーは、自分のタスクを完了するためにアドインを明示的に起動する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b728e-104">Without the event-based activation feature, a user has to explicitly launch an add-in to complete their tasks.</span></span> <span data-ttu-id="b728e-105">この機能により、特定のイベント (特に、すべてのアイテムに適用される操作) に基づいてアドインでタスクを実行できます。</span><span class="sxs-lookup"><span data-stu-id="b728e-105">This feature enables your add-in to run tasks based on certain events, particularly for operations that apply to every item.</span></span> <span data-ttu-id="b728e-106">作業ウィンドウと UI 非表示機能に統合することもできます。</span><span class="sxs-lookup"><span data-stu-id="b728e-106">You can also integrate with the task pane and UI-less functionality.</span></span> <span data-ttu-id="b728e-107">現在、次のイベントがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="b728e-107">At present, the following events are supported.</span></span>

- <span data-ttu-id="b728e-108">`OnNewMessageCompose`: 新しいメッセージを作成するときに (返信、全員に返信、および転送を含む)</span><span class="sxs-lookup"><span data-stu-id="b728e-108">`OnNewMessageCompose`: On composing a new message (includes reply, reply all, and forward)</span></span>
- <span data-ttu-id="b728e-109">`OnNewAppointmentOrganizer`: 新しい予定を作成するとき</span><span class="sxs-lookup"><span data-stu-id="b728e-109">`OnNewAppointmentOrganizer`: On creating a new appointment</span></span>

  > [!IMPORTANT]
  > <span data-ttu-id="b728e-110">この機能は、アイテムを編集するときにアクティブ化 **されません** 。たとえば、下書きまたは既存の予定です。</span><span class="sxs-lookup"><span data-stu-id="b728e-110">This feature does **not** activate on editing an item, for example, a draft or an existing appointment.</span></span>

<span data-ttu-id="b728e-111">このチュートリアルを終了すると、新しいメッセージが作成されるたびに実行されるアドインができます。</span><span class="sxs-lookup"><span data-stu-id="b728e-111">By the end of this walkthrough, you'll have an add-in that runs whenever a new message is created.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b728e-112">この機能は、Microsoft 365 サブスクリプションを使用する web 上の Outlook の [プレビュー](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="b728e-112">This feature is only supported for [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web with a Microsoft 365 subscription.</span></span> <span data-ttu-id="b728e-113">詳細については、この記事の「 [イベントに基づくライセンス認証機能をプレビューする方法](#how-to-preview-the-event-based-activation-feature) 」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b728e-113">See [How to preview the event-based activation feature](#how-to-preview-the-event-based-activation-feature) in this article for more details.</span></span>
>
> <span data-ttu-id="b728e-114">プレビュー機能は予告なしに変更される可能性があるため、運用アドインでは使用しないでください。</span><span class="sxs-lookup"><span data-stu-id="b728e-114">Because preview features are subject to change without notice, they shouldn't be used in production add-ins.</span></span>

## <a name="how-to-preview-the-event-based-activation-feature"></a><span data-ttu-id="b728e-115">イベントベースのライセンス認証機能をプレビューする方法</span><span class="sxs-lookup"><span data-stu-id="b728e-115">How to preview the event-based activation feature</span></span>

<span data-ttu-id="b728e-116">イベントに基づくライセンス認証機能をお試しください。</span><span class="sxs-lookup"><span data-stu-id="b728e-116">We invite you to try out the event-based activation feature!</span></span> <span data-ttu-id="b728e-117">GitHub を通じてフィードバックを提供することによって、自分のシナリオと改善方法をお知らせください (このページの最後にある **フィードバック** セクションを参照してください)。</span><span class="sxs-lookup"><span data-stu-id="b728e-117">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="b728e-118">この機能をプレビューするには:</span><span class="sxs-lookup"><span data-stu-id="b728e-118">To preview this feature:</span></span>

- <span data-ttu-id="b728e-119">CDN の **ベータ版** ライブラリを参照し https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) ます (。</span><span class="sxs-lookup"><span data-stu-id="b728e-119">Reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="b728e-120">TypeScript のコンパイルおよび IntelliSense 用の [型定義ファイル](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) は、CDN と、定義 [された](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts)定義ファイルにあります。</span><span class="sxs-lookup"><span data-stu-id="b728e-120">The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span></span> <span data-ttu-id="b728e-121">これらの種類は、でインストールでき `npm install --save-dev @types/office-js-preview` ます。</span><span class="sxs-lookup"><span data-stu-id="b728e-121">You can install these types with `npm install --save-dev @types/office-js-preview`.</span></span>
- <span data-ttu-id="b728e-122">[この要求フォーム](https://aka.ms/OWAPreview)を完了して送信することにより、Microsoft 365 アカウントを使用して、web 上の Outlook のプレビュービットへのアクセスを要求します。</span><span class="sxs-lookup"><span data-stu-id="b728e-122">Request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this request form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="b728e-123">テナントの準備が整ったことをお知らせします。</span><span class="sxs-lookup"><span data-stu-id="b728e-123">We'll let you know when your tenant is ready.</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="b728e-124">環境を設定する</span><span class="sxs-lookup"><span data-stu-id="b728e-124">Set up your environment</span></span>

<span data-ttu-id="b728e-125">Outlook の [クイックスタート](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) に記入します。このアドインプロジェクトは、Office アドイン用の [アプリ] ジェネレーターを使用して作成されます。</span><span class="sxs-lookup"><span data-stu-id="b728e-125">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="b728e-126">マニフェストを構成する</span><span class="sxs-lookup"><span data-stu-id="b728e-126">Configure the manifest</span></span>

<span data-ttu-id="b728e-127">アドインのイベントベースのアクティブ化を有効にするには、マニフェストで、 [ランタイム](../reference/manifest/runtimes.md) 要素と [launchevent](../reference/manifest/extensionpoint.md#launchevent-preview) 拡張点を構成する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b728e-127">To enable event-based activation of your add-in, you must configure the [Runtimes](../reference/manifest/runtimes.md) element and [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) extension point in the manifest.</span></span> <span data-ttu-id="b728e-128">ここで `DesktopFormFactor` は、サポートされているフォームファクターのみを示します。</span><span class="sxs-lookup"><span data-stu-id="b728e-128">For now, `DesktopFormFactor` is the only supported form factor.</span></span>

1. <span data-ttu-id="b728e-129">コードエディターで、[クイックスタート] プロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="b728e-129">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="b728e-130">プロジェクトのルートにある **manifest.xml** ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="b728e-130">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="b728e-131">`<VersionOverrides>`ノード全体 (open タグと close タグを含む) を選択し、次の XML に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="b728e-131">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML.</span></span>

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

<span data-ttu-id="b728e-132">Outlook on the Windows は JavaScript ファイルを使用しますが、web 上の Outlook は同じ JavaScript ファイルを参照する HTML ファイルを使用します。</span><span class="sxs-lookup"><span data-stu-id="b728e-132">Outlook on Windows uses a JavaScript file, while Outlook on the web uses an HTML file that references the same JavaScript file.</span></span> <span data-ttu-id="b728e-133">Outlook プラットフォームは、outlook クライアントに基づいて HTML と JavaScript のどちらを使用するかを決定するために、これらのファイルへの参照をマニフェストに提供する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b728e-133">You must provide references to both these files in the manifest as the Outlook platform ultimately determines whether to use HTML or JavaScript based on the Outlook client.</span></span> <span data-ttu-id="b728e-134">そのため、イベント処理を構成するには、要素内の HTML の場所を指定し、 `Runtime` その `Override` 子要素で、JavaScript ファイルの場所を指定するか、html で参照します。</span><span class="sxs-lookup"><span data-stu-id="b728e-134">As such, to configure event handling, provide the location of the HTML in the `Runtime` element, then in its `Override` child element provide the location of the JavaScript file inlined or referenced by the HTML.</span></span>

> [!TIP]
> <span data-ttu-id="b728e-135">Outlook アドインのマニフェストの詳細については、「 [outlook アドインのマニフェスト](manifests.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b728e-135">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-event-handling"></a><span data-ttu-id="b728e-136">イベント処理を実装する</span><span class="sxs-lookup"><span data-stu-id="b728e-136">Implement event handling</span></span>

<span data-ttu-id="b728e-137">選択したイベントの処理を実装する必要があります。</span><span class="sxs-lookup"><span data-stu-id="b728e-137">You have to implement handling for your selected events.</span></span>

<span data-ttu-id="b728e-138">このシナリオでは、新しいアイテムを作成するための処理を追加します。</span><span class="sxs-lookup"><span data-stu-id="b728e-138">In this scenario, you'll add handling for composing new items.</span></span>

1. <span data-ttu-id="b728e-139">同じクイックスタートプロジェクトから、コードエディターで **/src/commands/commands.js** を開きます。</span><span class="sxs-lookup"><span data-stu-id="b728e-139">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="b728e-140">関数の後 `action` に、次の JavaScript 関数を挿入します。</span><span class="sxs-lookup"><span data-stu-id="b728e-140">After the `action` function, insert the following JavaScript functions.</span></span>

    ```js
    function onMessageComposeHandler(event) {
      setSubject();
      event.completed();
    }
    function onAppointmentComposeHandler(event) {
      setSubject();
      event.completed();
    }
    function setSubject() {
      Office.context.mailbox.item.subject.setAsync("Set by an event-based add-in!");
    }
    ```

1. <span data-ttu-id="b728e-141">ファイルの末尾に、次のステートメントを追加します。</span><span class="sxs-lookup"><span data-stu-id="b728e-141">At the end of the file, add the following statements.</span></span>

    ```js
    g.onMessageComposeHandler = onMessageComposeHandler;
    g.onAppointmentComposeHandler = onAppointmentComposeHandler;
    ```

## <a name="try-it-out"></a><span data-ttu-id="b728e-142">試してみる</span><span class="sxs-lookup"><span data-stu-id="b728e-142">Try it out</span></span>

1. <span data-ttu-id="b728e-143">プロジェクトのルート ディレクトリから次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="b728e-143">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="b728e-144">このコマンドを実行すると、ローカル Web サーバーが起動します (まだ実行されていない場合)。</span><span class="sxs-lookup"><span data-stu-id="b728e-144">When you run this command, the local web server will start (if it's not already running).</span></span>

    ```command&nbsp;line
    npm start
    ```

    > [!IMPORTANT]
    > <span data-ttu-id="b728e-145">"サイドロードがサポートされていません" というエラーが表示された場合は、無視して続行できます。</span><span class="sxs-lookup"><span data-stu-id="b728e-145">If you see a "Sideload is not supported" error, you can ignore it and proceed.</span></span>

1. <span data-ttu-id="b728e-146">「[テスト用に Outlook アドインをサイドロードする](sideload-outlook-add-ins-for-testing.md)」の手順に従って Outlook アドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="b728e-146">Follow the instructions in [Sideload Outlook add-ins for testing](sideload-outlook-add-ins-for-testing.md) to sideload the add-in in Outlook.</span></span>

1. <span data-ttu-id="b728e-147">Outlook on the web で新しいメッセージを作成します。</span><span class="sxs-lookup"><span data-stu-id="b728e-147">In Outlook on the web, create a new message.</span></span>

    ![Web 上の Outlook で、作成時に件名が設定されたメッセージウィンドウのスクリーンショット。](../images/outlook-web-autolaunch.png)

## <a name="event-based-activation-behavior-and-limitations"></a><span data-ttu-id="b728e-149">イベントに基づくアクティブ化の動作と制限事項</span><span class="sxs-lookup"><span data-stu-id="b728e-149">Event-based activation behavior and limitations</span></span>

<span data-ttu-id="b728e-150">イベントに基づいてアクティブ化するアドインは、短時間で最大330秒のみが実行されるように設計されています。</span><span class="sxs-lookup"><span data-stu-id="b728e-150">Add-ins that activate based on events are designed to be short-running, up to 330 seconds only.</span></span> <span data-ttu-id="b728e-151">アドインで、 `event.completed` 起動イベントの処理が完了したことを通知するメソッドを呼び出すことをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="b728e-151">We recommend you have your add-in call the `event.completed` method to signal it has completed processing the launch event.</span></span> <span data-ttu-id="b728e-152">ユーザーが [新規作成] ウィンドウを閉じたときにもアドインは終了します。</span><span class="sxs-lookup"><span data-stu-id="b728e-152">The add-in also ends when the user closes the compose window.</span></span>

<span data-ttu-id="b728e-153">ユーザーが同じイベントにサブスクライブした複数のアドインを持っている場合、Outlook プラットフォームは特定の順序でアドインを起動しません。</span><span class="sxs-lookup"><span data-stu-id="b728e-153">If the user has multiple add-ins that subscribed to the same event, the Outlook platform launches the add-ins in no particular order.</span></span> <span data-ttu-id="b728e-154">現在、5つのイベントベースのアドインのみをアクティブに実行できます。</span><span class="sxs-lookup"><span data-stu-id="b728e-154">Currently, only five event-based add-ins can be actively running.</span></span> <span data-ttu-id="b728e-155">追加のアドインは、キューにプッシュされ、以前のアクティブなアドインが完了または非アクティブになったときとして実行されます。</span><span class="sxs-lookup"><span data-stu-id="b728e-155">Any additional add-ins are pushed to a queue then run as previously active add-ins are completed or deactivated.</span></span>

<span data-ttu-id="b728e-156">ユーザーは、アドインの実行が開始された現在のメールアイテムから切り替えることができます。</span><span class="sxs-lookup"><span data-stu-id="b728e-156">The user can switch or navigate away from the current mail item where the add-in started running.</span></span> <span data-ttu-id="b728e-157">起動されたアドインは、バックグラウンドで操作を終了します。</span><span class="sxs-lookup"><span data-stu-id="b728e-157">The add-in that was launched will finish its operation in the background.</span></span>

<span data-ttu-id="b728e-158">UI を変更または変更する一部の Office.js Api は、イベントベースのアドインからは許可されていません。ブロックされる Api を次に示します。</span><span class="sxs-lookup"><span data-stu-id="b728e-158">Some Office.js APIs that change or alter the UI are not allowed from event-based add-ins. The following are the blocked APIs.</span></span>

- <span data-ttu-id="b728e-159">`Office.context.mailbox`以下:</span><span class="sxs-lookup"><span data-stu-id="b728e-159">Under `Office.context.mailbox`:</span></span>
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- <span data-ttu-id="b728e-160">`Office.context.ui`以下:</span><span class="sxs-lookup"><span data-stu-id="b728e-160">Under `Office.context.ui`:</span></span>
  - `displayDialogAsync`
  - `messageParent`
- <span data-ttu-id="b728e-161">`Office.context.auth`以下:</span><span class="sxs-lookup"><span data-stu-id="b728e-161">Under `Office.context.auth`:</span></span>
  - `getAccessTokenAsync`

## <a name="see-also"></a><span data-ttu-id="b728e-162">関連項目</span><span class="sxs-lookup"><span data-stu-id="b728e-162">See also</span></span>

[<span data-ttu-id="b728e-163">Outlook アドインのマニフェスト</span><span class="sxs-lookup"><span data-stu-id="b728e-163">Outlook add-in manifests</span></span>](manifests.md)
