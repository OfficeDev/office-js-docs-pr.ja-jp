---
title: イベントベースのライセンス認証用に Outlook アドインを構成する (プレビュー)
description: イベントベースのライセンス認証用に Outlook アドインを構成する方法について説明します。
ms.topic: article
ms.date: 05/22/2020
localization_priority: Normal
ms.openlocfilehash: 73cdd4949b870d9bc5a5ad2006ce2081575558df
ms.sourcegitcommit: 77617f6ad06e07f5ff8078b26301748f73e2ee01
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/29/2020
ms.locfileid: "44413197"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a><span data-ttu-id="a55f4-103">イベントベースのライセンス認証用に Outlook アドインを構成する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="a55f4-103">Configure your Outlook add-in for event-based activation (preview)</span></span>

<span data-ttu-id="a55f4-104">イベントベースのアクティブ化機能がない場合、ユーザーは、自分のタスクを完了するためにアドインを明示的に起動する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a55f4-104">Without the event-based activation feature, a user has to explicitly launch an add-in to complete their tasks.</span></span> <span data-ttu-id="a55f4-105">この機能により、特定のイベント (特に、すべてのアイテムに適用される操作) に基づいてアドインでタスクを実行できます。</span><span class="sxs-lookup"><span data-stu-id="a55f4-105">This feature enables your add-in to run tasks based on certain events, particularly for operations that apply to every item.</span></span> <span data-ttu-id="a55f4-106">作業ウィンドウと UI 非表示機能に統合することもできます。</span><span class="sxs-lookup"><span data-stu-id="a55f4-106">You can also integrate with the task pane and UI-less functionality.</span></span> <span data-ttu-id="a55f4-107">現時点では、サポートされているイベントは次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="a55f4-107">At present, the supported events are as follows.</span></span>

- <span data-ttu-id="a55f4-108">`OnNewMessageCompose`: 新しいメッセージを作成するときに (返信、全員に返信、および転送を含む)</span><span class="sxs-lookup"><span data-stu-id="a55f4-108">`OnNewMessageCompose`: On composing a new message (includes reply, reply all, and forward)</span></span>
- <span data-ttu-id="a55f4-109">`OnNewAppointmentOrganizer`: 新しい予定を作成するとき</span><span class="sxs-lookup"><span data-stu-id="a55f4-109">`OnNewAppointmentOrganizer`: On creating a new appointment</span></span>

<span data-ttu-id="a55f4-110">このチュートリアルを終了すると、新しいメッセージが作成されるたびに実行されるアドインができます。</span><span class="sxs-lookup"><span data-stu-id="a55f4-110">By the end of this walkthrough, you'll have an add-in that runs whenever a new message is created.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="a55f4-111">この機能は、Office 365 サブスクリプションを使用する web 上の Outlook の[プレビュー](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="a55f4-111">This feature is only supported for [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web with an Office 365 subscription.</span></span> <span data-ttu-id="a55f4-112">詳細については、この記事の「[イベントに基づくライセンス認証機能をプレビューする方法](#how-to-preview-the-event-based-activation-feature)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a55f4-112">See [How to preview the event-based activation feature](#how-to-preview-the-event-based-activation-feature) in this article for more details.</span></span>
>
> <span data-ttu-id="a55f4-113">プレビュー機能は予告なしに変更される可能性があるため、運用アドインでは使用しないでください。</span><span class="sxs-lookup"><span data-stu-id="a55f4-113">Because preview features are subject to change without notice, they shouldn't be used in production add-ins.</span></span>

## <a name="how-to-preview-the-event-based-activation-feature"></a><span data-ttu-id="a55f4-114">イベントベースのライセンス認証機能をプレビューする方法</span><span class="sxs-lookup"><span data-stu-id="a55f4-114">How to preview the event-based activation feature</span></span>

<span data-ttu-id="a55f4-115">イベントに基づくライセンス認証機能をお試しください。</span><span class="sxs-lookup"><span data-stu-id="a55f4-115">We invite you to try out the event-based activation feature!</span></span> <span data-ttu-id="a55f4-116">GitHub を通じてフィードバックを提供することによって、自分のシナリオと改善方法をお知らせください (このページの最後にある**フィードバック**セクションを参照してください)。</span><span class="sxs-lookup"><span data-stu-id="a55f4-116">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="a55f4-117">この機能をプレビューするには:</span><span class="sxs-lookup"><span data-stu-id="a55f4-117">To preview this feature:</span></span>

- <span data-ttu-id="a55f4-118">CDN の**ベータ版**ライブラリを参照し https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) ます (。</span><span class="sxs-lookup"><span data-stu-id="a55f4-118">Reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="a55f4-119">TypeScript のコンパイルおよび IntelliSense 用の[型定義ファイル](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)は、CDN と、定義[された](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts)定義ファイルにあります。</span><span class="sxs-lookup"><span data-stu-id="a55f4-119">The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span></span> <span data-ttu-id="a55f4-120">これらの種類は、でインストールでき `npm install --save-dev @types/office-js-preview` ます。</span><span class="sxs-lookup"><span data-stu-id="a55f4-120">You can install these types with `npm install --save-dev @types/office-js-preview`.</span></span>
- <span data-ttu-id="a55f4-121">[この要求フォーム](https://aka.ms/OWAPreview)を完了して送信することにより、Microsoft 365 アカウントを使用して、web 上の Outlook のプレビュービットへのアクセスを要求します。</span><span class="sxs-lookup"><span data-stu-id="a55f4-121">Request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this request form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="a55f4-122">テナントの準備が整ったことをお知らせします。</span><span class="sxs-lookup"><span data-stu-id="a55f4-122">We'll let you know when your tenant is ready.</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="a55f4-123">環境を設定する</span><span class="sxs-lookup"><span data-stu-id="a55f4-123">Set up your environment</span></span>

<span data-ttu-id="a55f4-124">Outlook の[クイックスタート](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator)に記入します。このアドインプロジェクトは、Office アドイン用の [アプリ] ジェネレーターを使用して作成されます。</span><span class="sxs-lookup"><span data-stu-id="a55f4-124">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="a55f4-125">マニフェストを構成する</span><span class="sxs-lookup"><span data-stu-id="a55f4-125">Configure the manifest</span></span>

<span data-ttu-id="a55f4-126">アドインのイベントベースのアクティブ化を有効にするには、マニフェストで、[ランタイム](../reference/manifest/runtimes.md)要素と[launchevent](../reference/manifest/extensionpoint.md#launchevent-preview)拡張点を構成する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a55f4-126">To enable event-based activation of your add-in, you must configure the [Runtimes](../reference/manifest/runtimes.md) element and [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) extension point in the manifest.</span></span> <span data-ttu-id="a55f4-127">ここで `DesktopFormFactor` は、サポートされているフォームファクターのみを示します。</span><span class="sxs-lookup"><span data-stu-id="a55f4-127">For now, `DesktopFormFactor` is the only supported form factor.</span></span>

1. <span data-ttu-id="a55f4-128">コードエディターで、[クイックスタート] プロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="a55f4-128">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="a55f4-129">プロジェクトのルートにある**manifest.xml**ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="a55f4-129">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="a55f4-130">`<VersionOverrides>`ノード全体 (open タグと close タグを含む) を選択し、次の XML に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="a55f4-130">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML.</span></span>

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

<span data-ttu-id="a55f4-131">Outlook on the Windows は JavaScript ファイルを使用しますが、web 上の Outlook は同じ JavaScript ファイルを参照する HTML ファイルを使用します。</span><span class="sxs-lookup"><span data-stu-id="a55f4-131">Outlook on Windows uses a JavaScript file, while Outlook on the web uses an HTML file that references the same JavaScript file.</span></span> <span data-ttu-id="a55f4-132">Outlook プラットフォームは、outlook クライアントに基づいて HTML と JavaScript のどちらを使用するかを決定するために、これらのファイルへの参照をマニフェストに提供する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a55f4-132">You must provide references to both these files in the manifest as the Outlook platform ultimately determines whether to use HTML or JavaScript based on the Outlook client.</span></span> <span data-ttu-id="a55f4-133">そのため、イベント処理を構成するには、要素内の HTML の場所を指定し、 `Runtime` その `Override` 子要素で、JavaScript ファイルの場所を指定するか、html で参照します。</span><span class="sxs-lookup"><span data-stu-id="a55f4-133">As such, to configure event handling, provide the location of the HTML in the `Runtime` element, then in its `Override` child element provide the location of the JavaScript file inlined or referenced by the HTML.</span></span>

> [!TIP]
> <span data-ttu-id="a55f4-134">Outlook アドインのマニフェストの詳細については、「 [outlook アドインのマニフェスト](manifests.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="a55f4-134">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-event-handling"></a><span data-ttu-id="a55f4-135">イベント処理を実装する</span><span class="sxs-lookup"><span data-stu-id="a55f4-135">Implement event handling</span></span>

<span data-ttu-id="a55f4-136">選択したイベントの処理を実装する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a55f4-136">You have to implement handling for your selected events.</span></span>

<span data-ttu-id="a55f4-137">このシナリオでは、新しいアイテムを作成するための処理を追加します。</span><span class="sxs-lookup"><span data-stu-id="a55f4-137">In this scenario, you'll add handling for composing new items.</span></span>

1. <span data-ttu-id="a55f4-138">同じクイックスタートプロジェクトから、コードエディターで **/src/commands/commands.js**を開きます。</span><span class="sxs-lookup"><span data-stu-id="a55f4-138">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="a55f4-139">関数の後 `action` に、次の JavaScript 関数を挿入します。</span><span class="sxs-lookup"><span data-stu-id="a55f4-139">After the `action` function, insert the following JavaScript functions.</span></span>

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

1. <span data-ttu-id="a55f4-140">ファイルの末尾に、次のステートメントを追加します。</span><span class="sxs-lookup"><span data-stu-id="a55f4-140">At the end of the file, add the following statements.</span></span>

    ```js
    g.onMessageComposeHandler = onMessageComposeHandler;
    g.onAppointmentComposeHandler = onAppointmentComposeHandler;
    ```

## <a name="try-it-out"></a><span data-ttu-id="a55f4-141">試してみる</span><span class="sxs-lookup"><span data-stu-id="a55f4-141">Try it out</span></span>

1. <span data-ttu-id="a55f4-142">プロジェクトのルート ディレクトリから次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="a55f4-142">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="a55f4-143">このコマンドを実行すると、ローカル Web サーバーが起動します (まだ実行されていない場合)。</span><span class="sxs-lookup"><span data-stu-id="a55f4-143">When you run this command, the local web server will start (if it's not already running).</span></span>

    ```command&nbsp;line
    npm run dev-server
    ```

1. <span data-ttu-id="a55f4-144">「[テスト用に Outlook アドインをサイドロードする](sideload-outlook-add-ins-for-testing.md)」の手順に従って Outlook アドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="a55f4-144">Follow the instructions in [Sideload Outlook add-ins for testing](sideload-outlook-add-ins-for-testing.md) to sideload the add-in in Outlook.</span></span>

1. <span data-ttu-id="a55f4-145">Outlook on the web で新しいメッセージを作成します。</span><span class="sxs-lookup"><span data-stu-id="a55f4-145">In Outlook on the web, create a new message.</span></span>

    ![Web 上の Outlook で、作成時に件名が設定されたメッセージウィンドウのスクリーンショット。](../images/outlook-web-autolaunch.png)

## <a name="event-based-activation-behavior-and-limitations"></a><span data-ttu-id="a55f4-147">イベントに基づくアクティブ化の動作と制限事項</span><span class="sxs-lookup"><span data-stu-id="a55f4-147">Event-based activation behavior and limitations</span></span>

<span data-ttu-id="a55f4-148">イベントに基づいてアクティブ化するアドインは、短時間で最大330秒のみが実行されるように設計されています。</span><span class="sxs-lookup"><span data-stu-id="a55f4-148">Add-ins that activate based on events are designed to be short-running, up to 330 seconds only.</span></span> <span data-ttu-id="a55f4-149">アドインで、 `event.completed` 起動イベントの処理が完了したことを通知するメソッドを呼び出すことをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="a55f4-149">We recommend you have your add-in call the `event.completed` method to signal it has completed processing the launch event.</span></span> <span data-ttu-id="a55f4-150">ユーザーが [新規作成] ウィンドウを閉じたときにもアドインは終了します。</span><span class="sxs-lookup"><span data-stu-id="a55f4-150">The add-in also ends when the user closes the compose window.</span></span>

<span data-ttu-id="a55f4-151">ユーザーが同じイベントにサブスクライブした複数のアドインを持っている場合、Outlook プラットフォームは特定の順序でアドインを起動しません。</span><span class="sxs-lookup"><span data-stu-id="a55f4-151">If the user has multiple add-ins that subscribed to the same event, the Outlook platform launches the add-ins in no particular order.</span></span> <span data-ttu-id="a55f4-152">現在、5つのイベントベースのアドインのみをアクティブに実行できます。</span><span class="sxs-lookup"><span data-stu-id="a55f4-152">Currently, only five event-based add-ins can be actively running.</span></span> <span data-ttu-id="a55f4-153">追加のアドインは、キューにプッシュされ、以前のアクティブなアドインが完了または非アクティブになったときとして実行されます。</span><span class="sxs-lookup"><span data-stu-id="a55f4-153">Any additional add-ins are pushed to a queue then run as previously active add-ins are completed or deactivated.</span></span>

<span data-ttu-id="a55f4-154">ユーザーは、アドインの実行が開始された現在のメールアイテムから切り替えることができます。</span><span class="sxs-lookup"><span data-stu-id="a55f4-154">The user can switch or navigate away from the current mail item where the add-in started running.</span></span> <span data-ttu-id="a55f4-155">起動されたアドインは、バックグラウンドで操作を終了します。</span><span class="sxs-lookup"><span data-stu-id="a55f4-155">The add-in that was launched will finish its operation in the background.</span></span>

<span data-ttu-id="a55f4-156">UI を変更または変更する一部の Office .js Api は、イベントベースのアドインからは許可されません。ブロックされる Api を次に示します。</span><span class="sxs-lookup"><span data-stu-id="a55f4-156">Some Office.js APIs that change or alter the UI are not allowed from event-based add-ins. The following are the blocked APIs.</span></span>

- <span data-ttu-id="a55f4-157">`Office.context.mailbox`以下:</span><span class="sxs-lookup"><span data-stu-id="a55f4-157">Under `Office.context.mailbox`:</span></span>
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- <span data-ttu-id="a55f4-158">`Office.context.ui`以下:</span><span class="sxs-lookup"><span data-stu-id="a55f4-158">Under `Office.context.ui`:</span></span>
  - `displayDialogAsync`
  - `messageParent`
- <span data-ttu-id="a55f4-159">`Office.context.auth`以下:</span><span class="sxs-lookup"><span data-stu-id="a55f4-159">Under `Office.context.auth`:</span></span>
  - `getAccessTokenAsync`

## <a name="see-also"></a><span data-ttu-id="a55f4-160">関連項目</span><span class="sxs-lookup"><span data-stu-id="a55f4-160">See also</span></span>

[<span data-ttu-id="a55f4-161">Outlook アドインのマニフェスト</span><span class="sxs-lookup"><span data-stu-id="a55f4-161">Outlook add-in manifests</span></span>](manifests.md)
