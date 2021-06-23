---
title: イベント ベースのOutlook用にアドインを構成する
description: イベント ベースのアクティブ化Outlookアドインを構成する方法について学習します。
ms.topic: article
ms.date: 06/08/2021
localization_priority: Normal
ms.openlocfilehash: 07790ee84693596f4873bc04d53c1e76c3825b4d
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076792"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation"></a><span data-ttu-id="2d82c-103">イベント ベースのOutlook用にアドインを構成する</span><span class="sxs-lookup"><span data-stu-id="2d82c-103">Configure your Outlook add-in for event-based activation</span></span>

<span data-ttu-id="2d82c-104">イベント ベースのアクティブ化機能がない場合、ユーザーはタスクを完了するためにアドインを明示的に起動する必要があります。</span><span class="sxs-lookup"><span data-stu-id="2d82c-104">Without the event-based activation feature, a user has to explicitly launch an add-in to complete their tasks.</span></span> <span data-ttu-id="2d82c-105">この機能を使用すると、特定のイベントに基づいて、特にすべてのアイテムに適用される操作に基づいてタスクを実行できます。</span><span class="sxs-lookup"><span data-stu-id="2d82c-105">This feature enables your add-in to run tasks based on certain events, particularly for operations that apply to every item.</span></span> <span data-ttu-id="2d82c-106">作業ウィンドウや UI レス機能と統合することもできます。</span><span class="sxs-lookup"><span data-stu-id="2d82c-106">You can also integrate with the task pane and UI-less functionality.</span></span>

<span data-ttu-id="2d82c-107">このチュートリアルの最後には、新しいアイテムが作成され、件名が設定されるたびに実行されるアドインがあります。</span><span class="sxs-lookup"><span data-stu-id="2d82c-107">By the end of this walkthrough, you'll have an add-in that runs whenever a new item is created and sets the subject.</span></span>

> [!NOTE]
> <span data-ttu-id="2d82c-108">この機能のサポートは、要件セット [1.10 で導入されました](../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md)。</span><span class="sxs-lookup"><span data-stu-id="2d82c-108">Support for this feature was introduced in [requirement set 1.10](../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md).</span></span> <span data-ttu-id="2d82c-109">この要件セットをサポートする [クライアントおよびプラットフォーム](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2d82c-109">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="supported-events"></a><span data-ttu-id="2d82c-110">サポートされるイベント</span><span class="sxs-lookup"><span data-stu-id="2d82c-110">Supported events</span></span>

<span data-ttu-id="2d82c-111">現時点では、次のイベントは Web および web サイトWindows。</span><span class="sxs-lookup"><span data-stu-id="2d82c-111">At present, the following events are supported on the web and on Windows.</span></span>

|<span data-ttu-id="2d82c-112">イベント</span><span class="sxs-lookup"><span data-stu-id="2d82c-112">Event</span></span>|<span data-ttu-id="2d82c-113">説明</span><span class="sxs-lookup"><span data-stu-id="2d82c-113">Description</span></span>|<span data-ttu-id="2d82c-114">最小値</span><span class="sxs-lookup"><span data-stu-id="2d82c-114">Minimum</span></span><br><span data-ttu-id="2d82c-115">要件セット</span><span class="sxs-lookup"><span data-stu-id="2d82c-115">requirement set</span></span>|
|---|---|---|
|`OnNewMessageCompose`|<span data-ttu-id="2d82c-116">新しいメッセージを作成する場合 (返信、すべて返信、転送を含む) が、下書きなど編集時には作成されません。</span><span class="sxs-lookup"><span data-stu-id="2d82c-116">On composing a new message (includes reply, reply all, and forward) but not on editing, for example, a draft.</span></span>|<span data-ttu-id="2d82c-117">1.10</span><span class="sxs-lookup"><span data-stu-id="2d82c-117">1.10</span></span>|
|`OnNewAppointmentOrganizer`|<span data-ttu-id="2d82c-118">既存の予定の編集ではなく、新しい予定を作成する場合。</span><span class="sxs-lookup"><span data-stu-id="2d82c-118">On creating a new appointment but not on editing an existing one.</span></span>|<span data-ttu-id="2d82c-119">1.10</span><span class="sxs-lookup"><span data-stu-id="2d82c-119">1.10</span></span>|
|`OnMessageAttachmentsChanged`|<span data-ttu-id="2d82c-120">メッセージの作成中に添付ファイルを追加または削除する場合。</span><span class="sxs-lookup"><span data-stu-id="2d82c-120">On adding or removing attachments while composing a message.</span></span>|<span data-ttu-id="2d82c-121">プレビュー</span><span class="sxs-lookup"><span data-stu-id="2d82c-121">Preview</span></span>|
|`OnAppointmentAttachmentsChanged`|<span data-ttu-id="2d82c-122">予定の作成中に添付ファイルを追加または削除する場合。</span><span class="sxs-lookup"><span data-stu-id="2d82c-122">On adding or removing attachments while composing an appointment.</span></span>|<span data-ttu-id="2d82c-123">プレビュー</span><span class="sxs-lookup"><span data-stu-id="2d82c-123">Preview</span></span>|
|`OnMessageRecipientsChanged`|<span data-ttu-id="2d82c-124">メッセージの作成中に受信者を追加または削除する場合。</span><span class="sxs-lookup"><span data-stu-id="2d82c-124">On adding or removing recipients while composing a message.</span></span>|<span data-ttu-id="2d82c-125">プレビュー</span><span class="sxs-lookup"><span data-stu-id="2d82c-125">Preview</span></span>|
|`OnAppointmentAttendeesChanged`|<span data-ttu-id="2d82c-126">予定の作成中に出席者を追加または削除する場合。</span><span class="sxs-lookup"><span data-stu-id="2d82c-126">On adding or removing attendees while composing an appointment.</span></span>|<span data-ttu-id="2d82c-127">プレビュー</span><span class="sxs-lookup"><span data-stu-id="2d82c-127">Preview</span></span>|
|`OnAppointmentTimeChanged`|<span data-ttu-id="2d82c-128">予定の作成中に日付/時刻を変更する場合。</span><span class="sxs-lookup"><span data-stu-id="2d82c-128">On changing date/time while composing an appointment.</span></span>|<span data-ttu-id="2d82c-129">プレビュー</span><span class="sxs-lookup"><span data-stu-id="2d82c-129">Preview</span></span>|
|`OnAppointmentRecurrenceChanged`|<span data-ttu-id="2d82c-130">予定の作成中に定期的な詳細を追加、変更、または削除する場合。</span><span class="sxs-lookup"><span data-stu-id="2d82c-130">On adding, changing, or removing the recurrence details while composing an appointment.</span></span> <span data-ttu-id="2d82c-131">日付/時刻が変更された場合、 `OnAppointmentTimeChanged` イベントも発生します。</span><span class="sxs-lookup"><span data-stu-id="2d82c-131">If the date/time is changed, the `OnAppointmentTimeChanged` event will also be fired.</span></span>|<span data-ttu-id="2d82c-132">プレビュー</span><span class="sxs-lookup"><span data-stu-id="2d82c-132">Preview</span></span>|
|`OnInfoBarDismissClicked`|<span data-ttu-id="2d82c-133">メッセージまたは予定アイテムの作成中に通知を却下する場合。</span><span class="sxs-lookup"><span data-stu-id="2d82c-133">On dismissing a notification while composing a message or appointment item.</span></span> <span data-ttu-id="2d82c-134">通知を追加したアドインだけが通知されます。</span><span class="sxs-lookup"><span data-stu-id="2d82c-134">Only the add-in that added the notification will be notified.</span></span>|<span data-ttu-id="2d82c-135">プレビュー</span><span class="sxs-lookup"><span data-stu-id="2d82c-135">Preview</span></span>|

> [!IMPORTANT]
> <span data-ttu-id="2d82c-136">プレビュー中のイベントは、Microsoft 365のサブスクリプションとOutlook on the webでのみWindows。</span><span class="sxs-lookup"><span data-stu-id="2d82c-136">Events still in preview are only available with a Microsoft 365 subscription in Outlook on the web and on Windows.</span></span> <span data-ttu-id="2d82c-137">詳細については、「この記事 [でプレビューする方法」](#how-to-preview) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2d82c-137">For more details, see [How to preview](#how-to-preview) in this article.</span></span> <span data-ttu-id="2d82c-138">プレビュー イベントは、実稼働アドインでは使用できません。</span><span class="sxs-lookup"><span data-stu-id="2d82c-138">Preview events shouldn't be used in production add-ins.</span></span>

### <a name="how-to-preview"></a><span data-ttu-id="2d82c-139">プレビューする方法</span><span class="sxs-lookup"><span data-stu-id="2d82c-139">How to preview</span></span>

<span data-ttu-id="2d82c-140">プレビューで今すぐイベントを試してみてください。</span><span class="sxs-lookup"><span data-stu-id="2d82c-140">We invite you to try out the events now in preview!</span></span> <span data-ttu-id="2d82c-141">このページの最後にある「フィードバック」セクションをGitHubフィードバックを提供することで、お客様のシナリオと改善方法をお知らせします。</span><span class="sxs-lookup"><span data-stu-id="2d82c-141">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="2d82c-142">これらのイベントをプレビューするには、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="2d82c-142">To preview these events:</span></span>

- <span data-ttu-id="2d82c-143">次のOutlook on the web。</span><span class="sxs-lookup"><span data-stu-id="2d82c-143">For Outlook on the web:</span></span>
  - <span data-ttu-id="2d82c-144">[ターゲット リリースをテナントにMicrosoft 365します](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)。</span><span class="sxs-lookup"><span data-stu-id="2d82c-144">[Configure targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span>
  - <span data-ttu-id="2d82c-145">()**の** ベータ ライブラリを参照 https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) CDN。</span><span class="sxs-lookup"><span data-stu-id="2d82c-145">Reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="2d82c-146">TypeScript[のコンパイル](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)と定義の種類定義ファイルは、IntelliSenseと[DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts)にあるCDNです。</span><span class="sxs-lookup"><span data-stu-id="2d82c-146">The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span></span> <span data-ttu-id="2d82c-147">これらの種類は、 を使用してインストールできます `npm install --save-dev @types/office-js-preview` 。</span><span class="sxs-lookup"><span data-stu-id="2d82c-147">You can install these types with `npm install --save-dev @types/office-js-preview`.</span></span>
- <span data-ttu-id="2d82c-148">[OutlookのWindows:</span><span class="sxs-lookup"><span data-stu-id="2d82c-148">For Outlook on Windows:</span></span>
  - <span data-ttu-id="2d82c-149">必要な最小ビルドは 16.0.14026.20000 です。</span><span class="sxs-lookup"><span data-stu-id="2d82c-149">The minimum required build is 16.0.14026.20000.</span></span> <span data-ttu-id="2d82c-150">ベータ版[ビルドOfficeアクセスするには、Insider](https://insider.office.com)プログラムOffice参加します。</span><span class="sxs-lookup"><span data-stu-id="2d82c-150">Join the [Office Insider program](https://insider.office.com) for access to Office beta builds.</span></span>
  - <span data-ttu-id="2d82c-151">レジストリを構成します。</span><span class="sxs-lookup"><span data-stu-id="2d82c-151">Configure the registry.</span></span> <span data-ttu-id="2d82c-152">Outlookから読み込む代わりに、製品版とベータ版Office.jsのローカル コピーが含CDN。</span><span class="sxs-lookup"><span data-stu-id="2d82c-152">Outlook includes a local copy of the production and beta versions of Office.js instead of loading from the CDN.</span></span> <span data-ttu-id="2d82c-153">既定では、API のローカル実稼働コピーが参照されます。</span><span class="sxs-lookup"><span data-stu-id="2d82c-153">By default, the local production copy of the API is referenced.</span></span> <span data-ttu-id="2d82c-154">JavaScript API のローカル ベータ コピーに切り替Outlook、このレジストリ エントリを追加する必要があります。それ以外の場合は、ベータ版 API が見つからない場合があります。</span><span class="sxs-lookup"><span data-stu-id="2d82c-154">To switch to the local beta copy of the Outlook JavaScript APIs, you need to add this registry entry, otherwise beta APIs may not be found.</span></span>
    1. <span data-ttu-id="2d82c-155">レジストリ キーを作成します `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\WebExt\Developer` 。</span><span class="sxs-lookup"><span data-stu-id="2d82c-155">Create the registry key `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\WebExt\Developer`.</span></span>
    1. <span data-ttu-id="2d82c-156">という名前のエントリを `EnableBetaAPIsInJavaScript` 追加し、値をに設定します `1` 。</span><span class="sxs-lookup"><span data-stu-id="2d82c-156">Add an entry named `EnableBetaAPIsInJavaScript` and set the value to `1`.</span></span> <span data-ttu-id="2d82c-157">レジストリは次の図のようになります。</span><span class="sxs-lookup"><span data-stu-id="2d82c-157">The following image shows what the registry should look like.</span></span>

        ![EnableBetaAPIsInJavaScript レジストリ キー値を持つレジストリ エディターのスクリーンショット。](../images/outlook-beta-registry-key.png)

## <a name="set-up-your-environment"></a><span data-ttu-id="2d82c-159">環境を設定する</span><span class="sxs-lookup"><span data-stu-id="2d82c-159">Set up your environment</span></span>

<span data-ttu-id="2d82c-160">クイック スタート[Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator)完了し、Yeoman ジェネレーターを使用してアドイン プロジェクトを作成し、Office作成します。</span><span class="sxs-lookup"><span data-stu-id="2d82c-160">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="2d82c-161">マニフェストを構成する</span><span class="sxs-lookup"><span data-stu-id="2d82c-161">Configure the manifest</span></span>

<span data-ttu-id="2d82c-162">アドインのイベント ベースのアクティブ化を有効にするには、マニフェストのノードで [Runtimes](../reference/manifest/runtimes.md) 要素と [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent) 拡張ポイント `VersionOverridesV1_1` を構成する必要があります。</span><span class="sxs-lookup"><span data-stu-id="2d82c-162">To enable event-based activation of your add-in, you must configure the [Runtimes](../reference/manifest/runtimes.md) element and [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent) extension point in the `VersionOverridesV1_1` node of the manifest.</span></span> <span data-ttu-id="2d82c-163">今のところ、 `DesktopFormFactor` サポートされている唯一のフォーム ファクターです。</span><span class="sxs-lookup"><span data-stu-id="2d82c-163">For now, `DesktopFormFactor` is the only supported form factor.</span></span>

1. <span data-ttu-id="2d82c-164">コード エディターで、クイック スタート プロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="2d82c-164">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="2d82c-165">プロジェクトの **manifest.xml** にあるファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="2d82c-165">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="2d82c-166">ノード全体 (開くタグと閉じるタグを含む) を選択し、次の XML に置き換え `<VersionOverrides>` 、変更を保存します。</span><span class="sxs-lookup"><span data-stu-id="2d82c-166">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML, then save your changes.</span></span>

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
              <LaunchEvent Type="OnMessageAttachmentsChanged" FunctionName="onMessageAttachmentsChangedHandler" />
              <LaunchEvent Type="OnAppointmentAttachmentsChanged" FunctionName="onAppointmentAttachmentsChangedHandler" />
              <LaunchEvent Type="OnMessageRecipientsChanged" FunctionName="onMessageRecipientsChangedHandler" />
              <LaunchEvent Type="OnAppointmentAttendeesChanged" FunctionName="onAppointmentAttendeesChangedHandler" />
              <LaunchEvent Type="OnAppointmentTimeChanged" FunctionName="onAppointmentTimeChangedHandler" />
              <LaunchEvent Type="OnAppointmentRecurrenceChanged" FunctionName="onAppointmentRecurrenceChangedHandler" />
              <LaunchEvent Type="OnInfoBarDismissClicked" FunctionName="onInfobarDismissClickedHandler" />
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

<span data-ttu-id="2d82c-167">OutlookはWindows JavaScript ファイルを使用しますが、Outlook on the web同じ JavaScript ファイルを参照できる HTML ファイルを使用します。</span><span class="sxs-lookup"><span data-stu-id="2d82c-167">Outlook on Windows uses a JavaScript file, while Outlook on the web uses an HTML file that can reference the same JavaScript file.</span></span> <span data-ttu-id="2d82c-168">Outlook プラットフォームは最終的に、Outlook クライアントに基づいて HTML または JavaScript を使用するかどうかを決定するために、マニフェストのノードでこれらの両方のファイル `Resources` への参照を提供する必要があります。</span><span class="sxs-lookup"><span data-stu-id="2d82c-168">You must provide references to both these files in the `Resources` node of the manifest as the Outlook platform ultimately determines whether to use HTML or JavaScript based on the Outlook client.</span></span> <span data-ttu-id="2d82c-169">そのため、イベント処理を構成するには、要素内の HTML の場所を指定し、その子要素で HTML によってインライン化または参照される JavaScript ファイルの場所 `Runtime` `Override` を指定します。</span><span class="sxs-lookup"><span data-stu-id="2d82c-169">As such, to configure event handling, provide the location of the HTML in the `Runtime` element, then in its `Override` child element provide the location of the JavaScript file inlined or referenced by the HTML.</span></span>

> [!TIP]
> <span data-ttu-id="2d82c-170">アドインのマニフェストのOutlook詳細については、「Outlook[マニフェスト」を参照してください](manifests.md)。</span><span class="sxs-lookup"><span data-stu-id="2d82c-170">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-event-handling"></a><span data-ttu-id="2d82c-171">イベント処理の実装</span><span class="sxs-lookup"><span data-stu-id="2d82c-171">Implement event handling</span></span>

<span data-ttu-id="2d82c-172">選択したイベントの処理を実装する必要があります。</span><span class="sxs-lookup"><span data-stu-id="2d82c-172">You have to implement handling for your selected events.</span></span>

<span data-ttu-id="2d82c-173">このシナリオでは、新しいアイテムを作成する処理を追加します。</span><span class="sxs-lookup"><span data-stu-id="2d82c-173">In this scenario, you'll add handling for composing new items.</span></span>

1. <span data-ttu-id="2d82c-174">同じクイック スタート プロジェクトで、コード エディター **で ./src/commands/commands.js** ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="2d82c-174">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="2d82c-175">関数の `action` 後に、次の JavaScript 関数を挿入します。</span><span class="sxs-lookup"><span data-stu-id="2d82c-175">After the `action` function, insert the following JavaScript functions.</span></span>

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

1. <span data-ttu-id="2d82c-176">ファイルの末尾に次の JavaScript コードを追加します。</span><span class="sxs-lookup"><span data-stu-id="2d82c-176">Add the following JavaScript code at the end of the file.</span></span>

    ```js
    // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
    Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
    Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
    ```

1. <span data-ttu-id="2d82c-177">変更内容を保存します。</span><span class="sxs-lookup"><span data-stu-id="2d82c-177">Save your changes.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="2d82c-178">Windows: 現在、イベント ベースのアクティブ化の処理を実装する JavaScript ファイルではインポートはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="2d82c-178">Windows: At present, imports are not supported in the JavaScript file where you implement the handling for event-based activation.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="2d82c-179">試してみる</span><span class="sxs-lookup"><span data-stu-id="2d82c-179">Try it out</span></span>

1. <span data-ttu-id="2d82c-180">プロジェクトのルート ディレクトリから次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="2d82c-180">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="2d82c-181">このコマンドを実行すると、ローカル Web サーバーが (まだ実行されていない場合) 起動し、アドインがサイドロードされます。</span><span class="sxs-lookup"><span data-stu-id="2d82c-181">When you run this command, the local web server will start (if it's not already running) and your add-in will be sideloaded.</span></span>

    ```command&nbsp;line
    npm start
    ```

    > [!NOTE]
    > <span data-ttu-id="2d82c-182">アドインが自動的にサイドロードされていない場合は、サイドロード[Outlook](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually)アドインの手順に従ってテストを行い、Outlook でアドインを手動でサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="2d82c-182">If your add-in wasn't automatically sideloaded, then follow the instructions in [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually) to manually sideload the add-in in Outlook.</span></span>

1. <span data-ttu-id="2d82c-183">Outlook on the web で新しいメッセージを作成します。</span><span class="sxs-lookup"><span data-stu-id="2d82c-183">In Outlook on the web, create a new message.</span></span>

    ![作成時に件名が設定されているOutlook on the webウィンドウのスクリーンショット。](../images/outlook-web-autolaunch-1.png)

1. <span data-ttu-id="2d82c-185">[Outlook] でWindows新しいメッセージを作成します。</span><span class="sxs-lookup"><span data-stu-id="2d82c-185">In Outlook on Windows, create a new message.</span></span>

    ![作成時に件名が設定されているOutlookのWindowsウィンドウのスクリーンショット。](../images/outlook-win-autolaunch.png)

    > [!NOTE]
    > <span data-ttu-id="2d82c-187">localhost からアドインを実行している場合は、"申し訳ありませんが *、{your-add-in-name-here}* にアクセスできませんでした。</span><span class="sxs-lookup"><span data-stu-id="2d82c-187">If you're running your add-in from localhost and see the error "We're sorry, we couldn't access *{your-add-in-name-here}*.</span></span> <span data-ttu-id="2d82c-188">ネットワーク接続が確立されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="2d82c-188">Make sure you have a network connection.</span></span> <span data-ttu-id="2d82c-189">問題が解決しない場合は、後でもう一度お試しください。ループバックの除外を有効にする必要がある場合があります。</span><span class="sxs-lookup"><span data-stu-id="2d82c-189">If the problem continues, please try again later.", you may need to enable a loopback exemption.</span></span>
    >
    > 1. <span data-ttu-id="2d82c-190">Outlook を終了します。</span><span class="sxs-lookup"><span data-stu-id="2d82c-190">Close Outlook.</span></span>
    > 1. <span data-ttu-id="2d82c-191">タスク マネージャー **を開** き、タスク **msoadfsb.exeが** 実行されていないか確認します。</span><span class="sxs-lookup"><span data-stu-id="2d82c-191">Open the **Task Manager** and ensure that the **msoadfsb.exe** process is not running.</span></span>
    > 1. <span data-ttu-id="2d82c-192">次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="2d82c-192">Run the following command.</span></span>
    >
    >    ```command&nbsp;line
    >    call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_http___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
    >    ```
    >
    > 1. <span data-ttu-id="2d82c-193">Outlook を再起動します。</span><span class="sxs-lookup"><span data-stu-id="2d82c-193">Restart Outlook.</span></span>

## <a name="debug"></a><span data-ttu-id="2d82c-194">Debug</span><span class="sxs-lookup"><span data-stu-id="2d82c-194">Debug</span></span>

<span data-ttu-id="2d82c-195">アドインで起動イベント処理に変更を加える場合は、次の点に注意する必要があります。</span><span class="sxs-lookup"><span data-stu-id="2d82c-195">As you make changes to launch-event handling in your add-in, you should be aware that:</span></span>

- <span data-ttu-id="2d82c-196">マニフェストを更新した場合は、 [アドインを](sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in) 削除してから、もう一度サイドロードします。</span><span class="sxs-lookup"><span data-stu-id="2d82c-196">If you updated the manifest, [remove the add-in](sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in) then sideload it again.</span></span>
- <span data-ttu-id="2d82c-197">マニフェスト以外のファイルに変更を加えた場合は、OutlookでWindowsを閉じて再び開Outlook on the web。</span><span class="sxs-lookup"><span data-stu-id="2d82c-197">If you made changes to files other than the manifest, close and reopen Outlook on Windows, or refresh the browser tab running Outlook on the web.</span></span>

<span data-ttu-id="2d82c-198">独自の機能を実装する場合は、コードのデバッグが必要な場合があります。</span><span class="sxs-lookup"><span data-stu-id="2d82c-198">While implementing your own functionality, you may need to debug your code.</span></span> <span data-ttu-id="2d82c-199">イベント ベースのアドインのアクティブ化をデバッグする方法のガイダンスについては、「Debug [your event-based Outlook アドイン」を参照してください](debug-autolaunch.md)。</span><span class="sxs-lookup"><span data-stu-id="2d82c-199">For guidance on how to debug event-based add-in activation, see [Debug your event-based Outlook add-in](debug-autolaunch.md).</span></span>

<span data-ttu-id="2d82c-200">ランタイム ログは、この機能に対して、Windows。</span><span class="sxs-lookup"><span data-stu-id="2d82c-200">Runtime logging is also available for this feature on Windows.</span></span> <span data-ttu-id="2d82c-201">詳細については、「ランタイム ログを [使用してアドインをデバッグする」を参照してください](../testing/runtime-logging.md#runtime-logging-on-windows)。</span><span class="sxs-lookup"><span data-stu-id="2d82c-201">For more information, see [Debug your add-in with runtime logging](../testing/runtime-logging.md#runtime-logging-on-windows).</span></span>

## <a name="deploy-to-users"></a><span data-ttu-id="2d82c-202">ユーザーへの展開</span><span class="sxs-lookup"><span data-stu-id="2d82c-202">Deploy to users</span></span>

<span data-ttu-id="2d82c-203">イベント ベースのアドインを展開するには、イベント ベースのアドインを使用してマニフェストをアップロードMicrosoft 365 管理センター。</span><span class="sxs-lookup"><span data-stu-id="2d82c-203">You can deploy event-based add-ins by uploading the manifest through the Microsoft 365 admin center.</span></span> <span data-ttu-id="2d82c-204">管理ポータルで、ナビゲーション ウィンドウで [設定] セクションを展開し、[統合アプリ]**を選択します**。</span><span class="sxs-lookup"><span data-stu-id="2d82c-204">In the admin portal, expand the **Settings** section in the navigation pane then select **Integrated apps**.</span></span> <span data-ttu-id="2d82c-205">[統合アプリ **] ページで**、[カスタム アプリ] アップロード **を選択** します。</span><span class="sxs-lookup"><span data-stu-id="2d82c-205">On the **Integrated apps** page, choose the **Upload custom apps** action.</span></span>

![カスタム アプリのアクションを含む、Microsoft 365 管理センター統合アプリ ページアップロードスクリーンショット。](../images/outlook-deploy-event-based-add-ins.png)

<span data-ttu-id="2d82c-207">AppSource ストアと inclient ストア: イベント ベースのアドインを展開したり、既存のアドインを更新してイベント ベースのアクティブ化機能を含める機能をすぐに利用できる必要があります。</span><span class="sxs-lookup"><span data-stu-id="2d82c-207">AppSource and inclient stores: The ability to deploy event-based add-ins or update existing add-ins to include the event-based activation feature should be available soon.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="2d82c-208">イベント ベースのアドインは、管理者が管理する展開にのみ制限されます。</span><span class="sxs-lookup"><span data-stu-id="2d82c-208">Event-based add-ins are restricted to admin-managed deployments only.</span></span> <span data-ttu-id="2d82c-209">今のところ、ユーザーは AppSource ストアまたは inclient ストアからイベント ベースのアドインを取得できます。</span><span class="sxs-lookup"><span data-stu-id="2d82c-209">For now, users can't get event-based add-ins from AppSource or inclient stores.</span></span>

## <a name="event-based-activation-behavior-and-limitations"></a><span data-ttu-id="2d82c-210">イベント ベースのアクティブ化の動作と制限</span><span class="sxs-lookup"><span data-stu-id="2d82c-210">Event-based activation behavior and limitations</span></span>

<span data-ttu-id="2d82c-211">アドイン起動イベント ハンドラーは、実行時間が短く、軽量で、可能な限り非インバシブである必要があります。</span><span class="sxs-lookup"><span data-stu-id="2d82c-211">Add-in launch-event handlers are expected to be short-running, lightweight, and as noninvasive as possible.</span></span> <span data-ttu-id="2d82c-212">アクティブ化後、アドインはイベント ベースのアドインを実行できる最大時間である約 300 秒以内にタイム アウトします。アドインが起動イベントの処理を完了したというメッセージを表示するには、関連付けられたハンドラーにメソッドを呼び出す必要 `event.completed` があります。</span><span class="sxs-lookup"><span data-stu-id="2d82c-212">After activation, your add-in will time out within approximately 300 seconds, the maximum length of time allowed for running event-based add-ins. To signal that your add-in has completed processing a launch event, we recommend you have the associated handler call the `event.completed` method.</span></span> <span data-ttu-id="2d82c-213">(ステートメントの後に含まれる `event.completed` コードは、実行が保証されない点に注意してください)。アドインが処理するイベントがトリガーされるごとに、アドインが再アクティブ化され、関連付けられたイベント ハンドラーが実行され、タイムアウト ウィンドウがリセットされます。</span><span class="sxs-lookup"><span data-stu-id="2d82c-213">(Note that code included after the `event.completed` statement is not guaranteed to run.) Each time an event that your add-in handles is triggered, the add-in is reactivated and runs the associated event handler, and the timeout window is reset.</span></span> <span data-ttu-id="2d82c-214">アドインは、タイム アウト後に終了するか、ユーザーが作成ウィンドウを閉じるか、アイテムを送信します。</span><span class="sxs-lookup"><span data-stu-id="2d82c-214">The add-in ends after it times out, or the user closes the compose window or sends the item.</span></span>

<span data-ttu-id="2d82c-215">ユーザーが同じイベントにサブスクライブしている複数のアドインがある場合、Outlook プラットフォームは特定の順序でアドインを起動します。</span><span class="sxs-lookup"><span data-stu-id="2d82c-215">If the user has multiple add-ins that subscribed to the same event, the Outlook platform launches the add-ins in no particular order.</span></span> <span data-ttu-id="2d82c-216">現在、アクティブに実行できるイベント ベースのアドインは 5 つのみです。</span><span class="sxs-lookup"><span data-stu-id="2d82c-216">Currently, only five event-based add-ins can be actively running.</span></span>

<span data-ttu-id="2d82c-217">ユーザーは、アドインの実行を開始した現在のメール アイテムから切り替えまたは移動できます。</span><span class="sxs-lookup"><span data-stu-id="2d82c-217">The user can switch or navigate away from the current mail item where the add-in started running.</span></span> <span data-ttu-id="2d82c-218">起動されたアドインは、バックグラウンドで操作を終了します。</span><span class="sxs-lookup"><span data-stu-id="2d82c-218">The add-in that was launched will finish its operation in the background.</span></span>

<span data-ttu-id="2d82c-219">JavaScript ファイルでは、イベント ベースのアクティブ化の処理をクライアントで実装する場合、インポートはWindowsされません。</span><span class="sxs-lookup"><span data-stu-id="2d82c-219">Imports are not supported in the JavaScript file where you implement the handling for event-based activation in the Windows client.</span></span>

<span data-ttu-id="2d82c-220">UI Office.js変更する API の一部は、イベント ベースのアドインでは使用できない場合があります。ブロックされている API を次に示します。</span><span class="sxs-lookup"><span data-stu-id="2d82c-220">Some Office.js APIs that change or alter the UI are not allowed from event-based add-ins. The following are the blocked APIs:</span></span>

- <span data-ttu-id="2d82c-221">[ `OfficeRuntime.auth` : ] の下</span><span class="sxs-lookup"><span data-stu-id="2d82c-221">Under `OfficeRuntime.auth`:</span></span>
  - <span data-ttu-id="2d82c-222">`getAccessToken`(Windowsのみ)</span><span class="sxs-lookup"><span data-stu-id="2d82c-222">`getAccessToken` (Windows only)</span></span>
- <span data-ttu-id="2d82c-223">[ `Office.context.auth` : ] の下</span><span class="sxs-lookup"><span data-stu-id="2d82c-223">Under `Office.context.auth`:</span></span>
  - `getAccessToken`
  - `getAccessTokenAsync`
- <span data-ttu-id="2d82c-224">[ `Office.context.mailbox` : ] の下</span><span class="sxs-lookup"><span data-stu-id="2d82c-224">Under `Office.context.mailbox`:</span></span>
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- <span data-ttu-id="2d82c-225">[ `Office.context.mailbox.item` : ] の下</span><span class="sxs-lookup"><span data-stu-id="2d82c-225">Under `Office.context.mailbox.item`:</span></span>
  - `close`
- <span data-ttu-id="2d82c-226">[ `Office.context.ui` : ] の下</span><span class="sxs-lookup"><span data-stu-id="2d82c-226">Under `Office.context.ui`:</span></span>
  - `displayDialogAsync`
  - `messageParent`

## <a name="see-also"></a><span data-ttu-id="2d82c-227">関連項目</span><span class="sxs-lookup"><span data-stu-id="2d82c-227">See also</span></span>

- [<span data-ttu-id="2d82c-228">Outlook アドインのマニフェスト</span><span class="sxs-lookup"><span data-stu-id="2d82c-228">Outlook add-in manifests</span></span>](manifests.md)
- [<span data-ttu-id="2d82c-229">イベント ベースのアドインをデバッグする方法</span><span class="sxs-lookup"><span data-stu-id="2d82c-229">How to debug event-based add-ins</span></span>](debug-autolaunch.md)
- <span data-ttu-id="2d82c-230">PnP サンプル:[イベント Outlookアクティブ化を使用して署名を設定する](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/outlook-set-signature)</span><span class="sxs-lookup"><span data-stu-id="2d82c-230">PnP sample: [Use Outlook event-based activation to set the signature](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/outlook-set-signature)</span></span>