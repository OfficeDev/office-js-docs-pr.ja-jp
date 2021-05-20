---
title: イベント ベースのアクティブ化 (プレビュー) 用にOutlook アドインを構成する
description: イベント ベースのアクティブ化用にOutlook アドインを構成する方法について説明します。
ms.topic: article
ms.date: 05/18/2021
localization_priority: Normal
ms.openlocfilehash: 721f05e1c835e066744598ecb2bd416c6a6b0526
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555241"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a><span data-ttu-id="ef105-103">イベント ベースのアクティブ化 (プレビュー) 用にOutlook アドインを構成する</span><span class="sxs-lookup"><span data-stu-id="ef105-103">Configure your Outlook add-in for event-based activation (preview)</span></span>

<span data-ttu-id="ef105-104">イベントベースのアクティブ化機能を使用しない場合、ユーザーは、アドインを明示的に起動してタスクを完了する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ef105-104">Without the event-based activation feature, a user has to explicitly launch an add-in to complete their tasks.</span></span> <span data-ttu-id="ef105-105">この機能により、特定のイベントに基づいて、特にすべての項目に適用される操作に基づいて、アドインでタスクを実行できます。</span><span class="sxs-lookup"><span data-stu-id="ef105-105">This feature enables your add-in to run tasks based on certain events, particularly for operations that apply to every item.</span></span> <span data-ttu-id="ef105-106">作業ウィンドウと UI を使用する機能を統合することもできます。</span><span class="sxs-lookup"><span data-stu-id="ef105-106">You can also integrate with the task pane and UI-less functionality.</span></span>

<span data-ttu-id="ef105-107">このチュートリアルの最後に、新しい項目が作成され、件名を設定するたびに実行されるアドインが用意されています。</span><span class="sxs-lookup"><span data-stu-id="ef105-107">By the end of this walkthrough, you'll have an add-in that runs whenever a new item is created and sets the subject.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ef105-108">この機能は、web 上のOutlookで[プレビュー](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)し、Microsoft 365サブスクリプションでWindowsにのみサポートされます。</span><span class="sxs-lookup"><span data-stu-id="ef105-108">This feature is only supported for [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web and on Windows with a Microsoft 365 subscription.</span></span> <span data-ttu-id="ef105-109">詳細については、この記事 [の「イベントベースのアクティブ化機能をプレビューする方法](#how-to-preview-the-event-based-activation-feature) 」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ef105-109">For more details, see [How to preview the event-based activation feature](#how-to-preview-the-event-based-activation-feature) in this article.</span></span>
>
> <span data-ttu-id="ef105-110">プレビュー機能は予告なく変更される場合があるため、運用アドインで使用しないでください。</span><span class="sxs-lookup"><span data-stu-id="ef105-110">Because preview features are subject to change without notice, they shouldn't be used in production add-ins.</span></span>

## <a name="supported-events"></a><span data-ttu-id="ef105-111">サポートされるイベント</span><span class="sxs-lookup"><span data-stu-id="ef105-111">Supported events</span></span>

<span data-ttu-id="ef105-112">現在、以下のイベントがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="ef105-112">At present, the following events are supported.</span></span>

|<span data-ttu-id="ef105-113">イベント</span><span class="sxs-lookup"><span data-stu-id="ef105-113">Event</span></span>|<span data-ttu-id="ef105-114">説明</span><span class="sxs-lookup"><span data-stu-id="ef105-114">Description</span></span>|<span data-ttu-id="ef105-115">クライアント</span><span class="sxs-lookup"><span data-stu-id="ef105-115">Clients</span></span>|
|---|---|---|
|`OnNewMessageCompose`|<span data-ttu-id="ef105-116">新しいメッセージ (返信、全員への返信、転送を含む) の作成時に、編集時 (下書きなど) は作成しません。</span><span class="sxs-lookup"><span data-stu-id="ef105-116">On composing a new message (includes reply, reply all, and forward) but not on editing, for example, a draft.</span></span>|<span data-ttu-id="ef105-117">Windows, ウェブ</span><span class="sxs-lookup"><span data-stu-id="ef105-117">Windows, web</span></span>|
|`OnNewAppointmentOrganizer`|<span data-ttu-id="ef105-118">新しい予定を作成するが、既存の予定を編集する場合。</span><span class="sxs-lookup"><span data-stu-id="ef105-118">On creating a new appointment but not on editing an existing one.</span></span>|<span data-ttu-id="ef105-119">Windows, ウェブ</span><span class="sxs-lookup"><span data-stu-id="ef105-119">Windows, web</span></span>|
|`OnMessageAttachmentsChanged`|<span data-ttu-id="ef105-120">メッセージの作成中に添付ファイルを追加または削除する。</span><span class="sxs-lookup"><span data-stu-id="ef105-120">On adding or removing attachments while composing a message.</span></span>|<span data-ttu-id="ef105-121">Windows</span><span class="sxs-lookup"><span data-stu-id="ef105-121">Windows</span></span>|
|`OnAppointmentAttachmentsChanged`|<span data-ttu-id="ef105-122">予定の作成中に添付ファイルを追加または削除する。</span><span class="sxs-lookup"><span data-stu-id="ef105-122">On adding or removing attachments while composing an appointment.</span></span>|<span data-ttu-id="ef105-123">Windows</span><span class="sxs-lookup"><span data-stu-id="ef105-123">Windows</span></span>|
|`OnMessageRecipientsChanged`|<span data-ttu-id="ef105-124">メッセージの作成中に受信者を追加または削除する。</span><span class="sxs-lookup"><span data-stu-id="ef105-124">On adding or removing recipients while composing a message.</span></span>|<span data-ttu-id="ef105-125">Windows</span><span class="sxs-lookup"><span data-stu-id="ef105-125">Windows</span></span>|
|`OnAppointmentAttendeesChanged`|<span data-ttu-id="ef105-126">予定の作成中に出席者を追加または削除する。</span><span class="sxs-lookup"><span data-stu-id="ef105-126">On adding or removing attendees while composing an appointment.</span></span>|<span data-ttu-id="ef105-127">Windows</span><span class="sxs-lookup"><span data-stu-id="ef105-127">Windows</span></span>|
|`OnAppointmentTimeChanged`|<span data-ttu-id="ef105-128">予定の作成中に日付/時刻を変更する場合。</span><span class="sxs-lookup"><span data-stu-id="ef105-128">On changing date/time while composing an appointment.</span></span>|<span data-ttu-id="ef105-129">Windows</span><span class="sxs-lookup"><span data-stu-id="ef105-129">Windows</span></span>|
|`OnAppointmentRecurrenceChanged`|<span data-ttu-id="ef105-130">予定の作成中に定期的なアイテムの詳細を追加、変更、または削除する。</span><span class="sxs-lookup"><span data-stu-id="ef105-130">On adding, changing, or removing the recurrence details while composing an appointment.</span></span> <span data-ttu-id="ef105-131">日付/時刻が変更されると、 `OnAppointmentTimeChanged` イベントも発生します。</span><span class="sxs-lookup"><span data-stu-id="ef105-131">If the date/time is changed, the `OnAppointmentTimeChanged` event will also be fired.</span></span>|<span data-ttu-id="ef105-132">Windows</span><span class="sxs-lookup"><span data-stu-id="ef105-132">Windows</span></span>|
|`OnInfoBarDismissClicked`|<span data-ttu-id="ef105-133">メッセージまたは予定アイテムの作成中に通知を閉じる。</span><span class="sxs-lookup"><span data-stu-id="ef105-133">On dismissing a notification while composing a message or appointment item.</span></span> <span data-ttu-id="ef105-134">通知を追加したアドインのみが通知されます。</span><span class="sxs-lookup"><span data-stu-id="ef105-134">Only the add-in that added the notification will be notified.</span></span>|<span data-ttu-id="ef105-135">Windows</span><span class="sxs-lookup"><span data-stu-id="ef105-135">Windows</span></span>|

## <a name="how-to-preview-the-event-based-activation-feature"></a><span data-ttu-id="ef105-136">イベントベースのアクティブ化機能をプレビューする方法</span><span class="sxs-lookup"><span data-stu-id="ef105-136">How to preview the event-based activation feature</span></span>

<span data-ttu-id="ef105-137">イベントベースのアクティベーション機能を試してみてください!</span><span class="sxs-lookup"><span data-stu-id="ef105-137">We invite you to try out the event-based activation feature!</span></span> <span data-ttu-id="ef105-138">GitHubを通じてフィードバックを提供することで、お客様のシナリオと改善方法をお知らせください(このページの最後にある **フィードバック** セクションを参照)。</span><span class="sxs-lookup"><span data-stu-id="ef105-138">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="ef105-139">この機能をプレビューするには:</span><span class="sxs-lookup"><span data-stu-id="ef105-139">To preview this feature:</span></span>

- <span data-ttu-id="ef105-140">ウェブ上のOutlookの場合:</span><span class="sxs-lookup"><span data-stu-id="ef105-140">For Outlook on the web:</span></span>
  - <span data-ttu-id="ef105-141">[Microsoft 365 テナントで対象リリースを構成する](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center):</span><span class="sxs-lookup"><span data-stu-id="ef105-141">[Configure targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span>
  - <span data-ttu-id="ef105-142">CDN ( ) の **ベータ ライブラリ** を参照 https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) します。</span><span class="sxs-lookup"><span data-stu-id="ef105-142">Reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="ef105-143">TypeScript コンパイルとIntelliSenseの[型定義ファイル](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)は、CDNと[Typed にあります](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts)。</span><span class="sxs-lookup"><span data-stu-id="ef105-143">The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span></span> <span data-ttu-id="ef105-144">これらの種類は、 を使用してインストールできます `npm install --save-dev @types/office-js-preview` 。</span><span class="sxs-lookup"><span data-stu-id="ef105-144">You can install these types with `npm install --save-dev @types/office-js-preview`.</span></span>
- <span data-ttu-id="ef105-145">WindowsのOutlookの場合:</span><span class="sxs-lookup"><span data-stu-id="ef105-145">For Outlook on Windows:</span></span>
  - <span data-ttu-id="ef105-146">必要最小限のビルドは 16.0.14026.20000 です。</span><span class="sxs-lookup"><span data-stu-id="ef105-146">The minimum required build is 16.0.14026.20000.</span></span> <span data-ttu-id="ef105-147">ベータ版ビルドにアクセスするには[、Office Insider プログラム](https://insider.office.com)Office参加してください。</span><span class="sxs-lookup"><span data-stu-id="ef105-147">Join the [Office Insider program](https://insider.office.com) for access to Office beta builds.</span></span>
  - <span data-ttu-id="ef105-148">レジストリを構成します。</span><span class="sxs-lookup"><span data-stu-id="ef105-148">Configure the registry.</span></span> <span data-ttu-id="ef105-149">Outlookには、CDNから読み込む代わりに、Office.jsの実稼働およびベータ版のローカル コピーが含まれます。</span><span class="sxs-lookup"><span data-stu-id="ef105-149">Outlook includes a local copy of the production and beta versions of Office.js instead of loading from the CDN.</span></span> <span data-ttu-id="ef105-150">デフォルトでは、API のローカル本番コピーが参照されます。</span><span class="sxs-lookup"><span data-stu-id="ef105-150">By default, the local production copy of the API is referenced.</span></span> <span data-ttu-id="ef105-151">Outlookの JavaScript API のローカル ベータ 版に切り替えるには、このレジストリ エントリを追加する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ef105-151">To switch to the local beta copy of the Outlook JavaScript APIs, you need to add this registry entry, otherwise beta APIs may not be found.</span></span>
    1. <span data-ttu-id="ef105-152">レジストリ キーを作成 `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\WebExt\Developer` する:</span><span class="sxs-lookup"><span data-stu-id="ef105-152">Create the registry key `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\WebExt\Developer`.</span></span>
    1. <span data-ttu-id="ef105-153">という名前のエントリを追加 `EnableBetaAPIsInJavaScript` し、値を `1` に設定します。</span><span class="sxs-lookup"><span data-stu-id="ef105-153">Add an entry named `EnableBetaAPIsInJavaScript` and set the value to `1`.</span></span> <span data-ttu-id="ef105-154">レジストリは次の図のようになります。</span><span class="sxs-lookup"><span data-stu-id="ef105-154">The following image shows what the registry should look like.</span></span>

        ![レジストリ キー値を持つレジストリ エディターのスクリーンショット](../images/outlook-beta-registry-key.png)

## <a name="set-up-your-environment"></a><span data-ttu-id="ef105-156">環境を設定する</span><span class="sxs-lookup"><span data-stu-id="ef105-156">Set up your environment</span></span>

<span data-ttu-id="ef105-157">アドイン[Outlookクイック スタート](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator)を完了し、アドインの Office用の Yeoman ジェネレーターを使用してアドイン プロジェクトを作成します。</span><span class="sxs-lookup"><span data-stu-id="ef105-157">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="ef105-158">マニフェストを構成する</span><span class="sxs-lookup"><span data-stu-id="ef105-158">Configure the manifest</span></span>

<span data-ttu-id="ef105-159">アドインのイベント ベースのアクティブ化を有効にするには、マニフェストのノードで [Runtimes](../reference/manifest/runtimes.md) 要素と [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) 拡張ポイントを構成する必要があります `VersionOverridesV1_1` 。</span><span class="sxs-lookup"><span data-stu-id="ef105-159">To enable event-based activation of your add-in, you must configure the [Runtimes](../reference/manifest/runtimes.md) element and [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) extension point in the `VersionOverridesV1_1` node of the manifest.</span></span> <span data-ttu-id="ef105-160">現時点では、 `DesktopFormFactor` サポートされているフォーム ファクターは唯一です。</span><span class="sxs-lookup"><span data-stu-id="ef105-160">For now, `DesktopFormFactor` is the only supported form factor.</span></span>

1. <span data-ttu-id="ef105-161">コード エディターで、クイック スタート プロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="ef105-161">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="ef105-162">プロジェクトのルートにある **manifest.xml** ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="ef105-162">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="ef105-163">ノード全体 `<VersionOverrides>` (開くタグと閉じるタグを含む) を選択し、次の XML で置き換えて、変更を保存します。</span><span class="sxs-lookup"><span data-stu-id="ef105-163">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML, then save your changes.</span></span>

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
              <!-- Events supported on the web and on Windows. -->
              <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
              <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onAppointmentComposeHandler"/>
              <!-- Events supported only on Windows. -->
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

<span data-ttu-id="ef105-164">Windows上のOutlookは JavaScript ファイルを使用しますが、web 上のOutlookは同じ JavaScript ファイルを参照できる HTML ファイルを使用します。</span><span class="sxs-lookup"><span data-stu-id="ef105-164">Outlook on Windows uses a JavaScript file, while Outlook on the web uses an HTML file that can reference the same JavaScript file.</span></span> <span data-ttu-id="ef105-165">Outlook プラットフォームが最終的に `Resources` Outlook クライアントに基づいて HTML または JavaScript を使用するかどうかを決定するので、マニフェストのノードでこれらのファイルの両方への参照を指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ef105-165">You must provide references to both these files in the `Resources` node of the manifest as the Outlook platform ultimately determines whether to use HTML or JavaScript based on the Outlook client.</span></span> <span data-ttu-id="ef105-166">そのため、イベント処理を構成するには、要素内の HTML の場所を指定 `Runtime` し、その `Override` 子要素で、インライン化または HTML によって参照される JavaScript ファイルの場所を提供します。</span><span class="sxs-lookup"><span data-stu-id="ef105-166">As such, to configure event handling, provide the location of the HTML in the `Runtime` element, then in its `Override` child element provide the location of the JavaScript file inlined or referenced by the HTML.</span></span>

> [!TIP]
> <span data-ttu-id="ef105-167">Outlook アドインのマニフェストの詳細については、「アドイン マニフェスト[のOutlook」](manifests.md)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ef105-167">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-event-handling"></a><span data-ttu-id="ef105-168">イベント処理の実装</span><span class="sxs-lookup"><span data-stu-id="ef105-168">Implement event handling</span></span>

<span data-ttu-id="ef105-169">選択したイベントの処理を実装する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ef105-169">You have to implement handling for your selected events.</span></span>

<span data-ttu-id="ef105-170">このシナリオでは、新しいアイテムを作成するための処理を追加します。</span><span class="sxs-lookup"><span data-stu-id="ef105-170">In this scenario, you'll add handling for composing new items.</span></span>

1. <span data-ttu-id="ef105-171">同じクイック スタート プロジェクトから、コード エディターでファイル **./src/commands/commands.js** を開きます。</span><span class="sxs-lookup"><span data-stu-id="ef105-171">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="ef105-172">関数の後 `action` に、次の JavaScript 関数を挿入します。</span><span class="sxs-lookup"><span data-stu-id="ef105-172">After the `action` function, insert the following JavaScript functions.</span></span>

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

1. <span data-ttu-id="ef105-173">ファイルの末尾に次の JavaScript コードを追加します。</span><span class="sxs-lookup"><span data-stu-id="ef105-173">Add the following JavaScript code at the end of the file.</span></span>

    ```js
    // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
    Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
    Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
    ```

1. <span data-ttu-id="ef105-174">変更内容を保存します。</span><span class="sxs-lookup"><span data-stu-id="ef105-174">Save your changes.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="ef105-175">試してみる</span><span class="sxs-lookup"><span data-stu-id="ef105-175">Try it out</span></span>

1. <span data-ttu-id="ef105-176">プロジェクトのルート ディレクトリから次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="ef105-176">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="ef105-177">このコマンドを実行すると、ローカル Web サーバーが (まだ実行されていない場合) 起動し、アドインがサイドロードされます。</span><span class="sxs-lookup"><span data-stu-id="ef105-177">When you run this command, the local web server will start (if it's not already running) and your add-in will be sideloaded.</span></span>

    ```command&nbsp;line
    npm start
    ```

    > [!NOTE]
    > <span data-ttu-id="ef105-178">アドインが自動的にサイドロードされなかった場合は、「[サイドロード Outlook アドインをテスト用に](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually)実行する」の指示に従って、アドインを手動でOutlookサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="ef105-178">If your add-in wasn't automatically sideloaded, then follow the instructions in [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually) to manually sideload the add-in in Outlook.</span></span>

1. <span data-ttu-id="ef105-179">Outlook on the web で新しいメッセージを作成します。</span><span class="sxs-lookup"><span data-stu-id="ef105-179">In Outlook on the web, create a new message.</span></span>

    ![Web 上のOutlookのメッセージ ウィンドウのスクリーンショット (作成時に設定された件名)](../images/outlook-web-autolaunch-1.png)

1. <span data-ttu-id="ef105-181">Windows Outlookで、新しいメッセージを作成します。</span><span class="sxs-lookup"><span data-stu-id="ef105-181">In Outlook on Windows, create a new message.</span></span>

    ![作成時に設定された件名を持つWindowsのOutlookのメッセージ ウィンドウのスクリーンショット](../images/outlook-win-autolaunch.png)

    > [!NOTE]
    > <span data-ttu-id="ef105-183">localhost からアドインを実行していて、"申し訳ありませんが *、{アドイン名- here}* にアクセスできませんでした。</span><span class="sxs-lookup"><span data-stu-id="ef105-183">If you're running your add-in from localhost and see the error "We're sorry, we couldn't access *{your-add-in-name-here}*.</span></span> <span data-ttu-id="ef105-184">ネットワーク接続があることを確認します。</span><span class="sxs-lookup"><span data-stu-id="ef105-184">Make sure you have a network connection.</span></span> <span data-ttu-id="ef105-185">問題が解決しない場合は、後で再試行してください。</span><span class="sxs-lookup"><span data-stu-id="ef105-185">If the problem continues, please try again later.", you may need to enable a loopback exemption.</span></span>
    >
    > 1. <span data-ttu-id="ef105-186">Outlook を終了します。</span><span class="sxs-lookup"><span data-stu-id="ef105-186">Close Outlook.</span></span>
    > 1. <span data-ttu-id="ef105-187">タスク **マネージャ** を開き **、msoadfsb.exe** プロセスが実行されていないことを確認します。</span><span class="sxs-lookup"><span data-stu-id="ef105-187">Open the **Task Manager** and ensure that the **msoadfsb.exe** process is not running.</span></span>
    > 1. <span data-ttu-id="ef105-188">次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="ef105-188">Run the following command.</span></span>
    >
    >    ```command&nbsp;line
    >    call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_http___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
    >    ```
    >
    > 1. <span data-ttu-id="ef105-189">Outlook を再起動します。</span><span class="sxs-lookup"><span data-stu-id="ef105-189">Restart Outlook.</span></span>

## <a name="debug"></a><span data-ttu-id="ef105-190">Debug</span><span class="sxs-lookup"><span data-stu-id="ef105-190">Debug</span></span>

<span data-ttu-id="ef105-191">アドインで起動イベント処理を変更する場合は、次の点に注意する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ef105-191">As you make changes to launch-event handling in your add-in, you should be aware that:</span></span>

- <span data-ttu-id="ef105-192">マニフェストを更新した場合は、 [アドインを削除](sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in) してからサイドロードし直します。</span><span class="sxs-lookup"><span data-stu-id="ef105-192">If you updated the manifest, [remove the add-in](sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in) then sideload it again.</span></span>
- <span data-ttu-id="ef105-193">マニフェスト以外のファイルに変更を加えた場合は、WindowsでOutlookを閉じて再度開くか、web 上でOutlookを実行しているブラウザー タブを更新します。</span><span class="sxs-lookup"><span data-stu-id="ef105-193">If you made changes to files other than the manifest, close and reopen Outlook on Windows, or refresh the browser tab running Outlook on the web.</span></span>

<span data-ttu-id="ef105-194">独自の機能を実装する場合は、コードのデバッグが必要になる場合があります。</span><span class="sxs-lookup"><span data-stu-id="ef105-194">While implementing your own functionality, you may need to debug your code.</span></span> <span data-ttu-id="ef105-195">イベント ベースのアドインのアクティブ化をデバッグする方法については、「[イベント ベースのアドインOutlookデバッグ](debug-autolaunch.md)する 」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ef105-195">For guidance on how to debug event-based add-in activation, see [Debug your event-based Outlook add-in](debug-autolaunch.md).</span></span>

<span data-ttu-id="ef105-196">ランタイム ログは、Windowsでもこの機能で使用できます。</span><span class="sxs-lookup"><span data-stu-id="ef105-196">Runtime logging is also available for this feature on Windows.</span></span> <span data-ttu-id="ef105-197">詳細については、「 [ランタイム ログを使用したアドインのデバッグ](../testing/runtime-logging.md#runtime-logging-on-windows)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ef105-197">For more information, see [Debug your add-in with runtime logging](../testing/runtime-logging.md#runtime-logging-on-windows).</span></span>

## <a name="deploy-to-users"></a><span data-ttu-id="ef105-198">ユーザーへの展開</span><span class="sxs-lookup"><span data-stu-id="ef105-198">Deploy to users</span></span>

<span data-ttu-id="ef105-199">Microsoft 365管理センターからマニフェストをアップロードすることで、イベントベースのアドインを展開できます。</span><span class="sxs-lookup"><span data-stu-id="ef105-199">You can deploy event-based add-ins by uploading the manifest through the Microsoft 365 admin center.</span></span> <span data-ttu-id="ef105-200">管理ポータルで、ナビゲーション ウィンドウの **設定** セクションを展開し、[**統合アプリ**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="ef105-200">In the admin portal, expand the **Settings** section in the navigation pane then select **Integrated apps**.</span></span> <span data-ttu-id="ef105-201">[**統合アプリ**] ページで、[**カスタム アプリのアップロード]** アクションを選択します。</span><span class="sxs-lookup"><span data-stu-id="ef105-201">On the **Integrated apps** page, choose the **Upload custom apps** action.</span></span>

![Microsoft 365管理センターの [統合アプリ] ページのスクリーンショット (アップロードカスタム アプリ アクションなど)](../images/outlook-deploy-event-based-add-ins.png)

<span data-ttu-id="ef105-203">AppSource ストアとインクライアント ストア: イベント ベースのアドインを展開したり、既存のアドインを更新してイベント ベースのアクティブ化機能を含める機能は、すぐに利用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="ef105-203">AppSource and inclient stores: The ability to deploy event-based add-ins or update existing add-ins to include the event-based activation feature should be available soon.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ef105-204">イベントベースのアドインは、管理者が管理する展開のみに制限されます。</span><span class="sxs-lookup"><span data-stu-id="ef105-204">Event-based add-ins are restricted to admin-managed deployments only.</span></span> <span data-ttu-id="ef105-205">現時点では、ユーザーは AppSource ストアまたはインクライアント ストアからイベントベースのアドインを取得できません。</span><span class="sxs-lookup"><span data-stu-id="ef105-205">For now, users can't get event-based add-ins from AppSource or inclient stores.</span></span>

## <a name="event-based-activation-behavior-and-limitations"></a><span data-ttu-id="ef105-206">イベントベースのアクティブ化の動作と制限</span><span class="sxs-lookup"><span data-stu-id="ef105-206">Event-based activation behavior and limitations</span></span>

<span data-ttu-id="ef105-207">アドインの起動イベント ハンドラーは、短時間で軽量で、できるだけ非侵襲的であることが予想されます。</span><span class="sxs-lookup"><span data-stu-id="ef105-207">Add-in launch-event handlers are expected to be short-running, lightweight, and as noninvasive as possible.</span></span> <span data-ttu-id="ef105-208">アクティブ化後、アドインは約 300 秒以内にタイムアウトし、イベント ベースのアドインの実行に許容される最大時間が長くなります。アドインが起動イベントの処理を完了したことを知らせるために、関連付けられたハンドラーがメソッドを呼び出すことをお勧 `event.completed` めします。</span><span class="sxs-lookup"><span data-stu-id="ef105-208">After activation, your add-in will time out within approximately 300 seconds, the maximum length of time allowed for running event-based add-ins. To signal that your add-in has completed processing a launch event, we recommend you have the associated handler call the `event.completed` method.</span></span> <span data-ttu-id="ef105-209">(ステートメントの後に含まれるコード `event.completed` は、実行が保証されないことに注意してください。アドインが処理するイベントがトリガーされるたびに、アドインが再アクティブ化され、関連付けられたイベント ハンドラーが実行され、タイムアウト ウィンドウがリセットされます。</span><span class="sxs-lookup"><span data-stu-id="ef105-209">(Note that code included after the `event.completed` statement is not guaranteed to run.) Each time an event that your add-in handles is triggered, the add-in is reactivated and runs the associated event handler, and the timeout window is reset.</span></span> <span data-ttu-id="ef105-210">アドインはタイムアウト後に終了するか、ユーザーが作成ウィンドウを閉じるか、アイテムを送信します。</span><span class="sxs-lookup"><span data-stu-id="ef105-210">The add-in ends after it times out, or the user closes the compose window or sends the item.</span></span>

<span data-ttu-id="ef105-211">ユーザーが同じイベントにサブスクライブした複数のアドインを持っている場合、Outlook プラットフォームはアドインを順不同で起動します。</span><span class="sxs-lookup"><span data-stu-id="ef105-211">If the user has multiple add-ins that subscribed to the same event, the Outlook platform launches the add-ins in no particular order.</span></span> <span data-ttu-id="ef105-212">現在、アクティブに実行できるイベントベースのアドインは 5 つだけです。</span><span class="sxs-lookup"><span data-stu-id="ef105-212">Currently, only five event-based add-ins can be actively running.</span></span>

<span data-ttu-id="ef105-213">ユーザーは、アドインの実行を開始した現在のメール アイテムを切り替えたり、移動したりできます。</span><span class="sxs-lookup"><span data-stu-id="ef105-213">The user can switch or navigate away from the current mail item where the add-in started running.</span></span> <span data-ttu-id="ef105-214">起動されたアドインは、バックグラウンドで操作を終了します。</span><span class="sxs-lookup"><span data-stu-id="ef105-214">The add-in that was launched will finish its operation in the background.</span></span>

<span data-ttu-id="ef105-215">UI を変更または変更する一部のOffice.js API は、イベント ベースのアドインからは許可されません。ブロックされた API は次のとおりです。</span><span class="sxs-lookup"><span data-stu-id="ef105-215">Some Office.js APIs that change or alter the UI are not allowed from event-based add-ins. The following are the blocked APIs:</span></span>

- <span data-ttu-id="ef105-216">以下 `OfficeRuntime.auth` :</span><span class="sxs-lookup"><span data-stu-id="ef105-216">Under `OfficeRuntime.auth`:</span></span>
  - <span data-ttu-id="ef105-217">`getAccessToken`(Windowsのみ)</span><span class="sxs-lookup"><span data-stu-id="ef105-217">`getAccessToken` (Windows only)</span></span>
- <span data-ttu-id="ef105-218">以下 `Office.context.auth` :</span><span class="sxs-lookup"><span data-stu-id="ef105-218">Under `Office.context.auth`:</span></span>
  - `getAccessToken`
  - `getAccessTokenAsync`
- <span data-ttu-id="ef105-219">以下 `Office.context.mailbox` :</span><span class="sxs-lookup"><span data-stu-id="ef105-219">Under `Office.context.mailbox`:</span></span>
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- <span data-ttu-id="ef105-220">以下 `Office.context.mailbox.item` :</span><span class="sxs-lookup"><span data-stu-id="ef105-220">Under `Office.context.mailbox.item`:</span></span>
  - `close`
- <span data-ttu-id="ef105-221">以下 `Office.context.ui` :</span><span class="sxs-lookup"><span data-stu-id="ef105-221">Under `Office.context.ui`:</span></span>
  - `displayDialogAsync`
  - `messageParent`

## <a name="see-also"></a><span data-ttu-id="ef105-222">関連項目</span><span class="sxs-lookup"><span data-stu-id="ef105-222">See also</span></span>

- [<span data-ttu-id="ef105-223">Outlook アドインのマニフェスト</span><span class="sxs-lookup"><span data-stu-id="ef105-223">Outlook add-in manifests</span></span>](manifests.md)
- [<span data-ttu-id="ef105-224">イベント ベースのアドインをデバッグする方法</span><span class="sxs-lookup"><span data-stu-id="ef105-224">How to debug event-based add-ins</span></span>](debug-autolaunch.md)
