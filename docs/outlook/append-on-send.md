---
title: Outlook アドインで送信時に追加を実装する (プレビュー)
description: Outlook アドインでの追加-送信機能を実装する方法について説明します。
ms.topic: article
ms.date: 05/26/2020
localization_priority: Normal
ms.openlocfilehash: b9c834778d68e50806da908732cd0c8663ec6680
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093988"
---
# <a name="implement-append-on-send-in-your-outlook-add-in-preview"></a><span data-ttu-id="770d9-103">Outlook アドインで送信時に追加を実装する (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="770d9-103">Implement append on send in your Outlook add-in (preview)</span></span>

<span data-ttu-id="770d9-104">このチュートリアルを終了すると、メッセージが送信されたときに免責事項を挿入できる Outlook アドインが作成されます。</span><span class="sxs-lookup"><span data-stu-id="770d9-104">By the end of this walkthrough, you'll have an Outlook add-in that can insert a disclaimer when a message is sent.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="770d9-105">この機能は、現在、web 上の Outlook および Microsoft 365 サブスクリプションを使用した Windows の[プレビュー](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)でサポートされています。</span><span class="sxs-lookup"><span data-stu-id="770d9-105">This feature is currently supported for [preview](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web and Windows with a Microsoft 365 subscription.</span></span> <span data-ttu-id="770d9-106">詳細については、この記事の「[投稿の追加機能をプレビューする方法」を](#how-to-preview-the-append-on-send-feature)参照してください。</span><span class="sxs-lookup"><span data-stu-id="770d9-106">See [How to preview the append-on-send feature](#how-to-preview-the-append-on-send-feature) in this article for more details.</span></span>
>
> <span data-ttu-id="770d9-107">プレビュー機能は予告なしに変更される可能性があるため、運用アドインでは使用しないでください。</span><span class="sxs-lookup"><span data-stu-id="770d9-107">Because preview features are subject to change without notice, they shouldn't be used in production add-ins.</span></span>

## <a name="how-to-preview-the-append-on-send-feature"></a><span data-ttu-id="770d9-108">投稿の追加機能をプレビューする方法</span><span class="sxs-lookup"><span data-stu-id="770d9-108">How to preview the append-on-send feature</span></span>

<span data-ttu-id="770d9-109">投稿の追加機能をお試しください。</span><span class="sxs-lookup"><span data-stu-id="770d9-109">We invite you to try out the append-on-send feature!</span></span> <span data-ttu-id="770d9-110">GitHub を通じてフィードバックを提供することによって、自分のシナリオと改善方法をお知らせください (このページの最後にある**フィードバック**セクションを参照してください)。</span><span class="sxs-lookup"><span data-stu-id="770d9-110">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="770d9-111">この機能をプレビューするには:</span><span class="sxs-lookup"><span data-stu-id="770d9-111">To preview this feature:</span></span>

- <span data-ttu-id="770d9-112">CDN の**ベータ版**ライブラリを参照し https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) ます (。</span><span class="sxs-lookup"><span data-stu-id="770d9-112">Reference the **beta** library on the CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js).</span></span> <span data-ttu-id="770d9-113">TypeScript のコンパイルおよび IntelliSense 用の[型定義ファイル](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)は、CDN と、定義[された](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts)定義ファイルにあります。</span><span class="sxs-lookup"><span data-stu-id="770d9-113">The [type definition file](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) for TypeScript compilation and IntelliSense is found at the CDN and [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts).</span></span> <span data-ttu-id="770d9-114">これらの種類は、でインストールでき `npm install --save-dev @types/office-js-preview` ます。</span><span class="sxs-lookup"><span data-stu-id="770d9-114">You can install these types with `npm install --save-dev @types/office-js-preview`.</span></span>
- <span data-ttu-id="770d9-115">Windows の場合、より新しい Office ビルドにアクセスするには、 [Office Insider プログラム](https://insider.office.com)に参加する必要がある場合があります。</span><span class="sxs-lookup"><span data-stu-id="770d9-115">For Windows, you may need to join the [Office Insider program](https://insider.office.com) to access more recent Office builds.</span></span>
- <span data-ttu-id="770d9-116">Outlook on the web の場合は、 [Microsoft 365 テナントで対象となるリリースを構成](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center)します。</span><span class="sxs-lookup"><span data-stu-id="770d9-116">For Outlook on the web, [configure targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center).</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="770d9-117">環境を設定する</span><span class="sxs-lookup"><span data-stu-id="770d9-117">Set up your environment</span></span>

<span data-ttu-id="770d9-118">Outlook の[クイックスタート](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator)に記入します。このアドインプロジェクトは、Office アドイン用の [アプリ] ジェネレーターを使用して作成されます。</span><span class="sxs-lookup"><span data-stu-id="770d9-118">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="770d9-119">マニフェストを構成する</span><span class="sxs-lookup"><span data-stu-id="770d9-119">Configure the manifest</span></span>

<span data-ttu-id="770d9-120">アドインでの追加/送信機能を有効にするには、 `AppendOnSend` [extendedpermissions](../reference/manifest/extendedpermissions.md)のコレクションにアクセス許可を含める必要があります。</span><span class="sxs-lookup"><span data-stu-id="770d9-120">To enable the append-on-send feature in your add-in, you must include the `AppendOnSend` permission in the collection of [ExtendedPermissions](../reference/manifest/extendedpermissions.md).</span></span>

<span data-ttu-id="770d9-121">このシナリオでは、[操作の `action` **実行**] ボタンを選択するときに関数を実行する代わりに、関数を実行し `appendOnSend` ます。</span><span class="sxs-lookup"><span data-stu-id="770d9-121">For this scenario, instead of running the `action` function on choosing the **Perform an action** button, you'll be running the `appendOnSend` function.</span></span>

1. <span data-ttu-id="770d9-122">コードエディターで、[クイックスタート] プロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="770d9-122">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="770d9-123">プロジェクトのルートにある**manifest.xml**ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="770d9-123">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="770d9-124">`<VersionOverrides>`ノード全体 (open タグと close タグを含む) を選択し、次の XML に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="770d9-124">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML.</span></span>

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
            <DesktopFormFactor>
              <FunctionFile resid="Commands.Url" />
              <ExtensionPoint xsi:type="MessageComposeCommandSurface">
                <OfficeTab id="TabDefault">
                  <Group id="msgComposeGroup">
                    <Label resid="GroupLabel" />
                    <Control xsi:type="Button" id="msgComposeOpenPaneButton">
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
                        <FunctionName>appendDisclaimerOnSend</FunctionName>
                      </Action>
                    </Control>
                  </Group>
                </OfficeTab>
              </ExtensionPoint>

              <!-- Configure AppointmentOrganizerCommandSurface extension point to support
              append on sending a new appointment. -->

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
        <ExtendedPermissions>
          <ExtendedPermission>AppendOnSend</ExtendedPermission>
        </ExtendedPermissions>
      </VersionOverrides>
    </VersionOverrides>
    ```

> [!TIP]
> <span data-ttu-id="770d9-125">Outlook アドインのマニフェストの詳細については、「 [outlook アドインのマニフェスト](manifests.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="770d9-125">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-append-on-send-handling"></a><span data-ttu-id="770d9-126">追加オン送信処理を実装する</span><span class="sxs-lookup"><span data-stu-id="770d9-126">Implement append-on-send handling</span></span>

<span data-ttu-id="770d9-127">次に、送信イベントに追加を実装します。</span><span class="sxs-lookup"><span data-stu-id="770d9-127">Next, implement appending on the send event.</span></span>

<span data-ttu-id="770d9-128">このシナリオでは、ユーザーが送信するときに、免責事項をアイテムに追加することを実装します。</span><span class="sxs-lookup"><span data-stu-id="770d9-128">For this scenario, you'll implement appending a disclaimer to the item when the user sends.</span></span>

1. <span data-ttu-id="770d9-129">同じクイックスタートプロジェクトから、コードエディターで **/src/commands/commands.js**を開きます。</span><span class="sxs-lookup"><span data-stu-id="770d9-129">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="770d9-130">関数の後 `action` に、次の JavaScript 関数を挿入します。</span><span class="sxs-lookup"><span data-stu-id="770d9-130">After the `action` function, insert the following JavaScript function.</span></span>

    ```js
    function appendDisclaimerOnSend(event) {
      var appendText =
        '<p style = "color:blue"> <i>This and subsequent emails on the same topic are for discussion and information purposes only. Only those matters set out in a fully executed agreement are legally binding. This email may contain confidential information and should not be shared with any third party without the prior written agreement of Contoso. If you are not the intended recipient, take no action and contact the sender immediately.<br><br>Contoso Limited (company number 01624297) is a company registered in England and Wales whose registered office is at Contoso Campus, Thames Valley Park, Reading RG6 1WG</i></p>';  
      /**
        *************************************************************
         Ideal Usage - Call the getBodyType API. Use the coercionType
         it returns as the parameter value below.
        *************************************************************
      */
      Office.context.mailbox.item.body.appendOnSendAsync(
        appendText,
        {
          coercionType: Office.CoercionType.Html
        },
        function(asyncResult) {
          console.log(asyncResult);
        }
      );

      event.completed();
    }
    ```

1. <span data-ttu-id="770d9-131">ファイルの末尾に、次のステートメントを追加します。</span><span class="sxs-lookup"><span data-stu-id="770d9-131">At the end of the file, add the following statement.</span></span>

    ```js
    g.appendDisclaimerOnSend = appendDisclaimerOnSend;
    ```

## <a name="try-it-out"></a><span data-ttu-id="770d9-132">試してみる</span><span class="sxs-lookup"><span data-stu-id="770d9-132">Try it out</span></span>

1. <span data-ttu-id="770d9-133">プロジェクトのルート ディレクトリから次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="770d9-133">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="770d9-134">このコマンドを実行すると、ローカル web サーバーがまだ実行されていない場合は起動します。</span><span class="sxs-lookup"><span data-stu-id="770d9-134">When you run this command, the local web server will start if it's not already running.</span></span>

    ```command&nbsp;line
    npm run dev-server
    ```

1. <span data-ttu-id="770d9-135">「[テスト用に Outlook アドインをサイドロード](sideload-outlook-add-ins-for-testing.md)する」の手順に従います。</span><span class="sxs-lookup"><span data-stu-id="770d9-135">Follow the instructions in [Sideload Outlook add-ins for testing](sideload-outlook-add-ins-for-testing.md).</span></span>

1. <span data-ttu-id="770d9-136">新しいメッセージを作成し、[**宛先**] 行に自分を追加します。</span><span class="sxs-lookup"><span data-stu-id="770d9-136">Create a new message, and add yourself to the **To** line.</span></span>

1. <span data-ttu-id="770d9-137">リボンまたはオーバーフローメニューから、[**アクションを実行する**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="770d9-137">From the ribbon or overflow menu, choose **Perform an action**.</span></span>

1. <span data-ttu-id="770d9-138">メッセージを送信し、**受信トレイ**または**送信済みアイテム**フォルダーから開いて、追加の免責事項を表示します。</span><span class="sxs-lookup"><span data-stu-id="770d9-138">Send the message, then open it from your **Inbox** or **Sent Items** folder to view the appended disclaimer.</span></span>

    ![Web 上の Outlook で送信に追加された免責事項を含むメッセージ例のスクリーンショット。](../images/outlook-web-append-disclaimer.png)

## <a name="see-also"></a><span data-ttu-id="770d9-140">関連項目</span><span class="sxs-lookup"><span data-stu-id="770d9-140">See also</span></span>

[<span data-ttu-id="770d9-141">Outlook アドインのマニフェスト</span><span class="sxs-lookup"><span data-stu-id="770d9-141">Outlook add-in manifests</span></span>](manifests.md)
