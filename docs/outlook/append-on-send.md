---
title: Outlook アドインに送信時の追加機能を実装する
description: Outlook アドインに送信時追加機能を実装する方法について学習します。
ms.topic: article
ms.date: 02/01/2021
localization_priority: Normal
ms.openlocfilehash: 8b69fbbaef1d0f060f0675fe5c4948a70d935b7a
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234290"
---
# <a name="implement-append-on-send-in-your-outlook-add-in"></a><span data-ttu-id="91a80-103">Outlook アドインに送信時の追加機能を実装する</span><span class="sxs-lookup"><span data-stu-id="91a80-103">Implement append-on-send in your Outlook add-in</span></span>

<span data-ttu-id="91a80-104">このチュートリアルの最後には、メッセージの送信時に免責事項を挿入できる Outlook アドインがあります。</span><span class="sxs-lookup"><span data-stu-id="91a80-104">By the end of this walkthrough, you'll have an Outlook add-in that can insert a disclaimer when a message is sent.</span></span>

> [!NOTE]
> <span data-ttu-id="91a80-105">この機能のサポートは、要件セット 1.9 で導入されました。</span><span class="sxs-lookup"><span data-stu-id="91a80-105">Support for this feature was introduced in requirement set 1.9.</span></span> <span data-ttu-id="91a80-106">この要件セットをサポートする [クライアントおよびプラットフォーム](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="91a80-106">See [clients and platforms](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="set-up-your-environment"></a><span data-ttu-id="91a80-107">環境を設定する</span><span class="sxs-lookup"><span data-stu-id="91a80-107">Set up your environment</span></span>

<span data-ttu-id="91a80-108">Outlook クイック [スタートを完了](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) します。このクイック スタートでは、アドイン用の Yeoman ジェネレーターを使用してアドイン Office作成します。</span><span class="sxs-lookup"><span data-stu-id="91a80-108">Complete the [Outlook quick start](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) which creates an add-in project with the Yeoman generator for Office Add-ins.</span></span>

## <a name="configure-the-manifest"></a><span data-ttu-id="91a80-109">マニフェストを構成する</span><span class="sxs-lookup"><span data-stu-id="91a80-109">Configure the manifest</span></span>

<span data-ttu-id="91a80-110">アドインで送信時追加機能を有効にするには `AppendOnSend` [、ExtendedPermissions](../reference/manifest/extendedpermissions.md)のコレクションにアクセス許可を含める必要があります。</span><span class="sxs-lookup"><span data-stu-id="91a80-110">To enable the append-on-send feature in your add-in, you must include the `AppendOnSend` permission in the collection of [ExtendedPermissions](../reference/manifest/extendedpermissions.md).</span></span>

<span data-ttu-id="91a80-111">このシナリオでは、[操作の実行] ボタンを選択して関数を実行する代わりに、関数 `action` を実行 `appendOnSend` します。</span><span class="sxs-lookup"><span data-stu-id="91a80-111">For this scenario, instead of running the `action` function on choosing the **Perform an action** button, you'll be running the `appendOnSend` function.</span></span>

1. <span data-ttu-id="91a80-112">コード エディターで、クイック スタート プロジェクトを開きます。</span><span class="sxs-lookup"><span data-stu-id="91a80-112">In your code editor, open the quick start project.</span></span>

1. <span data-ttu-id="91a80-113">プロジェクトの **manifest.xml** にある新しいファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="91a80-113">Open the **manifest.xml** file located at the root of your project.</span></span>

1. <span data-ttu-id="91a80-114">ノード全体 (開いているタグと閉じるタグを含む) を選択し `<VersionOverrides>` 、次の XML に置き換えてください。</span><span class="sxs-lookup"><span data-stu-id="91a80-114">Select the entire `<VersionOverrides>` node (including open and close tags) and replace it with the following XML.</span></span>

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
> <span data-ttu-id="91a80-115">Outlook アドインのマニフェストの詳細については、Outlook アドインの [マニフェストを参照してください](manifests.md)。</span><span class="sxs-lookup"><span data-stu-id="91a80-115">To learn more about manifests for Outlook add-ins, see [Outlook add-in manifests](manifests.md).</span></span>

## <a name="implement-append-on-send-handling"></a><span data-ttu-id="91a80-116">送信時の追加処理を実装する</span><span class="sxs-lookup"><span data-stu-id="91a80-116">Implement append-on-send handling</span></span>

<span data-ttu-id="91a80-117">次に、送信イベントに追加を実装します。</span><span class="sxs-lookup"><span data-stu-id="91a80-117">Next, implement appending on the send event.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="91a80-118">アドインで送信[ `ItemSend` ](outlook-on-send-addins.md)時イベント処理も実装している場合は、送信時ハンドラーを呼び出してエラーを返します。このシナリオは `AppendOnSendAsync` サポートされていません。</span><span class="sxs-lookup"><span data-stu-id="91a80-118">If your add-in also implements [on-send event handling using `ItemSend`](outlook-on-send-addins.md), calling `AppendOnSendAsync` in the on-send handler returns an error as this scenario isn't supported.</span></span>

<span data-ttu-id="91a80-119">このシナリオでは、ユーザーが送信するときに免責事項をアイテムに追加する方法を実装します。</span><span class="sxs-lookup"><span data-stu-id="91a80-119">For this scenario, you'll implement appending a disclaimer to the item when the user sends.</span></span>

1. <span data-ttu-id="91a80-120">同じクイック スタート プロジェクトから、コード エディターで **ファイル ./src/commands/commands.js** を開きます。</span><span class="sxs-lookup"><span data-stu-id="91a80-120">From the same quick start project, open the file **./src/commands/commands.js** in your code editor.</span></span>

1. <span data-ttu-id="91a80-121">関数の `action` 後に、次の JavaScript 関数を挿入します。</span><span class="sxs-lookup"><span data-stu-id="91a80-121">After the `action` function, insert the following JavaScript function.</span></span>

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

1. <span data-ttu-id="91a80-122">ファイルの最後に、次のステートメントを追加します。</span><span class="sxs-lookup"><span data-stu-id="91a80-122">At the end of the file, add the following statement.</span></span>

    ```js
    g.appendDisclaimerOnSend = appendDisclaimerOnSend;
    ```

## <a name="try-it-out"></a><span data-ttu-id="91a80-123">試してみる</span><span class="sxs-lookup"><span data-stu-id="91a80-123">Try it out</span></span>

1. <span data-ttu-id="91a80-124">プロジェクトのルート ディレクトリから次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="91a80-124">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="91a80-125">このコマンドを実行すると、ローカル Web サーバーが実行されていない場合に開始され、アドインがサイドロードされます。</span><span class="sxs-lookup"><span data-stu-id="91a80-125">When you run this command, the local web server will start if it's not already running and your add-in will be sideloaded.</span></span> 

    ```command&nbsp;line
    npm start
    ```

1. <span data-ttu-id="91a80-126">新しいメッセージを作成し、自分自身を [To] 行に **追加** します。</span><span class="sxs-lookup"><span data-stu-id="91a80-126">Create a new message, and add yourself to the **To** line.</span></span>

1. <span data-ttu-id="91a80-127">リボンまたはオーバーフロー メニューで、[操作の実行 **] を選択します**。</span><span class="sxs-lookup"><span data-stu-id="91a80-127">From the ribbon or overflow menu, choose **Perform an action**.</span></span>

1. <span data-ttu-id="91a80-128">メッセージを送信し、受信トレイフォルダーまたは送信アイテム フォルダー **から** メッセージを開き、追加された免責事項を表示します。</span><span class="sxs-lookup"><span data-stu-id="91a80-128">Send the message, then open it from your **Inbox** or **Sent Items** folder to view the appended disclaimer.</span></span>

    ![Outlook on the web で送信時に免責事項が追加されたメッセージ例のスクリーンショット。](../images/outlook-web-append-disclaimer.png)

## <a name="see-also"></a><span data-ttu-id="91a80-130">関連項目</span><span class="sxs-lookup"><span data-stu-id="91a80-130">See also</span></span>

[<span data-ttu-id="91a80-131">Outlook アドインのマニフェスト</span><span class="sxs-lookup"><span data-stu-id="91a80-131">Outlook add-in manifests</span></span>](manifests.md)
