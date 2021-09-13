---
title: アドインに append-on-send をOutlookする
description: アドインに追加送信機能を実装する方法についてOutlookします。
ms.topic: article
ms.date: 02/01/2021
ms.localizationpriority: medium
ms.openlocfilehash: ff3411ba2527e0b6c99e5e5674811ff76e240ccf
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154237"
---
# <a name="implement-append-on-send-in-your-outlook-add-in"></a>アドインに append-on-send をOutlookする

このチュートリアルの最後には、メッセージが送信Outlook免責事項を挿入できる新しいアドインがあります。

> [!NOTE]
> この機能のサポートは、要件セット 1.9 で導入されました。 この要件セットをサポートする [クライアントおよびプラットフォーム](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) を参照してください。

## <a name="set-up-your-environment"></a>環境を設定する

クイック スタート[Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator)完了し、Yeoman ジェネレーターを使用してアドイン プロジェクトを作成し、Office作成します。

## <a name="configure-the-manifest"></a>マニフェストを構成する

アドインで追加送信機能を有効にするには `AppendOnSend` [、ExtendedPermissions](../reference/manifest/extendedpermissions.md)のコレクションにアクセス許可を含める必要があります。

このシナリオでは、[アクションの実行] ボタンを選択して関数を実行する代わりに、関数 `action` を実行 `appendOnSend` します。

1. コード エディターで、クイック スタート プロジェクトを開きます。

1. プロジェクトの **manifest.xml** にあるファイルを開きます。

1. ノード全体 (開 `<VersionOverrides>` くタグと閉じるタグを含む) を選択し、次の XML に置き換えてください。

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
> アドインのマニフェストのOutlook詳細については、「Outlook[マニフェスト」を参照してください](manifests.md)。

## <a name="implement-append-on-send-handling"></a>append-on-send 処理の実装

次に、送信イベントに追加を実装します。

> [!IMPORTANT]
> アドインで[on-send `ItemSend` ](outlook-on-send-addins.md)イベント処理も実装している場合は、このシナリオがサポートされていないので、オン送信ハンドラーを呼び出してエラー `AppendOnSendAsync` を返します。

このシナリオでは、ユーザーが送信するときにアイテムに免責事項を追加する方法を実装します。

1. 同じクイック スタート プロジェクトで、コード エディター **で ./src/commands/commands.js** ファイルを開きます。

1. 関数の `action` 後に、次の JavaScript 関数を挿入します。

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

1. ファイルの最後に、次のステートメントを追加します。

    ```js
    g.appendDisclaimerOnSend = appendDisclaimerOnSend;
    ```

## <a name="try-it-out"></a>試してみる

1. プロジェクトのルート ディレクトリから次のコマンドを実行します。 このコマンドを実行すると、ローカル Web サーバーがまだ実行されていない場合に起動し、アドインがサイドロードされます。 

    ```command&nbsp;line
    npm start
    ```

1. 新しいメッセージを作成し、[To] 行に **自分自身を追加** します。

1. リボンまたはオーバーフロー メニューから、[アクションの実行 **] を選択します**。

1. メッセージを送信し、受信トレイまたは送信アイテムフォルダーからメッセージを開き、追加された免責事項を表示します。

    ![メッセージの例のスクリーンショットで、免責事項が送信時に追加Outlook on the web。](../images/outlook-web-append-disclaimer.png)

## <a name="see-also"></a>関連項目

[Outlook アドインのマニフェスト](manifests.md)
