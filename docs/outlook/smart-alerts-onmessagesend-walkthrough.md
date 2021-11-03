---
title: スマート アラートと OnMessageSend イベントを Outlookで使用する (プレビュー)
description: イベント ベースのライセンス認証を使用して、Outlookで送信メッセージ イベントを処理する方法について学習します。
ms.topic: article
ms.date: 11/01/2021
ms.localizationpriority: medium
ms.openlocfilehash: 78e10f8609264d69ba32b78badc14c626c210d76
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/02/2021
ms.locfileid: "60681849"
---
# <a name="use-smart-alerts-and-the-onmessagesend-event-in-your-outlook-add-in-preview"></a>スマート アラートと OnMessageSend イベントを Outlookで使用する (プレビュー)

このイベントでは、スマート アラートを利用して、ユーザーがメッセージで [送信] を選択した後に `OnMessageSend` ロジックをOutlookします。  イベント ハンドラーを使用すると、ユーザーが送信される前に電子メールを改善する機会をユーザーに提供できます。 イベント `OnAppointmentSend` は似ていますが、予定に適用されます。

このチュートリアルの最後には、メッセージが送信されるたびに実行されるアドインが作成され、ユーザーが電子メールに記載されているドキュメントまたは画像を追加するのを忘れた場合にチェックされます。

> [!IMPORTANT]
> イベント `OnMessageSend` と `OnAppointmentSend` イベントはプレビューでのみ利用可能で、Microsoft 365のサブスクリプションOutlookでWindows。 詳細については、「How [to preview」を参照してください](autolaunch.md#how-to-preview)。 プレビュー イベントは、実稼働アドインでは使用できません。

## <a name="prerequisites"></a>前提条件

イベント `OnMessageSend` は、イベント ベースのアクティブ化機能を使用して利用できます。 この機能、利用可能なイベント、このイベントをプレビューする方法、デバッグ、機能の制限など、アドインを構成する方法については、「イベント ベースのアクティブ化のために[Outlook](autolaunch.md)アドインを構成する」を参照してください。

## <a name="set-up-your-environment"></a>環境を設定する

クイック スタート[Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator)完了し、Yeoman ジェネレーターを使用してアドイン プロジェクトを作成し、Office作成します。

## <a name="configure-the-manifest"></a>マニフェストを構成する

1. コード エディターで、クイック スタート プロジェクトを開きます。

1. プロジェクトの **manifest.xml** にあるファイルを開きます。

1. ノード全体 (開くタグと閉じるタグを含む) を選択し、次の XML に置き換え `<VersionOverrides>` 、変更を保存します。

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

          <!-- Enable launching the add-in on the included event. -->
          <ExtensionPoint xsi:type="LaunchEvent">
            <LaunchEvents>
              <LaunchEvent Type="OnMessageSend" FunctionName="onMessageSendHandler" SendMode="PromptUser" />
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

> [!TIP]
>
> - イベント **で使用可能な SendMode** オプション `OnMessageSend` については、「使用可能な [SendMode オプション」を参照してください](../reference/manifest/launchevent.md#available-sendmode-options-preview)。
> - アドインのマニフェストのOutlook詳細については、「Outlook[マニフェスト」を参照してください](manifests.md)。

## <a name="implement-event-handling"></a>イベント処理の実装

選択したイベントの処理を実装する必要があります。

このシナリオでは、メッセージを送信する処理を追加します。 アドインは、メッセージ内の特定のキーワードをチェックします。 これらのキーワードが見つかった場合は、添付ファイルが含まれています。 添付ファイルがない場合、アドインは、不足している可能性のある添付ファイルを追加することをユーザーに推奨します。

1. 同じクイック スタート プロジェクトで、コード エディター **で ./src/commands/commands.js** ファイルを開きます。

1. 関数の `action` 後に、次の JavaScript 関数を挿入します。

    ```js
    function onMessageSendHandler(event) {
      Office.context.mailbox.item.body.getAsync(
        "text",
        { "asyncContext": event },
        function (asyncResult) {
          var event = asyncResult.asyncContext;
          var body = "";
          if (asyncResult.status !== Office.AsyncResultStatus.Failed && asyncResult.value !== undefined) {
            body = asyncResult.value;
          }
        
          var arrayOfTerms = ["send", "picture", "document", "attachment"];
          for (var index = 0; index < arrayOfTerms.length; index++) {
            var term = arrayOfTerms[index].trim();
            const regex = RegExp(term, 'i');
            if (regex.test(body)) {
              matches.push(term);
            }
          }
        
          if (matches.length > 0) {
            // Let's verify if there's an attachment!
            Office.context.mailbox.item.getAttachmentsAsync(
              { "asyncContext": event },
              function(result){
                var event = asyncResult.asyncContext;
                if (result.value.length <= 0) {
                  var message = "Looks like you're forgetting to include an attachment?";
                  event.completed({ allowEvent: false, errorMessage: message });
                } else {
                  for (var i=0;i<result.value.length;i++) {
                    if(result.value[i].isInline == false) {
                      event.completed({ allowEvent: true });
                      return;
                    }
                  }
                    
                  var message = "Looks like you're forgetting to include an attachment?";
                  event.completed({ allowEvent: false, errorMessage: message });
                }
              });
            } else {
              event.completed({ allowEvent: true });
            }
          }
        );
    }
    ```

1. ファイルの末尾に次の JavaScript コードを追加します。

    ```js
    // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
    ```

1. 変更内容を保存します。

> [!IMPORTANT]
> Windows: 現在、イベント ベースのアクティブ化の処理を実装する JavaScript ファイルではインポートはサポートされていません。

## <a name="try-it-out"></a>試してみる

1. プロジェクトのルート ディレクトリから次のコマンドを実行します。 このコマンドを実行すると、ローカル Web サーバーが (まだ実行されていない場合) 起動し、アドインがサイドロードされます。

    ```command&nbsp;line
    npm start
    ```

    > [!NOTE]
    > アドインが自動的にサイドロードされていない場合は、サイドロード[Outlook](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually)アドインの手順に従ってテストを行い、Outlook でアドインを手動でサイドロードします。

1. [Outlook] でWindowsメッセージを作成し、件名を設定します。 本文に「ねえ、この犬の写真をチェックしてください!」のようなテキストを追加します。
1. メッセージを送信します。 添付ファイルの追加に関する推奨事項を示すダイアログがポップアップ表示されます。
1. 添付ファイルを追加してから、もう一度メッセージを送信します。 今回は警告が表示される必要はありません。

> [!NOTE]
> localhost からアドインを実行している場合は、"申し訳ありませんが *、{your-add-in-name-here}* にアクセスできませんでした。 ネットワーク接続が確立されている必要があります。 問題が解決しない場合は、後でもう一度お試しください。ループバックの除外を有効にする必要がある場合があります。
>
> 1. Outlook を終了します。
> 1. タスク マネージャー **を開** き、タスク **msoadfsb.exeが** 実行されていないか確認します。
> 1. 次のコマンドを実行します。
>
>    ```command&nbsp;line
>    call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_http___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
>    ```
>
> 1. Outlook を再起動します。

## <a name="see-also"></a>関連項目

- [Outlook アドインのマニフェスト](manifests.md)
- [イベント ベースのOutlook用にアドインを構成する](autolaunch.md)
- [イベント ベースのアドインをデバッグする方法](debug-autolaunch.md)
- [イベント ベースのアドインの AppSource Outlookオプション](autolaunch-store-options.md)
