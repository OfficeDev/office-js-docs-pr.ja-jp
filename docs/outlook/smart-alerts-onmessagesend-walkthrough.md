---
title: スマート アラートと OnMessageSend イベントを Outlookで使用する (プレビュー)
description: イベント ベースのライセンス認証を使用して、Outlookで送信メッセージ イベントを処理する方法について学習します。
ms.topic: article
ms.date: 03/03/2022
ms.localizationpriority: medium
ms.openlocfilehash: dba12ba6ae667f3f5db740495a58ffc425d3aef3
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340849"
---
# <a name="use-smart-alerts-and-the-onmessagesend-event-in-your-outlook-add-in-preview"></a>スマート アラートと OnMessageSend イベントを Outlookで使用する (プレビュー)

この`OnMessageSend`イベントでは、スマート アラートを利用して、ユーザーがメッセージ内で [送信]  を選択した後にロジックをOutlookします。 イベント ハンドラーを使用すると、ユーザーが送信される前に電子メールを改善する機会をユーザーに提供できます。 イベント `OnAppointmentSend` は似ていますが、予定に適用されます。

このチュートリアルの最後には、メッセージが送信されるたびに実行されるアドインが作成され、ユーザーが電子メールに記載されているドキュメントまたは画像を追加するのを忘れた場合にチェックされます。

> [!IMPORTANT]
> イベント`OnMessageSend`は`OnAppointmentSend`プレビューでのみ利用可能で、Microsoft 365のサブスクリプションOutlookでWindows。 詳細については、「How [to preview」を参照してください](autolaunch.md#how-to-preview)。 プレビュー イベントは、実稼働アドインでは使用できません。

## <a name="prerequisites"></a>前提条件

イベント `OnMessageSend` は、イベント ベースのアクティブ化機能を使用して利用できます。 この機能、利用可能なイベント、このイベントをプレビューする方法、デバッグ、機能の制限などについては、「イベント ベースのアクティブ化のために [Outlook](autolaunch.md) アドインを構成する」を参照してください。

## <a name="set-up-your-environment"></a>環境を設定する

クイック スタート[Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator)完了し、Yeoman ジェネレーターを使用してアドイン プロジェクトを作成し、Office作成します。

## <a name="configure-the-manifest"></a>マニフェストを構成する

1. コード エディターで、クイック スタート プロジェクトを開きます。

1. プロジェクトの **manifest.xml** にあるファイルを開きます。

1. **VersionOverrides ノード** 全体 (開いているタグと閉じるタグを含む) を選択し、次の XML に置き換えて、変更を保存します。

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
               This is used by Outlook on the web and Outlook on the new Mac UI preview. -->
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
        <bt:Url id="JSRuntime.Url" DefaultValue="https://localhost:3000/launchevent.js" />
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
> - イベント **で使用可能な SendMode** オプションについては、「 `OnMessageSend` 使用可能な [SendMode オプション」を参照してください](../reference/manifest/launchevent.md#available-sendmode-options-preview)。
> - アドインのマニフェストのOutlook詳細については、「Outlook[マニフェスト」を参照してください](manifests.md)。

## <a name="implement-event-handling"></a>イベント処理の実装

選択したイベントの処理を実装する必要があります。

このシナリオでは、メッセージを送信する処理を追加します。 アドインは、メッセージ内の特定のキーワードをチェックします。 これらのキーワードが見つかった場合は、添付ファイルが含まれています。 添付ファイルがない場合、アドインは、不足している可能性のある添付ファイルを追加することをユーザーに推奨します。

1. 同じクイック スタート プロジェクトから、/src/ ディレクトリの **下に launchevent という名前** の **新しいフォルダーを作成** します。

1. **./src/launchevent フォルダーで**、次の名前の新しいファイルを **launchevent.js**。

1. コード エディターで **ファイル ./src/launchevent/launchevent.js** を開き、次の JavaScript コードを追加します。

    ```js
    /*
    * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
    * See LICENSE in the project root for license information.
    */

    function onMessageSendHandler(event) {
      Office.context.mailbox.item.body.getAsync(
        "text",
        { "asyncContext": event },
        function (asyncResult) {
          let event = asyncResult.asyncContext;
          let body = "";
          let matches;
          if (asyncResult.status !== Office.AsyncResultStatus.Failed && asyncResult.value !== undefined) {
            body = asyncResult.value;
          }

          const arrayOfTerms = ["send", "picture", "document", "attachment"];
          for (let index = 0; index < arrayOfTerms.length; index++) {
            let term = arrayOfTerms[index].trim();
            const regex = RegExp(term, 'i');
            if (regex.test(body)) {
              matches.push(term);
            }
          }

          if (matches.length > 0) {
            // Let's verify if there's an attachment!
            Office.context.mailbox.item.getAttachmentsAsync(
              { "asyncContext": event },
              function(result) {
                let event = result.asyncContext;
                if (result.value.length <= 0) {
                  const message = "Looks like you're forgetting to include an attachment?";
                  event.completed({ allowEvent: false, errorMessage: message });
                } else {
                  for (let i = 0; i < result.value.length; i++) {
                    if (result.value[i].isInline == false) {
                      event.completed({ allowEvent: true });
                      return;
                    }
                  }
      
                  const message = "Looks like you forgot to include an attachment?";
                  event.completed({ allowEvent: false, errorMessage: message });
                }
              });
            } else {
              event.completed({ allowEvent: true });
            }
          }
        );
    }

    // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
    ```

1. 変更内容を保存します。

> [!IMPORTANT]
> Windows: 現在、イベント ベースのアクティブ化の処理を実装する JavaScript ファイルではインポートはサポートされていません。

## <a name="update-webpack-config-settings"></a>Webpackの機能設定を更新する

プロジェクトの **ルートwebpack.config.js** にあるファイルを開き、次の手順を実行します。

1. オブジェクト内の `plugins` 配列を見つけて `config` 、配列の先頭にこの新しいオブジェクトを追加します。

    ```js
    new CopyWebpackPlugin({
      patterns: [
        {
          from: "./src/launchevent/launchevent.js",
          to: "launchevent.js",
        },
      ],
    }),
    ```

1. 変更内容を保存します。

## <a name="try-it-out"></a>試してみる

1. プロジェクトのルート ディレクトリで次のコマンドを実行します。 実行すると、 `npm start`ローカル Web サーバーが起動し (まだ実行されていない場合)、アドインがサイドロードされます。

    ```command&nbsp;line
    npm run build
    ```
    ```command&nbsp;line
    npm start
    ```

    > [!NOTE]
    > アドインが自動的にサイドロードされていない場合は、「[サイドロード Outlook](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually) アドイン」の手順に従ってテストを行い、Outlook でアドインを手動でサイドロードします。

1. [Outlook] でWindowsメッセージを作成し、件名を設定します。 本文に「ねえ、この犬の写真をチェックしてください!」のようなテキストを追加します。
1. メッセージを送信します。 添付ファイルの追加に関する推奨事項を示すダイアログがポップアップ表示されます。
1. 添付ファイルを追加してから、もう一度メッセージを送信します。 今回は警告が表示される必要はありません。

[!INCLUDE [Loopback exemption note](../includes/outlook-loopback-exemption.md)]

## <a name="see-also"></a>関連項目

- [Outlook アドインのマニフェスト](manifests.md)
- [イベント ベースのOutlookアドインを構成する](autolaunch.md)
- [イベント ベースのアドインをデバッグする方法](debug-autolaunch.md)
- [イベント ベースのアドインの AppSource Outlookオプション](autolaunch-store-options.md)
