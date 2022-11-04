---
title: Outlook アドインでスマート アラートと OnMessageSend イベントと OnAppointmentSend イベントを使用する
description: イベント ベースのアクティブ化を使用して、Outlook アドインで送信時イベントを処理する方法について説明します。
ms.topic: article
ms.date: 11/2/2022
ms.localizationpriority: medium
ms.openlocfilehash: 408c3684d325a9cbdd4a3f6e489db636ff52e028
ms.sourcegitcommit: 9c65c19298bf749836e3db1b7cf5e8c1387a2bf2
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/04/2022
ms.locfileid: "68842090"
---
# <a name="use-smart-alerts-and-the-onmessagesend-and-onappointmentsend-events-in-your-outlook-add-in"></a>Outlook アドインでスマート アラートと OnMessageSend イベントと OnAppointmentSend イベントを使用する

イベントと `OnAppointmentSend` イベントは`OnMessageSend`スマート アラートを利用します。これにより、ユーザーが Outlook メッセージまたは予定で **[送信**] を選択した後にロジックを実行できます。 イベント ハンドラーを使用すると、送信前にメールや会議出席依頼を改善する機会をユーザーに提供できます。

次のチュートリアルでは、 イベントを使用します `OnMessageSend` 。 このチュートリアルの終わりまでに、メッセージが送信されるたびに実行されるアドインが用意され、ユーザーがメールで言及したドキュメントや画像を追加し忘れたかどうかを確認できます。

> [!NOTE]
> イベントと `OnAppointmentSend` イベントは`OnMessageSend`[、要件セット 1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12) で導入されました。 この要件セットをサポートする [クライアントおよびプラットフォーム](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets) を参照してください。

## <a name="prerequisites"></a>前提条件

イベントは `OnMessageSend` 、イベント ベースのアクティブ化機能を使用して使用できます。 この機能を使用するようにアドインを構成する方法、他の使用可能なイベントを使用する方法、アドインをデバッグする方法については、「 [イベントベースのアクティブ化のために Outlook アドインを構成する](autolaunch.md)」を参照してください。

### <a name="supported-clients-and-platforms"></a>サポートされているクライアントとプラットフォーム

次の表に、スマート アラート機能でサポートされているクライアントとサーバーの組み合わせの一覧を示します。該当する場合は、累積的な更新プログラムExchange Server最低限必要です。 除外された組み合わせはサポートされていません。

|Client|Exchange Online|Exchange 2019 オンプレミス (累積的な更新プログラム 12 以降)|Exchange 2016 オンプレミス (累積的な更新プログラム 22 以降) |
|-----|-----|-----|-----|
|**Windows**<br>バージョン 2206 (ビルド 15330.20196) 以降|はい|はい|はい|
|**Mac**<br>バージョン 16.65.827.0 以降|はい|該当なし|該当なし|
|**Web ブラウザー (モダン UI)**|はい|該当なし|該当なし|
|**iOS**|該当なし|該当なし|該当なし|
|**Android**|該当なし|該当なし|該当なし|

## <a name="set-up-your-environment"></a>環境を設定する

[Office アドイン用 Yeoman ジェネレーター](../develop/yeoman-generator-overview.md)を使用してアドイン プロジェクトを作成する [Outlook クイック スタート](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator)を完了します。

## <a name="configure-the-manifest"></a>マニフェストを構成する

マニフェストを構成するには、使用しているマニフェストの種類のタブを選択します。

# <a name="xml-manifest"></a>[XML マニフェスト](#tab/xmlmanifest)

1. コード エディターで、クイック スタート プロジェクトを開きます。

1. プロジェクトのルートにある **manifest.xml** ファイルを開きます。

1. ノード全体 **\<VersionOverrides\>** (開いているタグと閉じるタグを含む) を選択し、次の XML に置き換えてから、変更を保存します。

```XML
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    <Requirements>
      <bt:Sets DefaultMinVersion="1.12">
        <bt:Set Name="Mailbox" />
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- Event-based activation happens in a lightweight runtime.-->
        <Runtimes>
          <!-- HTML file including reference to or inline JavaScript event handlers.
               This is used by Outlook on the web and on the new Mac UI. -->
          <Runtime resid="WebViewRuntime.Url">
            <!-- JavaScript file containing event handlers. This is used by Outlook on Windows. -->
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
        <!-- Entry needed for Outlook on Windows. -->
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
> - イベントと `OnAppointmentSend` イベントで`OnMessageSend`使用できる **SendMode** オプションについては、「[使用可能な SendMode オプション](/javascript/api/manifest/launchevent#available-sendmode-options)」を参照してください。
> - Outlook アドインのマニフェストの詳細については、「 [Outlook アドイン マニフェスト](manifests.md)」を参照してください。

# <a name="teams-manifest-developer-preview"></a>[Teams マニフェスト (開発者プレビュー)](#tab/jsonmanifest)

> [!IMPORTANT]
> スマート アラートは、 [Office アドイン用の Teams マニフェスト (プレビュー)](../develop/json-manifest-overview.md) ではまだサポートされていません。 このタブは、今後使用するために使用されます。

1. **manifest.json** ファイルを開きます。

1. "extensions.runtimes" 配列に次のオブジェクトを追加します。 このマークアップについて、次の点に注意してください。

   - メールボックス要件セットの "minVersion" は "1.12" に設定 [されます。サポートされているイベント テーブル](autolaunch.md#supported-events) では、これがイベントをサポート `OnMessageSend` する要件セットの最小バージョンであることが指定されているためです。
   - ランタイムの "id" は、わかりやすい名前 "autorun_runtime" に設定されます。
   - "code" プロパティには、HTML ファイルに設定された子 "page" プロパティと、JavaScript ファイルに設定された子 "script" プロパティがあります。 これらのファイルは、後の手順で作成または編集します。 Office では、プラットフォームに応じて、これらの値の 1 つまたはもう 1 つを使用します。
       - Office on Windows では、JavaScript 専用ランタイムでイベント ハンドラーが実行され、JavaScript ファイルが直接読み込まれます。
       - Office on Mac と Web では、ブラウザー ランタイムでハンドラーが実行され、HTML ファイルが読み込まれます。 そのファイルには、JavaScript ファイルを `<script>` 読み込むタグが含まれています。
     詳細については、「 [Office アドインのランタイム](../testing/runtimes.md)」を参照してください。
   - "lifetime" プロパティは "short" に設定されています。これは、イベントがトリガーされたときにランタイムが起動し、ハンドラーが完了したときにシャットダウンすることを意味します。 (まれに、ハンドラーが完了する前にランタイムがシャットダウンする場合があります。 [「Office アドインのランタイム」を](../testing/runtimes.md)参照してください)。
   - イベントのハンドラーを実行する `OnMessageSend` アクションがあります。 ハンドラー関数は、後の手順で作成します。

    ```json
     {
        "requirements": {
            "capabilities": [
                {
                    "name": "Mailbox",
                    "minVersion": "1.12"
                }
            ]
        },
        "id": "autorun_runtime",
        "type": "general",
        "code": {
            "page": "https://localhost:3000/commands.html",
            "script": "https://localhost:3000/launchevent.js"
        },
        "lifetime": "short",
        "actions": [
            {
                "id": "onMessageSendHandler",
                "type": "executeFunction",
                "displayName": "onMessageSendHandler"
            }
        ]
    }
    ```

1. "extensions" 配列の オブジェクトのプロパティとして、次の "autoRunEvents" 配列を追加します。

    ```json
    "autoRunEvents": [
    
    ]
    ```

1. "autoRunEvents" 配列に次のオブジェクトを追加します。 このコードについては、次の点に注意してください。

   - イベント オブジェクトは、ハンドラー関数をイベントに `OnMessageSend` 割り当てます ( [サポートされているイベント テーブル](autolaunch.md#supported-events)で説明されているように、イベントの Teams マニフェスト名 "messageSending" を使用)。 "actionId" で指定する関数名は、前の手順の "actions" 配列の オブジェクトの "id" プロパティで使用される名前と一致する必要があります。
   - "sendMode" オプションは "promptUser" に設定されています。 つまり、メッセージがアドインが送信のために設定する条件を満たしていない場合、ユーザーは送信を取り消すか送信するように求められます。

    ```json
      {
          "requirements": {
              "capabilities": [
                  {
                      "name": "Mailbox",
                      "minVersion": "1.12"
                  }
              ],
              "scopes": [
                  "mail"
              ]
          },
          "events": [
            {
                "type": "messageSending",
                "actionId": "onMessageSendHandler",
                "options": {
                    "sendMode": "promptUser"
                }
            }
          ]
      }
    ```

---

## <a name="implement-event-handling"></a>イベント処理を実装する

選択したイベントの処理を実装する必要があります。

このシナリオでは、メッセージを送信するための処理を追加します。 アドインは、メッセージ内の特定のキーワードを確認します。 これらのキーワードのいずれかが見つかった場合は、添付ファイルがあるかどうかを確認します。 添付ファイルがない場合、アドインは、不足している可能性のある添付ファイルを追加することをお勧めします。

1. 同じクイック スタート プロジェクトから、**./src** ディレクトリの下に **launchevent** という名前の新しいフォルダーを作成します。

1. **./src/launchevent** フォルダーで、 という名前の新しいファイル **launchevent.js** 作成します。

1. コード エディターで **./src/launchevent/launchevent.js** ファイルを開き、次の JavaScript コードを追加します。

    ```js
    /*
    * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
    * See LICENSE in the project root for license information.
    */

    function onMessageSendHandler(event) {
      Office.context.mailbox.item.body.getAsync(
        "text",
        { asyncContext: event },
        getBodyCallback
      );
    }

    function getBodyCallback(asyncResult){
      let event = asyncResult.asyncContext;
      let body = "";
      if (asyncResult.status !== Office.AsyncResultStatus.Failed && asyncResult.value !== undefined) {
        body = asyncResult.value;
      } else {
        let message = "Failed to get body text";
        console.error(message);
        event.completed({ allowEvent: false, errorMessage: message });
        return;
      }

      let matches = hasMatches(body);
      if (matches) {
        Office.context.mailbox.item.getAttachmentsAsync(
          { asyncContext: event },
          getAttachmentsCallback);
      } else {
        event.completed({ allowEvent: true });
      }
    }

    function hasMatches(body) {
      if (body == null || body == "") {
        return false;
      }

      const arrayOfTerms = ["send", "picture", "document", "attachment"];
      for (let index = 0; index < arrayOfTerms.length; index++) {
        const term = arrayOfTerms[index].trim();
        const regex = RegExp(term, 'i');
        if (regex.test(body)) {
          return true;
        }
      }

      return false;
    }

    function getAttachmentsCallback(asyncResult) {
      let event = asyncResult.asyncContext;
      if (asyncResult.value.length > 0) {
        for (let i = 0; i < asyncResult.value.length; i++) {
          if (asyncResult.value[i].isInline == false) {
            event.completed({ allowEvent: true });
            return;
          }
        }

        event.completed({ allowEvent: false, errorMessage: "Looks like you forgot to include an attachment?" });
      } else {
        event.completed({ allowEvent: false, errorMessage: "Looks like you're forgetting to include an attachment?" });
      }
    }

    // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
    ```

## <a name="update-the-commands-html-file"></a>コマンド HTML ファイルを更新する

1. **./src/commands** フォルダーで、 **commands.html** を開きます。

1. 終了 **ヘッド** タグ (`</head>`) の直前に、イベント処理 JavaScript コードのスクリプト エントリを追加します。

   ```js
   <script type="text/javascript" src="../launchevent/launchevent.js"></script> 
   ```

1. 変更内容を保存します。

## <a name="update-webpack-config-settings"></a>Webpackの機能設定を更新する

1. プロジェクトのルート ディレクトリにある **webpack.config.js** ファイルを開き、次の手順を実行します。

1. オブジェクト内の `plugins` 配列を `config` 見つけて、この新しいオブジェクトを配列の先頭に追加します。

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

1. プロジェクトのルート ディレクトリで次のコマンドを実行します。 を実行 `npm start`すると、ローカル Web サーバーが起動し (まだ実行されていない場合)、アドインがサイドロードされます。

    ```command&nbsp;line
    npm run build
    ```

    ```command&nbsp;line
    npm start
    ```

    > [!NOTE]
    > アドインが自動的にサイドロードされなかった場合は、「 [テスト用の Outlook アドインをサイドロード](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually) する」の手順に従って、Outlook でアドインを手動でサイドロードします。

1. Outlook on Windows で、新しいメッセージを作成し、件名を設定します。 本文に、「ねえ、私の犬のこの写真をチェックしてください!」のようなテキストを追加します。
1. メッセージを送信します。 添付ファイルを追加するための推奨事項が表示されたダイアログが表示されます。

    ![ユーザーに添付ファイルを含めさせるダイアログ。](../images/outlook-win-smart-alert.png)

1. 添付ファイルを追加し、メッセージをもう一度送信します。 今回はアラートは表示されません。

## <a name="deploy-to-users"></a>ユーザーにデプロイする

他のイベント ベースのアドインと同様に、スマート アラート機能を使用するアドインは、組織の管理者が展開する必要があります。 Microsoft 365 管理センターを使用してアドインを展開する方法のガイダンスについては、「[イベント ベースのアクティブ化のために Outlook アドインを構成](autolaunch.md#deploy-to-users)する」の「**ユーザーに展開** する」セクションを参照してください。

> [!IMPORTANT]
> スマート アラート機能を使用するアドインは、マニフェストの [SendMode プロパティ](/javascript/api/manifest/launchevent#available-sendmode-options)が または `PromptUser` オプションに設定されている場合にのみ AppSource に`SoftBlock`発行できます。 アドインの **SendMode** プロパティが に `Block`設定されている場合は、AppSource の検証に失敗するため、組織の管理者のみが展開できます。 イベント ベースのアドインを AppSource に発行する方法の詳細については、「 [イベントベースの Outlook アドインの AppSource 一覧オプション](autolaunch-store-options.md)」を参照してください。

## <a name="smart-alerts-feature-behavior-and-scenarios"></a>スマート アラート機能の動作とシナリオ

**SendMode** オプションの説明と、使用するタイミングに関する推奨事項については、「[使用可能な SendMode オプション](/javascript/api/manifest/launchevent#available-sendmode-options)」を参照してください。 次に、特定のシナリオに対する機能の動作について説明します。

### <a name="add-in-is-unavailable"></a>アドインは使用できません

メッセージまたは予定の送信中にアドインが使用できない場合 (たとえば、アドインの読み込みを妨げるエラーが発生した場合)、ユーザーにアラートが表示されます。 ユーザーが使用できるオプションは、アドインに適用される **SendMode** オプションによって異なります。

または `SoftBlock` オプションを`PromptUser`使用する場合、ユーザーは [**Send Anyway]\(とにかく送信**\) を選択して、アドインをチェックせずにアイテムを送信するか、[**後で試** す] を選択して、アドインが再度使用可能になったときにアイテムを確認できます。

![アドインが使用できないことをユーザーに警告し、ユーザーにアイテムを今すぐ送信するオプションを表示するダイアログ。](../images/outlook-soft-block-promptUser-unavailable.png)

オプションが `Block` 使用されている場合、ユーザーはアドインが使用可能になるまでアイテムを送信できません。 (アドインが Teams マニフェスト (プレビュー) を使用している場合、オプションは `Block` サポートされていません)。

![アドインが使用できないことをユーザーに警告するダイアログ。 ユーザーは、アドインが再度使用できる場合にのみアイテムを送信できます。](../images/outlook-hard-block-unavailable.png)

### <a name="long-running-add-in-operations"></a>実行時間の長いアドイン操作

アドインが 5 秒を超えて実行されるが 5 分未満の場合、ユーザーは、アドインがメッセージまたは予定の処理に予想以上に時間がかかることを警告します。

オプションを使用する `PromptUser` 場合、ユーザーは [ **Send Anyway]\(とにかく送信\)** を選択して、アドインがチェックを完了せずにアイテムを送信できます。 または、ユーザーが [ **送信しない** ] を選択して、アドインの処理を停止することもできます。

![アドインがアイテムの処理に予想以上に時間がかかることをユーザーに通知するダイアログ。 ユーザーは、アドインがチェックを完了せずにアイテムを送信するか、アドインがアイテムの処理を停止することを選択できます。](../images/outlook-promptUser-long-running.png)

ただし、 または `Block` オプションを`SoftBlock`使用する場合、アドインが処理を完了するまで、ユーザーはアイテムを送信できません。

![アドインがアイテムの処理に予想以上に時間がかかることをユーザーに通知するダイアログ。 ユーザーは、アドインがアイテムの処理を完了するまで待ってから送信する必要があります。](../images/outlook-soft-hard-block-long-running.png)

`OnMessageSend` アドイン `OnAppointmentSend` は、実行時間が短く、軽量である必要があります。 実行時間の長い操作ダイアログを回避するには、または `OnAppointmentSend` イベントがアクティブ化される前に、他のイベントを使用して条件付きチェックを`OnMessageSend`処理します。 たとえば、ユーザーがすべてのメッセージまたは予定の添付ファイルを暗号化する必要がある場合は、 または `OnAppointmentAttachmentsChanged` イベントを`OnMessageAttachmentsChanged`使用してチェックを実行することを検討してください。

### <a name="add-in-timed-out"></a>アドインがタイムアウトしました

アドインが 5 分以上実行されると、タイムアウトになります。オプションを使用する `PromptUser` 場合、ユーザーは [ **Send Anyway]\(とにかく送信\)** を選択して、アドインがチェックを完了せずにアイテムを送信できます。 または、ユーザーは [ **送信しない**] を選択できます。

![アドイン プロセスがタイムアウトしたことをユーザーに通知するダイアログ。ユーザーは、アドインがチェックを完了せずにアイテムを送信するか、アイテムを送信しないことを選択できます。](../images/outlook-promptUser-timeout.png)

`SoftBlock`または `Block` オプションが使用されている場合、アドインがチェックを完了するまで、ユーザーはアイテムを送信できません。 ユーザーは、アドインを再アクティブ化するために、アイテムの再送信を試みる必要があります。

![アドイン プロセスがタイムアウトしたことをユーザーに通知するダイアログ。ユーザーは、メッセージまたは予定を送信する前に、アドインをアクティブ化するためにアイテムの再送信を試みる必要があります。](../images/outlook-soft-hard-block-timeout.png)

## <a name="limitations"></a>制限事項

`OnMessageSend`イベントと `OnAppointmentSend` イベントはイベント ベースのアクティブ化機能を通じてサポートされるため、これらのイベントの結果としてアクティブ化するアドインにも同じ機能制限が適用されます。 これらの制限事項の詳細については、「 [イベント ベースのアクティブ化の動作と制限事項](autolaunch.md#event-based-activation-behavior-and-limitations)」を参照してください。

これらの制約に加えて、マニフェストで宣言できるのは、 `OnMessageSend` イベントと `OnAppointmentSend` イベントのそれぞれ 1 つのインスタンスのみです。 複数または`OnAppointmentSend`イベント`OnMessageSend`が必要な場合は、それぞれを個別のアドインで宣言する必要があります。

event.completed メソッドの [errorMessage プロパティ](/javascript/api/office/office.addincommands.eventcompletedoptions) を使用して、アドインのシナリオに合わせてスマート アラート ダイアログ メッセージを変更できますが、次の内容はカスタマイズできません。

- ダイアログのタイトル バー。 アドインの名前は常にそこに表示されます。
- メッセージの形式。 たとえば、テキストのフォント サイズと色を変更したり、箇条書きを挿入したりすることはできません。
- ダイアログ オプション。 たとえば、[ **送信方法** ] オプションと **[送信しない** ] オプションは固定されており、選択した [SendMode オプション](/javascript/api/manifest/launchevent#available-sendmode-options) によって異なります。
- イベント ベースのアクティブ化処理と進行状況の情報ダイアログ。 たとえば、タイムアウト操作ダイアログや実行時間の長い操作ダイアログに表示されるテキストとオプションは変更できません。

## <a name="differences-between-smart-alerts-and-the-on-send-feature"></a>スマート アラートと送信時機能の違い

スマート アラートと [送信時機能](outlook-on-send-addins.md) を使用すると、送信前にメッセージや会議出席依頼を改善する機会がユーザーに提供されますが、スマート アラートは、ユーザーにさらなるアクションを求める柔軟性を高める新しい機能です。 2 つの機能の主な違いを次の表に示します。

> [!IMPORTANT]
> スマート アラートは、Teams マニフェスト (プレビュー) ではまだサポートされていません。 そのサポートの提供に間もなく取り組んでいます。

|属性|スマート アラート|送信中|
|-----|-----|-----|
|**サポートされる最小要件セット**|[メールボックス 1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)|[Mailbox 1.8](/javascript/api/requirement-sets/outlook/requirement-set-1.8/outlook-requirement-set-1.8)|
|**サポートされている Outlook クライアント**|-Windows<br>- Web ブラウザー (モダン UI)<br>- Mac (新しい UI)|-Windows<br>- Web ブラウザー (クラシックおよびモダン UI)<br>- Mac (クラシックおよび新しい UI) |
|**サポートされるイベント**|**XML マニフェスト**<br>- `OnMessageSend`<br>- `OnAppointmentSend`<br><br>**Teams マニフェスト (プレビュー)**<br>- "messageSending"<br>- "appointmentSending"|**XML マニフェスト**<br>- `ItemSend`<br><br>**Teams マニフェスト (プレビュー)**<br>- サポートされていません|
|**マニフェスト拡張プロパティ**|**XML マニフェスト**<br>- `LaunchEvent`<br><br>**Teams マニフェスト (プレビュー)**<br>- "autoRunEvents"|**XML マニフェスト**<br>- `Events`<br><br>**Teams マニフェスト (プレビュー)**<br>- サポートされていません|
|**サポートされている送信モード オプション**|- ユーザーにプロンプトを表示する<br>- ソフト ブロック<br>- ブロック (アドインが Teams マニフェスト (プレビュー) を使用している場合はサポートされません)|ブロック|
|**アドインでサポートされているイベントの最大数**|1 つの `OnMessageSend` イベントと 1 つの `OnAppointmentSend` イベント。|1 つの `ItemSend` イベント。|
|**アドインのデプロイ**|プロパティが または `PromptUser` オプションに設定されている場合`SendMode`、アドインを AppSource に`SoftBlock`発行できます。 それ以外の場合は、組織の管理者がアドインを展開する必要があります。|アドインを AppSource に発行することはできません。 組織の管理者が展開する必要があります。|
|**アドインのインストールの追加構成**|マニフェストがMicrosoft 365 管理センターにアップロードされた後は、追加の構成は必要ありません。|組織のコンプライアンス標準と使用される Outlook クライアントに応じて、アドインをインストールするように特定のメールボックス ポリシーを構成する必要があります。|

## <a name="see-also"></a>関連項目

- [Outlook アドインのマニフェスト](manifests.md)
- [イベント ベースのアクティブ化のために Outlook アドインを構成する](autolaunch.md)
- [イベント ベースのアドインをデバッグする方法](debug-autolaunch.md)
- [イベント ベースの Outlook アドインの AppSource 登録情報オプション](autolaunch-store-options.md)
- [Office アドインのコード サンプル: Outlook スマート アラートを使用する](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-check-item-categories)
