---
title: イベント ベースのアクティブ化のために Outlook アドインを構成する
description: イベント ベースのアクティブ化のために Outlook アドインを構成する方法について説明します。
ms.topic: article
ms.date: 10/13/2022
ms.localizationpriority: medium
ms.openlocfilehash: ce2821ed5d226ff2c6a2b3c718d5711689523ac6
ms.sourcegitcommit: d402c37fc3388bd38761fedf203a7d10fce4e899
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/21/2022
ms.locfileid: "68664680"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation"></a>イベント ベースのアクティブ化のために Outlook アドインを構成する

イベント ベースのアクティブ化機能がないと、ユーザーはアドインを明示的に起動してタスクを完了する必要があります。 この機能を使用すると、アドインは特定のイベント (特にすべてのアイテムに適用される操作) に基づいてタスクを実行できます。 作業ウィンドウと関数コマンドと統合することもできます。

このチュートリアルの終わりまでに、新しい項目が作成され、件名が設定されるたびに実行されるアドインが用意されています。

> [!NOTE]
> この機能のサポートは [要件セット 1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10) で導入され、追加のイベントが後続の要件セットで使用できるようになりました。 イベントの最小要件セットと、それをサポートするクライアントとプラットフォームの詳細については、「[Exchange サーバーと Outlook クライアントでサポートされる](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients)[サポートされるイベント](#supported-events)と要件セット」を参照してください。
>
> iOS または Android 上の Outlook では、イベント ベースのアクティブ化はサポートされていません。

## <a name="supported-events"></a>サポートされるイベント

次の表に、現在使用可能なイベントと、各イベントでサポートされているクライアントを示します。 イベントが発生すると、ハンドラーは、イベントの種類に固有の詳細を `event` 含むオブジェクトを受け取ります。 **[説明]** 列には、該当する場合は関連オブジェクトへのリンクが含まれます。

|イベント標準名</br>および XML マニフェスト名|Teams マニフェスト名|説明|最小要件セットとサポートされているクライアント|
|---|---|---|---|
|`OnNewMessageCompose`| newMessageComposeCreated |新しいメッセージの作成時 (返信、全員への返信、転送を含む) が、下書きなど編集時には含まれません。|[1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10)<br><br>- Windows<sup>1</sup><br>- Web ブラウザー<br>- 新しい Mac UI |
|`OnNewAppointmentOrganizer`|newAppointmentOrganizerCreated|新しい予定を作成するが、既存の予定を編集していない場合。|[1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10)<br><br>- Windows<sup>1</sup><br>- Web ブラウザー<br>- 新しい Mac UI |
|`OnMessageAttachmentsChanged`|messageAttachmentsChanged|メッセージの作成中に添付ファイルを追加または削除する場合。<br><br>イベント固有のデータ オブジェクト: [AttachmentsChangedEventArgs](/javascript/api/outlook/office.attachmentschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>- Web ブラウザー<br>- 新しい Mac UI|
|`OnAppointmentAttachmentsChanged`|appointmentAttachmentsChanged|予定の作成中に添付ファイルを追加または削除する場合。<br><br>イベント固有のデータ オブジェクト: [AttachmentsChangedEventArgs](/javascript/api/outlook/office.attachmentschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>- Web ブラウザー<br>- 新しい Mac UI|
|`OnMessageRecipientsChanged`|messageRecipientsChanged|メッセージの作成中に受信者を追加または削除する場合。<br><br>イベント固有のデータ オブジェクト: [RecipientsChangedEventArgs](/javascript/api/outlook/office.recipientschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>- Web ブラウザー<br>- 新しい Mac UI|
|`OnAppointmentAttendeesChanged`|appointmentAttendeesChanged|予定の作成中に出席者を追加または削除する場合。<br><br>イベント固有のデータ オブジェクト: [RecipientsChangedEventArgs](/javascript/api/outlook/office.recipientschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>- Web ブラウザー<br>- 新しい Mac UI|
|`OnAppointmentTimeChanged`|appointmentTimeChanged|予定の作成中に日付/時刻を変更する場合。<br><br>イベント固有のデータ オブジェクト: [AppointmentTimeChangedEventArgs](/javascript/api/outlook/office.appointmenttimechangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>- Web ブラウザー<br>- 新しい Mac UI|
|`OnAppointmentRecurrenceChanged`|appointmentRecurrenceChanged|予定の作成中に繰り返しの詳細を追加、変更、または削除する場合。 日付/時刻が変更されると、 `OnAppointmentTimeChanged` イベントも発生します。<br><br>イベント固有のデータ オブジェクト: [RecurrenceChangedEventArgs](/javascript/api/outlook/office.recurrencechangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>- Web ブラウザー<br>- 新しい Mac UI|
|`OnInfoBarDismissClicked`|infoBarDismissClicked|メッセージまたは予定アイテムの作成中に通知を閉じる場合。 通知を追加したアドインのみが通知されます。<br><br>イベント固有のデータ オブジェクト: [InfobarClickedEventArgs](/javascript/api/outlook/office.infobarclickedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<sup>1</sup><br>- Web ブラウザー<br>- 新しい Mac UI|
|`OnMessageSend`|messageSending|メッセージ 項目の送信時。 詳細については、 [スマート アラートのチュートリアルを](smart-alerts-onmessagesend-walkthrough.md)参照してください。|[1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)<br><br>- Windows<sup>1</sup><br>- Web ブラウザー|
|`OnAppointmentSend`|appointmentSending|予定アイテムを送信する場合。 詳細については、 [スマート アラートのチュートリアルを](smart-alerts-onmessagesend-walkthrough.md)参照してください。|[1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)<br><br>- Windows<sup>1</sup><br>- Web ブラウザー|
|`OnMessageCompose`|messageComposeOpened|新しいメッセージの作成 (返信、全員への返信、転送を含む) または下書きを編集する場合。|[1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)<br><br>- Windows<sup>1</sup><br>- Web ブラウザー|
|`OnAppointmentOrganizer`|appointmentOrganizerOpened|新しい予定を作成するか、既存の予定を編集する場合。|[1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)<br><br>- Windows<sup>1</sup><br>- Web ブラウザー|

> [!NOTE]
> Outlook on Windows の <sup>1 つの</sup>イベント ベースのアドインを実行するには、少なくとも Windows 10 バージョン 1903 (ビルド 18362) または Windows Server 2019 バージョン 1903 が必要です。

## <a name="set-up-your-environment"></a>環境を設定する

[Office アドイン用 Yeoman ジェネレーター](../develop/yeoman-generator-overview.md)を使用してアドイン プロジェクトを作成する [Outlook クイック スタート](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator)を完了します。

> [!NOTE]
> [Office アドイン用の Teams マニフェスト (プレビュー)](../develop/json-manifest-overview.md) を使用する場合は、[Outlook クイック スタートで Teams マニフェスト (プレビュー)](../quickstarts/outlook-quickstart-json-manifest.md) を使用して代替クイック スタートを完了しますが、[**試してみる**] セクションの後にすべてのセクションをスキップします。

## <a name="configure-the-manifest"></a>マニフェストを構成する

マニフェストを構成するには、使用しているマニフェストの種類のタブを選択します。

# <a name="xml-manifest"></a>[XML マニフェスト](#tab/xmlmanifest)

アドインのイベント ベースのアクティブ化を有効にするには、マニフェストのノードで [Runtimes](/javascript/api/manifest/runtimes) 要素と [LaunchEvent](/javascript/api/manifest/extensionpoint#launchevent) 拡張ポイントを構成する `VersionOverridesV1_1` 必要があります。 現時点では、 `DesktopFormFactor` サポートされている唯一のフォーム ファクターです。

1. コード エディターで、クイック スタート プロジェクトを開きます。

1. プロジェクトのルートにある **manifest.xml** ファイルを開きます。

1. ノード全体 **\<VersionOverrides\>** (開いているタグと閉じるタグを含む) を選択し、次の XML に置き換えてから、変更を保存します。

```XML
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    <Requirements>
      <bt:Sets DefaultMinVersion="1.10">
        <bt:Set Name="Mailbox" />
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- Event-based activation happens in a lightweight runtime.-->
        <Runtimes>
          <!-- HTML file including reference to or inline JavaScript event handlers.
               This is used by Outlook on the web and Outlook on the new Mac UI. -->
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

          <!-- Enable launching the add-in on the included events. -->
          <ExtensionPoint xsi:type="LaunchEvent">
            <LaunchEvents>
              <LaunchEvent Type="OnNewMessageCompose" FunctionName="onNewMessageComposeHandler"/>
              <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onNewAppointmentComposeHandler"/>
              
              <!-- Other available events -->
              <!--
              <LaunchEvent Type="OnMessageAttachmentsChanged" FunctionName="onMessageAttachmentsChangedHandler" />
              <LaunchEvent Type="OnAppointmentAttachmentsChanged" FunctionName="onAppointmentAttachmentsChangedHandler" />
              <LaunchEvent Type="OnMessageRecipientsChanged" FunctionName="onMessageRecipientsChangedHandler" />
              <LaunchEvent Type="OnAppointmentAttendeesChanged" FunctionName="onAppointmentAttendeesChangedHandler" />
              <LaunchEvent Type="OnAppointmentTimeChanged" FunctionName="onAppointmentTimeChangedHandler" />
              <LaunchEvent Type="OnAppointmentRecurrenceChanged" FunctionName="onAppointmentRecurrenceChangedHandler" />
              <LaunchEvent Type="OnInfoBarDismissClicked" FunctionName="onInfobarDismissClickedHandler" />
              <LaunchEvent Type="OnMessageSend" FunctionName="onMessageSendHandler" SendMode="PromptUser" />
              <LaunchEvent Type="OnAppointmentSend" FunctionName="onAppointmentSendHandler" SendMode="PromptUser" />
              <LaunchEvent Type="OnMessageCompose" FunctionName="onMessageComposeHandler" />
              <LaunchEvent Type="OnAppointmentOrganizer" FunctionName="onAppointmentOrganizerHandler" />
              -->
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

Outlook on Windows では JavaScript ファイルが使用されますが、Outlook on the webと新しい Mac UI では、同じ JavaScript ファイルを参照できる HTML ファイルが使用されます。 Outlook プラットフォームは最終的に Outlook クライアントに基づいて HTML と JavaScript のどちらを使用するかを決定するため、マニフェストのノードで `Resources` これらの両方のファイルへの参照を指定する必要があります。 そのため、イベント処理を構成するには、要素内の HTML の場所を **\<Runtime\>** 指定し、その `Override` 子要素で、HTML によってインライン化または参照される JavaScript ファイルの場所を指定します。

# <a name="teams-manifest-developer-preview"></a>[Teams マニフェスト (開発者プレビュー)](#tab/jsonmanifest)

1. **manifest.json** ファイルを開きます。

1. "extensions.runtimes" 配列に次のオブジェクトを追加します。 このマークアップについて、次の点に注意してください。

   - メールボックス要件セットの "minVersion" は "1.10" に設定されています。これは、この記事の前の表で、 と `OnNewAppointmentCompose` イベントをサポート`OnNewMessageCompose`する要件セットの最小バージョンであることが指定されているためです。
   - ランタイムの "id" は、わかりやすい名前 "autorun_runtime" に設定されます。
   - "code" プロパティには、HTML ファイルに設定された子 "page" プロパティと、JavaScript ファイルに設定された子 "script" プロパティがあります。 これらのファイルは、後の手順で作成または編集します。 Office では、プラットフォームに応じてこれらの値のいずれかを使用します。
       - Windows 上の Office では、JavaScript 専用ランタイムでイベント ハンドラーが実行され、JavaScript ファイルが直接読み込まれます。
       - Office on Mac と Web では、ブラウザー ランタイムでハンドラーが実行され、HTML ファイルが読み込まれます。 そのファイルには、JavaScript ファイルを `<script>` 読み込むタグが含まれています。
     詳細については、「 [Office アドインのランタイム](../testing/runtimes.md)」を参照してください。
   - "lifetime" プロパティは "short" に設定されています。これは、いずれかのイベントがトリガーされたときにランタイムが起動し、ハンドラーが完了するとシャットダウンすることを意味します。 (まれに、ハンドラーが完了する前にランタイムがシャットダウンする場合があります。 [「Office アドインのランタイム」を](../testing/runtimes.md)参照してください)。
   - ランタイムで実行できる "アクション" には 2 種類あります。 これらのアクションに対応する関数は、後の手順で作成します。

    ```json
     {
        "requirements": {
            "capabilities": [
                {
                    "name": "Mailbox",
                    "minVersion": "1.10"
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
                "id": "onNewMessageComposeHandler",
                "type": "executeFunction",
                "displayName": "onNewMessageComposeHandler"
            },
            {
                "id": "onNewAppointmentComposeHandler",
                "type": "executeFunction",
                "displayName": "onNewAppointmentComposeHandler"
            }
        ]
    }
    ```

1. "extensions" 配列の オブジェクトのプロパティとして、次の "autoRunEvents" 配列を追加します。

    ```json
    "autoRunEvents": [
    
    ]
    ```

1. "autoRunEvents" 配列に次のオブジェクトを追加します。 "events" プロパティは、この記事の前の表で説明したようにハンドラーをイベントにマップします。 ハンドラー名は、前の手順の "actions" 配列内のオブジェクトの "id" プロパティで使用されているものと一致する必要があります。

    ```json
      {
          "requirements": {
              "capabilities": [
                  {
                      "name": "Mailbox",
                      "minVersion": "1.10"
                  }
              ],
              "scopes": [
                  "mail"
              ]
          },
          "events": [
              {
                  "type": "newMessageComposeCreated",
                  "actionId": "onNewMessageComposeHandler"
              },
              {
                  "type": "newAppointmentOrganizerCreated",
                  "actionId": "onNewAppointmentComposeHandler"
              }
          ]
      }
    ```

---

> [!TIP]
>
> - アドインのランタイムの詳細については、「 [Office アドインのランタイム](../testing/runtimes.md)」を参照してください。
> - Outlook アドインのマニフェストの詳細については、「 [Outlook アドイン マニフェスト](manifests.md)」を参照してください。

## <a name="implement-event-handling"></a>イベント処理を実装する

選択したイベントの処理を実装する必要があります。

このシナリオでは、新しい項目を作成するための処理を追加します。

1. 同じクイック スタート プロジェクトから、**./src** ディレクトリの下に **launchevent** という名前の新しいフォルダーを作成します。

1. **./src/launchevent** フォルダーで、 という名前の新しいファイル **launchevent.js** 作成します。

1. コード エディターで **./src/launchevent/launchevent.js** ファイルを開き、次の JavaScript コードを追加します。

    ```js
    /*
    * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
    * See LICENSE in the project root for license information.
    */

    function onNewMessageComposeHandler(event) {
      setSubject(event);
    }
    function onNewAppointmentComposeHandler(event) {
      setSubject(event);
    }
    function setSubject(event) {
      Office.context.mailbox.item.subject.setAsync(
        "Set by an event-based add-in!",
        {
          "asyncContext": event
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

    // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
    Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
    Office.actions.associate("onNewAppointmentComposeHandler", onNewAppointmentComposeHandler);
    ```

1. 変更内容を保存します。

> [!IMPORTANT]
> Windows: 現時点では、イベント ベースのアクティブ化の処理を実装する JavaScript ファイルでは、インポートはサポートされていません。

## <a name="update-the-commands-html-file"></a>コマンド HTML ファイルを更新する

1. **./src/commands** フォルダーで、 **commands.html** を開きます。

1. 終了 **ヘッド** タグ (`</head>`) の直前に、イベント処理 JavaScript コードを含めるスクリプト エントリを追加します。

    ```html
    <script type="text/javascript" src="../launchevent/launchevent.js"></script>
    ```

1. 変更内容を保存します。

## <a name="update-webpack-config-settings"></a>Webpackの機能設定を更新する

1. プロジェクトのルート ディレクトリにある **webpack.config.js** ファイルを開き、次の手順を実行します。

1. オブジェクト内の `plugins` 配列を `config` 見つけて、配列の先頭にこの新しいオブジェクトを追加します。

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

1. Outlook on the web で新しいメッセージを作成します。

    ![compose に件名が設定されたOutlook on the webのメッセージ ウィンドウ。](../images/outlook-web-autolaunch-1.png)

1. 新しい Mac UI の Outlook で、新しいメッセージを作成します。

    ![新しい Mac UI 上の Outlook のメッセージ ウィンドウ。件名は compose に設定されています。](../images/outlook-mac-autolaunch.png)

1. Outlook on Windows で、新しいメッセージを作成します。

    ![Outlook on Windows のメッセージ ウィンドウ。件名が compose に設定されています。](../images/outlook-win-autolaunch.png)

## <a name="debug"></a>デバッグ

アドインで起動イベント処理に変更を加える際には、次の点に注意する必要があります。

- マニフェストを更新した場合は、 [アドインを削除](sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in)してから、もう一度サイドロードします。 Windows で Outlook を使用している場合は、Outlook を閉じてもう一度開きます。
- マニフェスト以外のファイルに変更を加えた場合は、Outlook on Windows を閉じてもう一度開くか、Outlook on the webを実行しているブラウザー タブを更新します。

独自の機能を実装する際に、コードのデバッグが必要になる場合があります。 イベント ベースのアドインのアクティブ化をデバッグする方法のガイダンスについては、「 [イベント ベースの Outlook アドインをデバッグ](debug-autolaunch.md)する」を参照してください。

ランタイム ログは、Windows のこの機能でも使用できます。 詳細については、「 [ランタイム ログを使用してアドインをデバッグ](../testing/runtime-logging.md#runtime-logging-on-windows)する」を参照してください。

[!INCLUDE [Loopback exemption note](../includes/outlook-loopback-exemption.md)]

## <a name="deploy-to-users"></a>ユーザーにデプロイする

Microsoft 365 管理センターを使用してマニフェストをアップロードすることで、イベント ベースのアドインをデプロイできます。 管理ポータルで、ナビゲーション ウィンドウの **[設定]** セクションを展開し、[ **統合アプリ**] を選択します。 [ **統合アプリ** ] ページで、[ **カスタム アプリのアップロード** ] アクションを選択します。

![[カスタム アプリのアップロード] アクションが強調表示されているMicrosoft 365 管理センターの [統合アプリ] ページ。](../images/outlook-deploy-event-based-add-ins.png)

> [!IMPORTANT]
> イベント ベースのアドインは、管理者が管理するデプロイのみに制限されます。 ユーザーは、AppSource またはアプリ内 Office ストアからイベント ベースのアドインをアクティブ化できません。 詳細については、 [イベント ベースの Outlook アドインの AppSource 登録情報オプションに関するページを参照してください](autolaunch-store-options.md)。

[!INCLUDE [outlook-smart-alerts-deployment](../includes/outlook-smart-alerts-deployment.md)]

## <a name="event-based-activation-behavior-and-limitations"></a>イベント ベースのアクティブ化の動作と制限事項

アドイン起動イベント ハンドラーは、実行時間が短く、軽量で、可能な限り非侵襲的である必要があります。 アクティブ化後、アドインは約 300 秒以内にタイムアウトします。これは、イベント ベースのアドインの実行に許可される最大時間です。アドインが起動イベントの処理を完了したことを通知するには、関連付けられているイベント ハンドラーが メソッドを `event.completed` 呼び出す必要があります。 (ステートメントの後に `event.completed` 含まれるコードは、実行が保証されていないことに注意してください)。アドインが処理するイベントがトリガーされるたびに、アドインが再アクティブ化され、関連付けられているイベント ハンドラーが実行され、タイムアウト ウィンドウがリセットされます。 アドインは、タイムアウト後に終了するか、ユーザーが作成ウィンドウを閉じるか、アイテムを送信します。

ユーザーが同じイベントをサブスクライブした複数のアドインがある場合、Outlook プラットフォームは特定の順序でアドインを起動します。 現時点では、アクティブに実行できるイベント ベースのアドインは 5 つだけです。

サポートされているすべての Outlook クライアントで、アドインの実行を完了するには、アドインがアクティブ化された現在のメール アイテムにユーザーを残す必要があります。 現在のアイテムから移動すると (たとえば、別の新規作成ウィンドウやタブに切り替えるなど)、アドイン操作が終了します。 また、アドインは、ユーザーが作成しているメッセージまたは予定を送信したときにも操作を停止します。

Windows クライアントでイベント ベースのアクティブ化の処理を実装する JavaScript ファイルでは、インポートはサポートされていません。

UI を変更または変更する一部のOffice.js API は、イベント ベースのアドインからは許可されません。ブロックされている API を次に示します。

- の下 `Office.context.auth`:
  - `getAccessToken`
  - `getAccessTokenAsync`
    > [!NOTE]
    > [OfficeRuntime.auth](/javascript/api/office-runtime/officeruntime.auth) は、イベント ベースのアクティブ化とシングル サインオン (SSO) をサポートするすべての Outlook バージョンでサポートされていますが、 [Office.auth](/javascript/api/office/office.auth) は特定の Outlook ビルドでのみサポートされています。 詳細については、「 [イベント ベースのアクティブ化を使用する Outlook アドインでシングル サインオン (SSO) を有効にする」を参照](use-sso-in-event-based-activation.md)してください。
- の下 `Office.context.mailbox`:
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- の下 `Office.context.mailbox.item`:
  - `close`
- の下 `Office.context.ui`:
  - `displayDialogAsync`
  - `messageParent`

### <a name="requesting-external-data"></a>外部データの要求

外部データを要求するには [、Fetch](https://developer.mozilla.org/docs/Web/API/Fetch_API) などの API を使用するか、サーバーと対話するための HTTP 要求を発行する標準 Web API である [XMLHttpRequest (XHR)](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest) を使用します。

XMLHttpRequest オブジェクトを使用する場合は、 [同じ配信元ポリシー](https://developer.mozilla.org/docs/Web/Security/Same-origin_policy) と単純な [CORS (クロスオリジン リソース共有)](https://developer.mozilla.org/docs/Web/HTTP/CORS) を必要とする追加のセキュリティ対策を使用する必要があることに注意してください。

[単純な CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS#simple_requests) 実装:

- Cookie を使用できません。
- 、、、 などの`GET``HEAD``POST`単純なメソッドのみをサポートします。
- フィールド名 `Accept`、、 `Accept-Language`または `Content-Language`を含む単純なヘッダーを受け入れます。
- コンテンツ タイプが `Content-Type`、、または `multipart/form-data`である場合は`application/x-www-form-urlencoded`、 を`text/plain`使用できます。
- によって `XMLHttpRequest.upload`返される オブジェクトにイベント リスナーを登録することはできません。
- 要求でオブジェクトを使用 `ReadableStream` することはできません。

> [!NOTE]
> 完全な CORS サポートは、Outlook on the web、Mac、および Windows で利用できます (バージョン 2201 以降、ビルド 16.0.14813.10000)。

## <a name="see-also"></a>関連項目

- [Outlook アドインのマニフェスト](manifests.md)
- [イベント ベースのアドインをデバッグする方法](debug-autolaunch.md)
- [イベント ベースの Outlook アドインの AppSource 登録情報オプション](autolaunch-store-options.md)
- [スマート アラートと OnMessageSend チュートリアル](smart-alerts-onmessagesend-walkthrough.md)
- Office アドインのコード サンプル:
  - [Outlook イベント ベースのアクティブ化を使用して添付ファイルを暗号化し、会議出席依頼の出席者を処理し、予定の日付/時刻の変更に対応する](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-encrypt-attachments)
  - [Outlook イベントベースのアクティブ化を使用して署名を設定する](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-set-signature)
  - [Outlook イベントベースのアクティブ化を使用して、外部受信者をタグ付けする](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-tag-external)
  - [Outlook スマート アラートを使用する](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-check-item-categories)
