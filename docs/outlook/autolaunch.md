---
title: イベント ベースのOutlookアドインを構成する
description: イベント ベースのアクティブ化Outlookアドインを構成する方法について学習します。
ms.topic: article
ms.date: 03/09/2022
ms.localizationpriority: medium
ms.openlocfilehash: bd6dfab38d59d5e120ca9672df8eb3ac6e7654e7
ms.sourcegitcommit: 287a58de82a09deeef794c2aa4f32280efbbe54a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/28/2022
ms.locfileid: "64496909"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation"></a>イベント ベースのOutlookアドインを構成する

イベント ベースのアクティブ化機能がない場合、ユーザーはタスクを完了するためにアドインを明示的に起動する必要があります。 この機能を使用すると、特定のイベントに基づいて、特にすべてのアイテムに適用される操作に基づいてタスクを実行できます。 作業ウィンドウや UI レス機能と統合することもできます。

このチュートリアルの最後には、新しいアイテムが作成され、件名が設定されるたびに実行されるアドインがあります。

> [!NOTE]
> この機能のサポートは、要件セット [1.10 で導入されました](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10)。 この要件セットをサポートする [クライアントおよびプラットフォーム](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients) を参照してください。

## <a name="supported-events"></a>サポートされるイベント

次の表に、現在利用可能なイベントと、各イベントでサポートされているクライアントの一覧を示します。 イベントが発生すると、ハンドラーはイベントの種類に `event` 固有の詳細を含む可能性があるオブジェクトを受け取ります。 [ **説明]** 列には、該当する関連オブジェクトへのリンクが含まれます。

> [!IMPORTANT]
> プレビュー中のイベントは、Microsoft 365 サブスクリプションと、次の表に示す限られた一連のサポートされているクライアントでのみ使用できます。 クライアント構成の詳細については、「 [この記事でプレビューする方法」](#how-to-preview) を参照してください。 プレビュー イベントは、実稼働アドインでは使用できません。

|イベント|説明|最小要件セットとサポートされているクライアント|
|---|---|---|
|`OnNewMessageCompose`|新しいメッセージを作成する場合 (返信、すべて返信、転送を含む) が、下書きなど編集時には作成されません。|[1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10)<br><br>- Windows<br>- Web ブラウザー<br>- 新しい Mac UI プレビュー|
|`OnNewAppointmentOrganizer`|既存の予定の編集ではなく、新しい予定を作成する場合。|[1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10)<br><br>- Windows<br>- Web ブラウザー<br>- 新しい Mac UI プレビュー|
|`OnMessageAttachmentsChanged`|メッセージの作成中に添付ファイルを追加または削除する場合。<br><br>イベント固有のデータ オブジェクト: [AttachmentsChangedEventArgs](/javascript/api/outlook/office.attachmentschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<br>- Web ブラウザー|
|`OnAppointmentAttachmentsChanged`|予定の作成中に添付ファイルを追加または削除する場合。<br><br>イベント固有のデータ オブジェクト: [AttachmentsChangedEventArgs](/javascript/api/outlook/office.attachmentschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<br>- Web ブラウザー|
|`OnMessageRecipientsChanged`|メッセージの作成中に受信者を追加または削除する場合。<br><br>イベント固有のデータ オブジェクト: [RecipientsChangedEventArgs](/javascript/api/outlook/office.recipientschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<br>- Web ブラウザー|
|`OnAppointmentAttendeesChanged`|予定の作成中に出席者を追加または削除する場合。<br><br>イベント固有のデータ オブジェクト: [RecipientsChangedEventArgs](/javascript/api/outlook/office.recipientschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<br>- Web ブラウザー|
|`OnAppointmentTimeChanged`|予定の作成中に日付/時刻を変更する場合。<br><br>イベント固有のデータ オブジェクト: [AppointmentTimeChangedEventArgs](/javascript/api/outlook/office.appointmenttimechangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<br>- Web ブラウザー|
|`OnAppointmentRecurrenceChanged`|予定の作成中に定期的な詳細を追加、変更、または削除する場合。 日付/時刻が変更された場合、イベント `OnAppointmentTimeChanged` も発生します。<br><br>イベント固有のデータ オブジェクト: [RecurrenceChangedEventArgs](/javascript/api/outlook/office.recurrencechangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<br>- Web ブラウザー|
|`OnInfoBarDismissClicked`|メッセージまたは予定アイテムの作成中に通知を却下する場合。 通知を追加したアドインだけが通知されます。<br><br>イベント固有のデータ オブジェクト: [InfobarClickedEventArgs](/javascript/api/outlook/office.infobarclickedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>- Windows<br>- Web ブラウザー|
|`OnMessageSend`|メッセージ アイテムの送信時。 詳細については、「スマート アラート」 [のチュートリアルを参照してください](smart-alerts-onmessagesend-walkthrough.md)。|[プレビュー](/javascript/api/requirement-sets/outlook/preview-requirement-set/outlook-requirement-set-preview)<br><br>- Windows|
|`OnAppointmentSend`|予定アイテムの送信時。 詳細については、「スマート アラート」 [のチュートリアルを参照してください](smart-alerts-onmessagesend-walkthrough.md)。|[プレビュー](/javascript/api/requirement-sets/outlook/preview-requirement-set/outlook-requirement-set-preview)<br><br>- Windows|

### <a name="how-to-preview"></a>プレビューする方法

プレビューで今すぐイベントを試してみてください。 このページの最後にある「フィードバック」セクションをGitHubフィードバックを提供することで、お客様のシナリオと改善方法をお知らせします。

使用可能な場合にこれらのイベントをプレビューするには、次の手順を実行します。

- 次のOutlook on the web:
  - [ターゲット リリースをテナントにMicrosoft 365します。](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)
  - () **の** ベータ ライブラリを参照CDN。https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) TypeScript コンパイルおよび IntelliSense の [型定義ファイル](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)は CDN で見つかり、[DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts) にあります。 これらの型は、`npm install --save-dev @types/office-js-preview` を使用してインストールできます。
- 新Outlook Mac UI プレビューの詳細については、次の操作を行います。
  - 最小必要なビルドは 16.54 (21101001)。 Insider プログラムOffice[参加](https://insider.office.com/join/Mac)し、ベータ版のビルドにアクセスするためのベータ Officeを選択します。
- [OutlookのWindows:
  - 必要な最小ビルドは 16.0.14511.10000 です。 Insider プログラムOffice[参加](https://insider.office.com/join/windows)し、ベータ版のビルドにアクセスするためのベータ Officeを選択します。

## <a name="set-up-your-environment"></a>環境を設定する

クイック スタート[Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator)完了し、Yeoman ジェネレーターを使用してアドイン プロジェクトを作成し、Office作成します。

## <a name="configure-the-manifest"></a>マニフェストを構成する

アドインのイベント ベースのアクティブ化を有効にするには、マニフェストのノードで [Runtimes](/javascript/api/manifest/runtimes) 要素と [LaunchEvent](/javascript/api/manifest/extensionpoint#launchevent) `VersionOverridesV1_1` 拡張ポイントを構成する必要があります。 今のところ、 `DesktopFormFactor` サポートされている唯一のフォーム ファクターです。

1. コード エディターで、クイック スタート プロジェクトを開きます。

1. プロジェクトの **manifest.xml** にあるファイルを開きます。

1. ノード全体 (開く `<VersionOverrides>` タグと閉じるタグを含む) を選択し、次の XML に置き換え、変更を保存します。

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

          <!-- Enable launching the add-in on the included events. -->
          <ExtensionPoint xsi:type="LaunchEvent">
            <LaunchEvents>
              <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
              <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onAppointmentComposeHandler"/>
              
              <!-- Other available events (currently released) -->
              <!--
              <LaunchEvent Type="OnMessageAttachmentsChanged" FunctionName="onMessageAttachmentsChangedHandler" />
              <LaunchEvent Type="OnAppointmentAttachmentsChanged" FunctionName="onAppointmentAttachmentsChangedHandler" />
              <LaunchEvent Type="OnMessageRecipientsChanged" FunctionName="onMessageRecipientsChangedHandler" />
              <LaunchEvent Type="OnAppointmentAttendeesChanged" FunctionName="onAppointmentAttendeesChangedHandler" />
              <LaunchEvent Type="OnAppointmentTimeChanged" FunctionName="onAppointmentTimeChangedHandler" />
              <LaunchEvent Type="OnAppointmentRecurrenceChanged" FunctionName="onAppointmentRecurrenceChangedHandler" />
              <LaunchEvent Type="OnInfoBarDismissClicked" FunctionName="onInfobarDismissClickedHandler" />
              -->

              <!-- Other available events (currently in preview) -->
              <!--
              <LaunchEvent Type="OnMessageSend" FunctionName="onMessageSendHandler" SendMode="PromptUser" />
              <LaunchEvent Type="OnAppointmentSend" FunctionName="onAppointmentSendHandler" SendMode="PromptUser" />
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

OutlookはWindows JavaScript ファイルを使用しますが、Outlook on the web および新しい Mac UI プレビューでは、同じ JavaScript ファイルを参照できる HTML ファイルを使用します。 `Resources` Outlook プラットフォームは最終的に、Outlook クライアントに基づいて HTML または JavaScript を使用するかどうかを決定するために、マニフェストのノードでこれらの両方のファイルへの参照を提供する必要があります。 そのため、イベント処理を構成するには、要素内の HTML `Runtime` `Override` の場所を指定し、その子要素で HTML によってインライン化または参照される JavaScript ファイルの場所を指定します。

> [!TIP]
> アドインのマニフェストのOutlook詳細については、「Outlook[マニフェスト」を参照してください](manifests.md)。

## <a name="implement-event-handling"></a>イベント処理の実装

選択したイベントの処理を実装する必要があります。

このシナリオでは、新しいアイテムを作成する処理を追加します。

1. 同じクイック スタート プロジェクトから、./src ディレクトリの **下に launchevent という** 名前の **新しいフォルダーを作成** します。

1. **./src/launchevent フォルダーで**、次の名前の新しいファイルを **launchevent.js**。

1. コード エディターで **ファイル ./src/launchevent/launchevent.js** を開き、次の JavaScript コードを追加します。

    ```js
    /*
    * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
    * See LICENSE in the project root for license information.
    */

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
    Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
    Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
    ```

1. 変更内容を保存します。

> [!IMPORTANT]
> Windows: 現在、イベント ベースのアクティブ化の処理を実装する JavaScript ファイルではインポートはサポートされていません。

## <a name="update-the-commands-html-file"></a>コマンドの HTML ファイルを更新する

1. **./src/commands フォルダーで**、ファイルを **開** commands.html。

1. 終了ヘッド タグ (`<\head>`**) の直前** に、イベント処理 JavaScript コードを含めるスクリプト エントリを追加します。

    ```html
    <script type="text/javascript" src="../launchevent/launchevent.js"></script>
    ```

1. 変更内容を保存します。

## <a name="update-webpack-config-settings"></a>Webpackの機能設定を更新する

1. プロジェクトの **ルートwebpack.config.js** にあるファイルを開き、次の手順を実行します。

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

1. Outlook on the web で新しいメッセージを作成します。

    ![作成時に件名が設定されているOutlook on the webウィンドウのスクリーンショット。](../images/outlook-web-autolaunch-1.png)

1. 新Outlook Mac UI プレビューで、新しいメッセージを作成します。

    ![新しい Mac UI プレビュー Outlookメッセージ ウィンドウのスクリーンショットを作成時に設定します。](../images/outlook-mac-autolaunch.png)

1. [Outlook] でWindows新しいメッセージを作成します。

    ![作成時に設定された件名OutlookのWindowsウィンドウのスクリーンショット。](../images/outlook-win-autolaunch.png)

## <a name="debug"></a>デバッグ

アドインで起動イベント処理に変更を加える場合は、次の点に注意する必要があります。

- マニフェストを更新した場合は、 [アドインを](sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in)削除してから、もう一度サイドロードします。 アプリでアプリを使用しているOutlook、Windows閉じて再度開きます。
- マニフェスト以外のファイルに変更を加えた場合は、OutlookでWindowsを閉じて再び開くか、Outlook on the web を実行しているブラウザー タブを更新します。

独自の機能を実装する場合は、コードのデバッグが必要な場合があります。 イベント ベースのアドインのアクティブ化をデバッグする方法のガイダンスについては、「イベント ベースのアドインをデバッグするOutlook[参照してください](debug-autolaunch.md)。

ランタイム ログは、この機能で使用することもできます。Windows。 詳細については、「ランタイム ログを [使用してアドインをデバッグする」を参照してください](../testing/runtime-logging.md#runtime-logging-on-windows)。

[!INCLUDE [Loopback exemption note](../includes/outlook-loopback-exemption.md)]

## <a name="deploy-to-users"></a>ユーザーへの展開

イベント ベースのアドインを展開するには、イベント ベースのアドインを使用してマニフェストをアップロードMicrosoft 365 管理センター。 管理ポータルで、ナビゲーション ウィンドウで [**設定] セクション** を展開し、[統合アプリ **] を選択します**。 [統合アプリ **] ページで**、[カスタム アプリ] アップロード **を選択** します。

![カスタム アプリのアクションを含む、Microsoft 365 管理センターアプリの統合アップロードのスクリーンショット。](../images/outlook-deploy-event-based-add-ins.png)

AppSource とアプリ内 Office ストア: イベント ベースのアドインを展開したり、イベント ベースのアクティブ化機能を含める既存のアドインを更新したりする機能は、すぐに利用できる必要があります。

> [!IMPORTANT]
> イベント ベースのアドインは、管理者が管理する展開にのみ制限されます。 今のところ、ユーザーは AppSource またはアプリ内アドインからイベント ベースのアドインを取得Officeできます。 詳細については、「AppSource リスト オプション」を参照して、イベント ベースのOutlook[を参照してください](autolaunch-store-options.md)。

## <a name="event-based-activation-behavior-and-limitations"></a>イベント ベースのアクティブ化の動作と制限

アドイン起動イベント ハンドラーは、実行時間が短く、軽量で、可能な限り非インバシブである必要があります。 アクティブ化後、アドインはイベント ベースのアドインを実行できる最大時間である約 300 秒以内にタイム アウトします。アドインが起動イベントの処理を完了したというメッセージを表示するには、関連付けられたハンドラーにメソッドを呼び出 `event.completed` す必要があります。 (ステートメントの後に含まれる `event.completed` コードは、実行が保証されない点に注意してください)。アドインが処理するイベントがトリガーされるごとに、アドインが再アクティブ化され、関連付けられたイベント ハンドラーが実行され、タイムアウト ウィンドウがリセットされます。 アドインは、タイム アウト後に終了するか、ユーザーが作成ウィンドウを閉じるか、アイテムを送信します。

ユーザーが同じイベントにサブスクライブしている複数のアドインがある場合、Outlook プラットフォームは特定の順序でアドインを起動します。 現在、アクティブに実行できるイベント ベースのアドインは 5 つのみです。

ユーザーは、アドインの実行を開始した現在のメール アイテムから切り替えまたは移動できます。 起動されたアドインは、バックグラウンドで操作を終了します。

JavaScript ファイルでは、イベント ベースのアクティブ化の処理をクライアントで実装する場合、インポートはWindowsされません。

UI Office.js変更する API の一部は、イベント ベースのアドインでは使用できない場合があります。ブロックされている API を次に示します。

- [ : ] の下`Office.context.auth`
  - `getAccessToken`
  - `getAccessTokenAsync`
    > [!NOTE]
    > `OfficeRuntime.auth` がサポートされています。 詳細については、「イベント ベースのライセンス認証を使用する Outlookでシングル サインオン [(SSO) を有効にする」を参照してください](use-sso-in-event-based-activation.md)。
- [ : ] の下`Office.context.mailbox`
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- [ : ] の下`Office.context.mailbox.item`
  - `close`
- [ : ] の下`Office.context.ui`
  - `displayDialogAsync`
  - `messageParent`

### <a name="requesting-external-data"></a>外部データの要求

外部データを要求するには、 [Fetch](https://developer.mozilla.org/docs/Web/API/Fetch_API) のような API を使用するか、サーバーとやり取りするための HTTP 要求を発行する標準 Web API [である XmlHttpRequest (XHR)](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest) を使用します。

XmlHttpRequests を作成する場合は、追加のセキュリティ対策を使用する必要があります[](https://developer.mozilla.org/docs/Web/Security/Same-origin_policy)。同じオリジン ポリシーと単純な [CORS が必要です](https://www.w3.org/TR/cors/)。

単純な CORS 実装では Cookie を使用できません。単純なメソッド (GET、HEAD、POST) のみをサポートします。 単純な CORS はフィールド名`Accept`、 `Accept-Language`、`Content-Language`の簡単なヘッダーを受け入れます。 コンテンツ タイプが 、 `Content-Type` である場合は、単純な CORS `application/x-www-form-urlencoded`でヘッダーを `text/plain`使用できます `multipart/form-data`。

CORS の完全なサポートは近日公開予定です。

## <a name="see-also"></a>関連項目

- [Outlook アドインのマニフェスト](manifests.md)
- [イベント ベースのアドインをデバッグする方法](debug-autolaunch.md)
- [イベント ベースのアドインの AppSource Outlookオプション](autolaunch-store-options.md)
- [スマート アラートと OnMessageSend のチュートリアル](smart-alerts-onmessagesend-walkthrough.md)
- PnP サンプル:
  - [イベント ベースOutlookを使用して、添付ファイルの暗号化、会議出席依頼の出席者の処理、予定の日時の変更への対応を行います。](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-encrypt-attachments)
  - [Outlook イベントベースのアクティブ化を使用して署名を設定する](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-set-signature)
  - [Outlook イベントベースのアクティブ化を使用して、外部受信者をタグ付けする](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-tag-external)
