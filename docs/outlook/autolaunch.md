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
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a>イベント ベースのアクティブ化 (プレビュー) 用にOutlook アドインを構成する

イベントベースのアクティブ化機能を使用しない場合、ユーザーは、アドインを明示的に起動してタスクを完了する必要があります。 この機能により、特定のイベントに基づいて、特にすべての項目に適用される操作に基づいて、アドインでタスクを実行できます。 作業ウィンドウと UI を使用する機能を統合することもできます。

このチュートリアルの最後に、新しい項目が作成され、件名を設定するたびに実行されるアドインが用意されています。

> [!IMPORTANT]
> この機能は、web 上のOutlookで[プレビュー](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)し、Microsoft 365サブスクリプションでWindowsにのみサポートされます。 詳細については、この記事 [の「イベントベースのアクティブ化機能をプレビューする方法](#how-to-preview-the-event-based-activation-feature) 」を参照してください。
>
> プレビュー機能は予告なく変更される場合があるため、運用アドインで使用しないでください。

## <a name="supported-events"></a>サポートされるイベント

現在、以下のイベントがサポートされています。

|イベント|説明|クライアント|
|---|---|---|
|`OnNewMessageCompose`|新しいメッセージ (返信、全員への返信、転送を含む) の作成時に、編集時 (下書きなど) は作成しません。|Windows, ウェブ|
|`OnNewAppointmentOrganizer`|新しい予定を作成するが、既存の予定を編集する場合。|Windows, ウェブ|
|`OnMessageAttachmentsChanged`|メッセージの作成中に添付ファイルを追加または削除する。|Windows|
|`OnAppointmentAttachmentsChanged`|予定の作成中に添付ファイルを追加または削除する。|Windows|
|`OnMessageRecipientsChanged`|メッセージの作成中に受信者を追加または削除する。|Windows|
|`OnAppointmentAttendeesChanged`|予定の作成中に出席者を追加または削除する。|Windows|
|`OnAppointmentTimeChanged`|予定の作成中に日付/時刻を変更する場合。|Windows|
|`OnAppointmentRecurrenceChanged`|予定の作成中に定期的なアイテムの詳細を追加、変更、または削除する。 日付/時刻が変更されると、 `OnAppointmentTimeChanged` イベントも発生します。|Windows|
|`OnInfoBarDismissClicked`|メッセージまたは予定アイテムの作成中に通知を閉じる。 通知を追加したアドインのみが通知されます。|Windows|

## <a name="how-to-preview-the-event-based-activation-feature"></a>イベントベースのアクティブ化機能をプレビューする方法

イベントベースのアクティベーション機能を試してみてください! GitHubを通じてフィードバックを提供することで、お客様のシナリオと改善方法をお知らせください(このページの最後にある **フィードバック** セクションを参照)。

この機能をプレビューするには:

- ウェブ上のOutlookの場合:
  - [Microsoft 365 テナントで対象リリースを構成する](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center):
  - CDN ( ) の **ベータ ライブラリ** を参照 https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) します。 TypeScript コンパイルとIntelliSenseの[型定義ファイル](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)は、CDNと[Typed にあります](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts)。 これらの種類は、 を使用してインストールできます `npm install --save-dev @types/office-js-preview` 。
- WindowsのOutlookの場合:
  - 必要最小限のビルドは 16.0.14026.20000 です。 ベータ版ビルドにアクセスするには[、Office Insider プログラム](https://insider.office.com)Office参加してください。
  - レジストリを構成します。 Outlookには、CDNから読み込む代わりに、Office.jsの実稼働およびベータ版のローカル コピーが含まれます。 デフォルトでは、API のローカル本番コピーが参照されます。 Outlookの JavaScript API のローカル ベータ 版に切り替えるには、このレジストリ エントリを追加する必要があります。
    1. レジストリ キーを作成 `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\WebExt\Developer` する:
    1. という名前のエントリを追加 `EnableBetaAPIsInJavaScript` し、値を `1` に設定します。 レジストリは次の図のようになります。

        ![レジストリ キー値を持つレジストリ エディターのスクリーンショット](../images/outlook-beta-registry-key.png)

## <a name="set-up-your-environment"></a>環境を設定する

アドイン[Outlookクイック スタート](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator)を完了し、アドインの Office用の Yeoman ジェネレーターを使用してアドイン プロジェクトを作成します。

## <a name="configure-the-manifest"></a>マニフェストを構成する

アドインのイベント ベースのアクティブ化を有効にするには、マニフェストのノードで [Runtimes](../reference/manifest/runtimes.md) 要素と [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) 拡張ポイントを構成する必要があります `VersionOverridesV1_1` 。 現時点では、 `DesktopFormFactor` サポートされているフォーム ファクターは唯一です。

1. コード エディターで、クイック スタート プロジェクトを開きます。

1. プロジェクトのルートにある **manifest.xml** ファイルを開きます。

1. ノード全体 `<VersionOverrides>` (開くタグと閉じるタグを含む) を選択し、次の XML で置き換えて、変更を保存します。

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

Windows上のOutlookは JavaScript ファイルを使用しますが、web 上のOutlookは同じ JavaScript ファイルを参照できる HTML ファイルを使用します。 Outlook プラットフォームが最終的に `Resources` Outlook クライアントに基づいて HTML または JavaScript を使用するかどうかを決定するので、マニフェストのノードでこれらのファイルの両方への参照を指定する必要があります。 そのため、イベント処理を構成するには、要素内の HTML の場所を指定 `Runtime` し、その `Override` 子要素で、インライン化または HTML によって参照される JavaScript ファイルの場所を提供します。

> [!TIP]
> Outlook アドインのマニフェストの詳細については、「アドイン マニフェスト[のOutlook」](manifests.md)を参照してください。

## <a name="implement-event-handling"></a>イベント処理の実装

選択したイベントの処理を実装する必要があります。

このシナリオでは、新しいアイテムを作成するための処理を追加します。

1. 同じクイック スタート プロジェクトから、コード エディターでファイル **./src/commands/commands.js** を開きます。

1. 関数の後 `action` に、次の JavaScript 関数を挿入します。

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

1. ファイルの末尾に次の JavaScript コードを追加します。

    ```js
    // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
    Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
    Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
    ```

1. 変更内容を保存します。

## <a name="try-it-out"></a>試してみる

1. プロジェクトのルート ディレクトリから次のコマンドを実行します。 このコマンドを実行すると、ローカル Web サーバーが (まだ実行されていない場合) 起動し、アドインがサイドロードされます。

    ```command&nbsp;line
    npm start
    ```

    > [!NOTE]
    > アドインが自動的にサイドロードされなかった場合は、「[サイドロード Outlook アドインをテスト用に](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually)実行する」の指示に従って、アドインを手動でOutlookサイドロードします。

1. Outlook on the web で新しいメッセージを作成します。

    ![Web 上のOutlookのメッセージ ウィンドウのスクリーンショット (作成時に設定された件名)](../images/outlook-web-autolaunch-1.png)

1. Windows Outlookで、新しいメッセージを作成します。

    ![作成時に設定された件名を持つWindowsのOutlookのメッセージ ウィンドウのスクリーンショット](../images/outlook-win-autolaunch.png)

    > [!NOTE]
    > localhost からアドインを実行していて、"申し訳ありませんが *、{アドイン名- here}* にアクセスできませんでした。 ネットワーク接続があることを確認します。 問題が解決しない場合は、後で再試行してください。
    >
    > 1. Outlook を終了します。
    > 1. タスク **マネージャ** を開き **、msoadfsb.exe** プロセスが実行されていないことを確認します。
    > 1. 次のコマンドを実行します。
    >
    >    ```command&nbsp;line
    >    call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_http___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
    >    ```
    >
    > 1. Outlook を再起動します。

## <a name="debug"></a>Debug

アドインで起動イベント処理を変更する場合は、次の点に注意する必要があります。

- マニフェストを更新した場合は、 [アドインを削除](sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in) してからサイドロードし直します。
- マニフェスト以外のファイルに変更を加えた場合は、WindowsでOutlookを閉じて再度開くか、web 上でOutlookを実行しているブラウザー タブを更新します。

独自の機能を実装する場合は、コードのデバッグが必要になる場合があります。 イベント ベースのアドインのアクティブ化をデバッグする方法については、「[イベント ベースのアドインOutlookデバッグ](debug-autolaunch.md)する 」を参照してください。

ランタイム ログは、Windowsでもこの機能で使用できます。 詳細については、「 [ランタイム ログを使用したアドインのデバッグ](../testing/runtime-logging.md#runtime-logging-on-windows)」を参照してください。

## <a name="deploy-to-users"></a>ユーザーへの展開

Microsoft 365管理センターからマニフェストをアップロードすることで、イベントベースのアドインを展開できます。 管理ポータルで、ナビゲーション ウィンドウの **設定** セクションを展開し、[**統合アプリ**] を選択します。 [**統合アプリ**] ページで、[**カスタム アプリのアップロード]** アクションを選択します。

![Microsoft 365管理センターの [統合アプリ] ページのスクリーンショット (アップロードカスタム アプリ アクションなど)](../images/outlook-deploy-event-based-add-ins.png)

AppSource ストアとインクライアント ストア: イベント ベースのアドインを展開したり、既存のアドインを更新してイベント ベースのアクティブ化機能を含める機能は、すぐに利用できるようになります。

> [!IMPORTANT]
> イベントベースのアドインは、管理者が管理する展開のみに制限されます。 現時点では、ユーザーは AppSource ストアまたはインクライアント ストアからイベントベースのアドインを取得できません。

## <a name="event-based-activation-behavior-and-limitations"></a>イベントベースのアクティブ化の動作と制限

アドインの起動イベント ハンドラーは、短時間で軽量で、できるだけ非侵襲的であることが予想されます。 アクティブ化後、アドインは約 300 秒以内にタイムアウトし、イベント ベースのアドインの実行に許容される最大時間が長くなります。アドインが起動イベントの処理を完了したことを知らせるために、関連付けられたハンドラーがメソッドを呼び出すことをお勧 `event.completed` めします。 (ステートメントの後に含まれるコード `event.completed` は、実行が保証されないことに注意してください。アドインが処理するイベントがトリガーされるたびに、アドインが再アクティブ化され、関連付けられたイベント ハンドラーが実行され、タイムアウト ウィンドウがリセットされます。 アドインはタイムアウト後に終了するか、ユーザーが作成ウィンドウを閉じるか、アイテムを送信します。

ユーザーが同じイベントにサブスクライブした複数のアドインを持っている場合、Outlook プラットフォームはアドインを順不同で起動します。 現在、アクティブに実行できるイベントベースのアドインは 5 つだけです。

ユーザーは、アドインの実行を開始した現在のメール アイテムを切り替えたり、移動したりできます。 起動されたアドインは、バックグラウンドで操作を終了します。

UI を変更または変更する一部のOffice.js API は、イベント ベースのアドインからは許可されません。ブロックされた API は次のとおりです。

- 以下 `OfficeRuntime.auth` :
  - `getAccessToken`(Windowsのみ)
- 以下 `Office.context.auth` :
  - `getAccessToken`
  - `getAccessTokenAsync`
- 以下 `Office.context.mailbox` :
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- 以下 `Office.context.mailbox.item` :
  - `close`
- 以下 `Office.context.ui` :
  - `displayDialogAsync`
  - `messageParent`

## <a name="see-also"></a>関連項目

- [Outlook アドインのマニフェスト](manifests.md)
- [イベント ベースのアドインをデバッグする方法](debug-autolaunch.md)
