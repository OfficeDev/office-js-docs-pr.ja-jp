---
title: イベント ベースOutlook用にアドインを構成する (プレビュー)
description: イベント ベースのアクティブ化Outlookアドインを構成する方法について学習します。
ms.topic: article
ms.date: 05/04/2021
localization_priority: Normal
ms.openlocfilehash: 0052f08e9c6a3903f4adb48efca3ff29a6d21467
ms.sourcegitcommit: 8fbc7c7eb47875bf022e402b13858695a8536ec5
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/06/2021
ms.locfileid: "52253324"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a>イベント ベースOutlook用にアドインを構成する (プレビュー)

イベント ベースのアクティブ化機能がない場合、ユーザーはタスクを完了するためにアドインを明示的に起動する必要があります。 この機能を使用すると、特定のイベントに基づいて、特にすべてのアイテムに適用される操作に基づいてタスクを実行できます。 作業ウィンドウや UI レス機能と統合することもできます。

このチュートリアルの最後には、新しいアイテムが作成され、件名が設定されるたびに実行されるアドインがあります。

> [!IMPORTANT]
> この機能は、Web 上[](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)および Outlookサブスクリプションでのプレビュー WindowsでのみMicrosoft 365されます。 詳細については、この記事の [「イベント ベースのアクティブ化機能](#how-to-preview-the-event-based-activation-feature) をプレビューする方法」を参照してください。
>
> プレビュー機能は予告なしに変更される可能性があるため、実稼働アドインでは使用できません。

## <a name="supported-events"></a>サポートされるイベント

現時点では、次のイベントがサポートされています。

|イベント|説明|クライアント|
|---|---|---|
|`OnNewMessageCompose`|新しいメッセージを作成する場合 (返信、すべて返信、転送を含む) が、下書きなど編集時には作成されません。|Windows Web|
|`OnNewAppointmentOrganizer`|既存の予定の編集ではなく、新しい予定を作成する場合。|Windows Web|
|`OnMessageAttachmentsChanged`|メッセージの作成中に添付ファイルを追加または削除する場合。|Windows|
|`OnAppointmentAttachmentsChanged`|予定の作成中に添付ファイルを追加または削除する場合。|Windows|
|`OnMessageRecipientsChanged`|メッセージの作成中に受信者を追加または削除する場合。|Windows|
|`OnAppointmentAttendeesChanged`|予定の作成中に出席者を追加または削除する場合。|Windows|
|`OnAppointmentTimeChanged`|予定の作成中に日付/時刻を変更する場合。|Windows|
|`OnAppointmentRecurrenceChanged`|予定の作成中に定期的な詳細を追加、変更、または削除する場合。 日付/時刻が変更された場合、 `OnAppointmentTimeChanged` イベントも発生します。|Windows|
|`OnInfoBarDismissClicked`|メッセージまたは予定アイテムの作成中に通知を却下する場合。 通知を追加したアドインだけが通知されます。|Windows|

## <a name="how-to-preview-the-event-based-activation-feature"></a>イベント ベースのアクティブ化機能をプレビューする方法

イベント ベースのアクティブ化機能を試してみてください。 このページの最後にある「フィードバック」セクションをGitHubフィードバックを提供することで、お客様のシナリオと改善方法をお知らせします。

この機能をプレビューするには、次の方法を使用します。

- Web Outlookの詳細については、次の情報を参照してください。
  - [ターゲット リリースをテナントにMicrosoft 365します](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)。
  - ()**の** ベータ ライブラリを参照 https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) CDN。 TypeScript[のコンパイルと](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)定義の種類定義ファイルは、IntelliSenseと[DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts)にあるCDNです。 これらの種類は、 を使用してインストールできます `npm install --save-dev @types/office-js-preview` 。
- [OutlookのWindows:
  - 必要な最小ビルドは 16.0.14026.20000 です。 ベータ版[ビルドOfficeアクセスするには、Insider](https://insider.office.com)プログラムOffice参加します。
  - レジストリを構成します。
    1. レジストリ キーを作成します `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\WebExt\Developer` 。
    1. という名前のエントリを `EnableBetaAPIsInJavaScript` 追加し、値をに設定します `1` 。

## <a name="set-up-your-environment"></a>環境を設定する

クイック スタート[Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator)完了し、Yeoman ジェネレーターを使用してアドイン プロジェクトを作成し、Office作成します。

## <a name="configure-the-manifest"></a>マニフェストを構成する

アドインのイベント ベースのアクティブ化を有効にするには、マニフェストのノードで [Runtimes](../reference/manifest/runtimes.md) 要素と [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) 拡張ポイント `VersionOverridesV1_1` を構成する必要があります。 今のところ、 `DesktopFormFactor` サポートされている唯一のフォーム ファクターです。

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

          <!-- Enable launching the add-in on the included events. -->
          <ExtensionPoint xsi:type="LaunchEvent">
            <LaunchEvents>
              <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
              <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onAppointmentComposeHandler"/>
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

OutlookはWindows JavaScript ファイルを使用しますが、web 上Outlookは同じ JavaScript ファイルを参照できる HTML ファイルを使用します。 Outlook プラットフォームは最終的に、Outlook クライアントに基づいて HTML または JavaScript を使用するかどうかを決定するために、マニフェストのノードでこれらの両方のファイル `Resources` への参照を提供する必要があります。 そのため、イベント処理を構成するには、要素内の HTML の場所を指定し、その子要素で HTML によってインライン化または参照される JavaScript ファイルの場所 `Runtime` `Override` を指定します。

> [!TIP]
> アドインのマニフェストのOutlook詳細については、「Outlook[マニフェスト」を参照してください](manifests.md)。

## <a name="implement-event-handling"></a>イベント処理の実装

選択したイベントの処理を実装する必要があります。

このシナリオでは、新しいアイテムを作成する処理を追加します。

1. 同じクイック スタート プロジェクトで、コード エディター **で ./src/commands/commands.js** ファイルを開きます。

1. 関数の `action` 後に、次の JavaScript 関数を挿入します。

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

1. Office アドインの **Yeoman** ジェネレーターによって生成されたこのプロジェクトを使用して、web 上の Outlook で機能する関数については、ファイルの末尾に次のステートメントを追加します。

    ```js
    g.onMessageComposeHandler = onMessageComposeHandler;
    g.onAppointmentComposeHandler = onAppointmentComposeHandler;
    ```

1. 関数がファイル上でOutlook **動作** Windows、ファイルの末尾に次の JavaScript コードを追加します。

    ```js
    if (Office.actions) {
      // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
      Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
      Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
    }
    ```

    **注**: チェックを `Office.actions` 実行すると、web 上Outlookこれらのステートメントが無視されます。

1. 変更内容を保存します。

## <a name="try-it-out"></a>試してみる

1. プロジェクトのルート ディレクトリから次のコマンドを実行します。 このコマンドを実行すると、ローカル Web サーバーが (まだ実行されていない場合) 起動し、アドインがサイドロードされます。

    ```command&nbsp;line
    npm start
    ```

1. Outlook on the web で新しいメッセージを作成します。

    ![作成時に件名が設定Outlook Web 上のメッセージ ウィンドウのスクリーンショット](../images/outlook-web-autolaunch-1.png)

1. [Outlook] でWindows新しいメッセージを作成します。

    ![作成時に件名が設定されているOutlookのWindowsウィンドウのスクリーンショット](../images/outlook-win-autolaunch.png)

    > [!NOTE]
    > "localhost からこのアドインを開くことができません" というエラーが表示される場合は、ループバックの除外を有効にする必要があります。
    >
    > 1. Outlook を終了します。
    > 2. タスク マネージャー **を開** き、タスク **msoadfs.exeが** 実行されていないか確認します。
    > 3. 次のコマンドを実行します。
    >
    >     ```command&nbsp;line
    >     call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_http___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
    >     ```
    >
    > 4. Outlook を再起動します。

## <a name="debug"></a>Debug

独自の機能を実装する場合は、コードのデバッグが必要な場合があります。 イベント ベースのアドインのアクティブ化をデバッグする方法のガイダンスについては、「Debug [your event-based Outlook アドイン」を参照してください](debug-autolaunch.md)。

## <a name="event-based-activation-behavior-and-limitations"></a>イベント ベースのアクティブ化の動作と制限

イベントに基づいてアクティブ化するアドインは、実行時間が短く、軽量で、可能な限り非侵襲的である必要があります。 アドインが起動イベントの処理を完了したと知らされる場合は、アドインでメソッドを呼び出す必要 `event.completed` があります。 この呼び出しが行われた場合、アドインはイベント ベースのアドインを実行できる最大時間である約 300 秒以内にタイム アウトします。ユーザーが作成ウィンドウを閉じると、アドインも終了します。

ユーザーが同じイベントにサブスクライブしている複数のアドインがある場合、Outlook プラットフォームは特定の順序でアドインを起動します。 現在、アクティブに実行できるイベント ベースのアドインは 5 つのみです。 追加のアドインはキューにプッシュされ、以前にアクティブだったアドインが完了または非アクティブ化されると実行されます。

ユーザーは、アドインの実行を開始した現在のメール アイテムから切り替えまたは移動できます。 起動されたアドインは、バックグラウンドで操作を終了します。

UI Office.js変更する API の一部は、イベント ベースのアドインでは使用できない場合があります。ブロックされている API を次に示します。

- [ `Office.context.auth` : ] の下
  - `getAccessToken`
  - `getAccessTokenAsync`
- [ `Office.context.mailbox` : ] の下
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- [ `Office.context.mailbox.item` : ] の下
  - `close`
- [ `Office.context.ui` : ] の下
  - `displayDialogAsync`
  - `messageParent`

## <a name="see-also"></a>関連項目

- [Outlook アドインのマニフェスト](manifests.md)
- [イベント ベースのアドインをデバッグする方法](debug-autolaunch.md)
