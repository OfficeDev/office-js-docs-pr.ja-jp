---
title: イベント ベースのアクティブ化用に Outlook アドインを構成する (プレビュー)
description: イベント ベースのアクティブ化用に Outlook アドインを構成する方法について学習します。
ms.topic: article
ms.date: 01/06/2021
localization_priority: Normal
ms.openlocfilehash: d6893733af52bba7917531b2e8d5a442ce3dcd77
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839832"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a>イベント ベースのアクティブ化用に Outlook アドインを構成する (プレビュー)

イベント ベースのアクティブ化機能がない場合、ユーザーは自分のタスクを完了するためにアドインを明示的に起動する必要があります。 この機能を使用すると、特にすべてのアイテムに適用される操作に関して、特定のイベントに基づいてタスクを実行できます。 作業ウィンドウと UI を使用する機能を統合することもできます。 現時点では、次のイベントがサポートされています。

- `OnNewMessageCompose`: 新しいメッセージの作成時 (返信、全員に返信、転送を含む)
- `OnNewAppointmentOrganizer`: 新しい予定の作成時

  > [!IMPORTANT]
  > この機能は **、下** 書きや既存の予定など、アイテムの編集時にはアクティブ化されない。

このチュートリアルの終わりまでに、新しいメッセージが作成されるたびに実行されるアドインが作成されます。

> [!IMPORTANT]
> この機能は、Microsoft 365 サブスクリプションを使用した Outlook on the web でのプレビューでのみサポートされます。 [](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) 詳細 [については、この記事のイベント ベース](#how-to-preview-the-event-based-activation-feature) のアクティブ化機能をプレビューする方法を参照してください。
>
> プレビュー機能は予告なしに変更されることがありますので、実稼働アドインでは使用できません。

## <a name="how-to-preview-the-event-based-activation-feature"></a>イベント ベースのアクティブ化機能をプレビューする方法

イベント ベースのアクティブ化機能をお試しください。 GitHub を通じてフィードバックを提供することで、お客様のシナリオと改善方法をお知らせします (このページの最後にあるフィードバックセクションをご覧ください)。

この機能をプレビューするには、次の方法を使用します。

- CDN で **ベータ** ライブラリを参照します ( https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) . TypeScript [のコンパイルと](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) 読み取りIntelliSense定義ファイルは CDN と [DefinitelyTyped で見つかりました](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts)。 次の種類を使用してインストールできます `npm install --save-dev @types/office-js-preview` 。
- [Microsoft 365 テナントで対象指定リリースを構成します](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)。

## <a name="set-up-your-environment"></a>環境を設定する

Outlook クイック [スタートを完了](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) します。このクイック スタートでは、アドイン用の Yeoman ジェネレーターを使用してアドイン Office作成します。

## <a name="configure-the-manifest"></a>マニフェストを構成する

アドインのイベント ベースのアクティブ化を有効にするには、マニフェストで [Runtimes](../reference/manifest/runtimes.md) 要素と [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) 拡張ポイントを構成する必要があります。 今のところ、 `DesktopFormFactor` サポートされているフォーム ファクターは 1 つのみです。

1. コード エディターで、クイック スタート プロジェクトを開きます。

1. プロジェクトの **manifest.xml** にある新しいファイルを開きます。

1. ノード全体 `<VersionOverrides>` (開いているタグと閉じるタグを含む) を選択し、次の XML に置き換えてください。

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
  </VersionOverrides>
</VersionOverrides>
```

Outlook on Windows では JavaScript ファイルを使用し、Outlook on the web は同じ JavaScript ファイルを参照する HTML ファイルを使用します。 Outlook プラットフォームは最終的に Outlook クライアントに基づいて HTML または JavaScript を使用するかどうかを決定しますので、マニフェストでこれらの両方のファイルへの参照を提供する必要があります。 そのため、イベント処理を構成するには、要素内の HTML の場所を指定し、その子要素で、HTML によってインライン化または参照される JavaScript ファイルの場所を提供します `Runtime` `Override` 。

> [!TIP]
> Outlook アドインのマニフェストの詳細については、Outlook アドインの [マニフェストを参照してください](manifests.md)。

## <a name="implement-event-handling"></a>イベント処理を実装する

選択したイベントの処理を実装する必要があります。

このシナリオでは、新しいアイテムを作成する処理を追加します。

1. 同じクイック スタート プロジェクトから、コード エディターで **ファイル ./src/commands/commands.js** を開きます。

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

1. ファイルの最後に、次のステートメントを追加します。

    ```js
    g.onMessageComposeHandler = onMessageComposeHandler;
    g.onAppointmentComposeHandler = onAppointmentComposeHandler;
    ```

## <a name="try-it-out"></a>試してみる

1. プロジェクトのルート ディレクトリから次のコマンドを実行します。 このコマンドを実行すると、ローカル Web サーバーが起動します (まだ実行されていない場合)。

    ```command&nbsp;line
    npm run dev-server
    ```

1. 「[テスト用に Outlook アドインをサイドロードする](sideload-outlook-add-ins-for-testing.md)」の手順に従って Outlook アドインをサイドロードします。

1. Outlook on the web で新しいメッセージを作成します。

    ![Outlook on the web のメッセージ ウィンドウのスクリーンショット。件名が新規作成に設定されています。](../images/outlook-web-autolaunch.png)

## <a name="event-based-activation-behavior-and-limitations"></a>イベント ベースのアクティブ化の動作と制限事項

イベントに基づいてアクティブ化するアドインは、最大 330 秒の短い実行に設計されています。 起動イベントの処理が完了したというシグナルをアドインで呼び出す方法 `event.completed` をお勧めします。 ユーザーが作成ウィンドウを閉じると、アドインも終了します。

ユーザーが同じイベントにサブスクライブした複数のアドインを持っている場合、Outlook プラットフォームは特定の順序でアドインを起動します。 現在、アクティブに実行できるイベント ベースのアドインは 5 つのみです。 追加のアドインはキューにプッシュされ、以前にアクティブだったアドインが完了または非アクティブ化されると実行されます。

ユーザーは、アドインの実行を開始した現在のメール アイテムから切り替えまたは移動できます。 起動されたアドインは、バックグラウンドで操作を完了します。

UI Office.js変更する API の一部は、イベント ベースのアドインでは許可されません。ブロックされる API を次に示します。

- Under `Office.context.mailbox` :
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- Under `Office.context.ui` :
  - `displayDialogAsync`
  - `messageParent`
- Under `Office.context.auth` :
  - `getAccessToken`
  - `getAccessTokenAsync`

## <a name="see-also"></a>関連項目

[Outlook アドインのマニフェスト](manifests.md)
