---
title: イベントベースのライセンス認証用に Outlook アドインを構成する (プレビュー)
description: イベントベースのライセンス認証用に Outlook アドインを構成する方法について説明します。
ms.topic: article
ms.date: 08/11/2020
localization_priority: Normal
ms.openlocfilehash: f5df8c1efe5e1e5c4c83b1536e90d8f38729dcc3
ms.sourcegitcommit: 65c15a9040279901ea7ff7f522d86c8fddb98e14
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/14/2020
ms.locfileid: "46672723"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a>イベントベースのライセンス認証用に Outlook アドインを構成する (プレビュー)

イベントベースのアクティブ化機能がない場合、ユーザーは、自分のタスクを完了するためにアドインを明示的に起動する必要があります。 この機能により、特定のイベント (特に、すべてのアイテムに適用される操作) に基づいてアドインでタスクを実行できます。 作業ウィンドウと UI 非表示機能に統合することもできます。 現時点では、サポートされているイベントは次のとおりです。

- `OnNewMessageCompose`: 新しいメッセージを作成するときに (返信、全員に返信、および転送を含む)
- `OnNewAppointmentOrganizer`: 新しい予定を作成するとき

このチュートリアルを終了すると、新しいメッセージが作成されるたびに実行されるアドインができます。

> [!IMPORTANT]
> この機能は、Microsoft 365 サブスクリプションを使用する web 上の Outlook の [プレビュー](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) でのみサポートされています。 詳細については、この記事の「 [イベントに基づくライセンス認証機能をプレビューする方法](#how-to-preview-the-event-based-activation-feature) 」を参照してください。
>
> プレビュー機能は予告なしに変更される可能性があるため、運用アドインでは使用しないでください。

## <a name="how-to-preview-the-event-based-activation-feature"></a>イベントベースのライセンス認証機能をプレビューする方法

イベントに基づくライセンス認証機能をお試しください。 GitHub を通じてフィードバックを提供することによって、自分のシナリオと改善方法をお知らせください (このページの最後にある **フィードバック** セクションを参照してください)。

この機能をプレビューするには:

- CDN の **ベータ版** ライブラリを参照し https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) ます (。 TypeScript のコンパイルおよび IntelliSense 用の [型定義ファイル](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) は、CDN と、定義 [された](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts)定義ファイルにあります。 これらの種類は、でインストールでき `npm install --save-dev @types/office-js-preview` ます。
- [この要求フォーム](https://aka.ms/OWAPreview)を完了して送信することにより、Microsoft 365 アカウントを使用して、web 上の Outlook のプレビュービットへのアクセスを要求します。 テナントの準備が整ったことをお知らせします。

## <a name="set-up-your-environment"></a>環境を設定する

Outlook の [クイックスタート](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) に記入します。このアドインプロジェクトは、Office アドイン用の [アプリ] ジェネレーターを使用して作成されます。

## <a name="configure-the-manifest"></a>マニフェストを構成する

アドインのイベントベースのアクティブ化を有効にするには、マニフェストで、 [ランタイム](../reference/manifest/runtimes.md) 要素と [launchevent](../reference/manifest/extensionpoint.md#launchevent-preview) 拡張点を構成する必要があります。 ここで `DesktopFormFactor` は、サポートされているフォームファクターのみを示します。

1. コードエディターで、[クイックスタート] プロジェクトを開きます。

1. プロジェクトのルートにある **manifest.xml** ファイルを開きます。

1. `<VersionOverrides>`ノード全体 (open タグと close タグを含む) を選択し、次の XML に置き換えます。

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

Outlook on the Windows は JavaScript ファイルを使用しますが、web 上の Outlook は同じ JavaScript ファイルを参照する HTML ファイルを使用します。 Outlook プラットフォームは、outlook クライアントに基づいて HTML と JavaScript のどちらを使用するかを決定するために、これらのファイルへの参照をマニフェストに提供する必要があります。 そのため、イベント処理を構成するには、要素内の HTML の場所を指定し、 `Runtime` その `Override` 子要素で、JavaScript ファイルの場所を指定するか、html で参照します。

> [!TIP]
> Outlook アドインのマニフェストの詳細については、「 [outlook アドインのマニフェスト](manifests.md)」を参照してください。

## <a name="implement-event-handling"></a>イベント処理を実装する

選択したイベントの処理を実装する必要があります。

このシナリオでは、新しいアイテムを作成するための処理を追加します。

1. 同じクイックスタートプロジェクトから、コードエディターで **/src/commands/commands.js** を開きます。

1. 関数の後 `action` に、次の JavaScript 関数を挿入します。

    ```js
    function onMessageComposeHandler(event) {
      setSubject();
      event.completed();
    }
    function onAppointmentComposeHandler(event) {
      setSubject();
      event.completed();
    }
    function setSubject() {
      Office.context.mailbox.item.subject.setAsync("Set by an event-based add-in!");
    }
    ```

1. ファイルの末尾に、次のステートメントを追加します。

    ```js
    g.onMessageComposeHandler = onMessageComposeHandler;
    g.onAppointmentComposeHandler = onAppointmentComposeHandler;
    ```

## <a name="try-it-out"></a>試してみる

1. プロジェクトのルート ディレクトリから次のコマンドを実行します。 このコマンドを実行すると、ローカル Web サーバーが起動します (まだ実行されていない場合)。

    ```command&nbsp;line
    npm start
    ```

1. 「[テスト用に Outlook アドインをサイドロードする](sideload-outlook-add-ins-for-testing.md)」の手順に従って Outlook アドインをサイドロードします。

1. Outlook on the web で新しいメッセージを作成します。

    ![Web 上の Outlook で、作成時に件名が設定されたメッセージウィンドウのスクリーンショット。](../images/outlook-web-autolaunch.png)

## <a name="event-based-activation-behavior-and-limitations"></a>イベントに基づくアクティブ化の動作と制限事項

イベントに基づいてアクティブ化するアドインは、短時間で最大330秒のみが実行されるように設計されています。 アドインで、 `event.completed` 起動イベントの処理が完了したことを通知するメソッドを呼び出すことをお勧めします。 ユーザーが [新規作成] ウィンドウを閉じたときにもアドインは終了します。

ユーザーが同じイベントにサブスクライブした複数のアドインを持っている場合、Outlook プラットフォームは特定の順序でアドインを起動しません。 現在、5つのイベントベースのアドインのみをアクティブに実行できます。 追加のアドインは、キューにプッシュされ、以前のアクティブなアドインが完了または非アクティブになったときとして実行されます。

ユーザーは、アドインの実行が開始された現在のメールアイテムから切り替えることができます。 起動されたアドインは、バックグラウンドで操作を終了します。

UI を変更または変更する一部の Office.js Api は、イベントベースのアドインからは許可されていません。ブロックされる Api を次に示します。

- `Office.context.mailbox`以下:
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- `Office.context.ui`以下:
  - `displayDialogAsync`
  - `messageParent`
- `Office.context.auth`以下:
  - `getAccessTokenAsync`

## <a name="see-also"></a>関連項目

[Outlook アドインのマニフェスト](manifests.md)
