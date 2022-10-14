---
title: 複数のメッセージで Outlook アドインをアクティブ化する (プレビュー)
description: 複数のメッセージが選択されているときに Outlook アドインをアクティブ化する方法について説明します。
ms.topic: article
ms.date: 10/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 9d81d698facfc4535b3945d8cee4c97492fc8a88
ms.sourcegitcommit: 5544cf174d145e356e33866e2480bde999514ada
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/14/2022
ms.locfileid: "68574145"
---
# <a name="activate-your-outlook-add-in-on-multiple-messages-preview"></a>複数のメッセージで Outlook アドインをアクティブ化する (プレビュー)

アイテムの複数選択機能を使用すると、Outlook アドインで、選択した複数のメッセージを 1 回の実行でアクティブ化し、操作を実行できるようになりました。 Customer Relationship Management (CRM) システムへのメッセージのアップロード、多数のアイテムの分類など、特定の操作をワンクリックで簡単に完了できるようになりました。

次のセクションでは、読み取りモードで複数のメッセージの件名行を取得するようにアドインを構成する方法について説明します。

> [!IMPORTANT]
> アイテムの複数選択機能は、Outlook on Windows の Microsoft 365 サブスクリプションのプレビューでのみ使用できます。 プレビューの機能は、運用環境のアドインでは使用しないでください。この機能をテストまたは開発環境でテストし、GitHub を通じて体験に関するフィードバックをお待ちしております (このページの最後にある **フィードバック** セクションを参照してください)。

> [!NOTE]
> アイテムの複数選択機能は、Teams [マニフェスト (プレビュー)](../develop/json-manifest-overview.md) では現在サポートされていませんが、チームはこれを利用できるように取り組んでいます。

## <a name="prerequisites-to-preview-item-multi-select"></a>アイテムの複数選択をプレビューするための前提条件

複数選択機能をプレビューするには、バージョン 2209 (ビルド 15629.20110) 以降の Outlook on Windows をインストールします。 インストールが完了したら、 [Office Insider プログラム](https://insider.office.com/join/windows) に参加し、 **ベータ チャネル** オプションを選択して Office ベータ ビルドにアクセスします。

## <a name="set-up-your-environment"></a>環境を設定する

Office アドイン[用 Yeoman ジェネレーター](../develop/yeoman-generator-overview.md)を使用してアドイン プロジェクトを作成するには[、Outlook クイック スタート](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator)を完了します。

## <a name="configure-the-manifest"></a>マニフェストを構成する

選択した複数のメッセージでアドインをアクティブにするには、[SupportsMultiSelect](/javascript/api/manifest/action?view=outlook-js-preview&preserve-view=true#supportsmultiselect-preview) 子要素を要素に追加し、その値`true`を **\<Action\>** . アイテムの複数選択は現時点ではメッセージのみをサポートするため、要素の **\<ExtensionPoint\>**`xsi:type`属性値を設定するか、または`MessageComposeCommandSurface`に設定する`MessageReadCommandSurface`必要があります。

1. 優先するコード エディターで、作成した Outlook クイック スタート プロジェクトを開きます。

1. プロジェクトのルートにある **manifest.xml** ファイルを開きます。

1. 要素に値を **\<Permissions\>**`ReadWriteMailbox`割り当てます。

    ```xml
    <Permissions>ReadWriteMailbox</Permissions>
    ```

1. ノード全体 **\<VersionOverrides\>** を選択し、次の XML に置き換えます。

    ```xml
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
        <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
            <Requirements>
                <bt:Sets DefaultMinVersion="1.12">
                  <bt:Set Name="Mailbox"/>
                </bt:Sets>
            </Requirements>
            <Hosts>
                <Host xsi:type="MailHost">
                    <DesktopFormFactor>
                        <!-- Message Read mode-->
                        <ExtensionPoint xsi:type="MessageReadCommandSurface">
                            <OfficeTab id="TabDefault">
                                <Group id="msgReadGroup">
                                    <Label resid="GroupLabel"/>
                                    <Control xsi:type="Button" id="msgReadOpenPaneButton">
                                        <Label resid="TaskpaneButton.Label"/>
                                        <Supertip>
                                            <Title resid="TaskpaneButton.Label"/>
                                            <Description resid="TaskpaneButton.Tooltip"/>
                                        </Supertip>
                                        <Icon>
                                            <bt:Image size="16" resid="Icon.16x16"/>
                                            <bt:Image size="32" resid="Icon.32x32"/>
                                            <bt:Image size="80" resid="Icon.80x80"/>
                                        </Icon>
                                        <Action xsi:type="ShowTaskpane">
                                            <SourceLocation resid="Taskpane.Url"/>
                                            <!-- Enables your add-in to activate on multiple selected messages. -->
                                            <SupportsMultiSelect>true</SupportsMultiSelect>
                                        </Action>
                                    </Control>
                                </Group>
                            </OfficeTab>
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
                  <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
                </bt:Urls>
                <bt:ShortStrings>
                  <bt:String id="GroupLabel" DefaultValue="Item Multi-select"/>
                  <bt:String id="TaskpaneButton.Label" DefaultValue="Show Taskpane"/>
                </bt:ShortStrings>
                <bt:LongStrings>
                  <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Opens a pane which displays an option to retrieve the subject line of selected messages."/>
                </bt:LongStrings>
            </Resources>
        </VersionOverrides>
    </VersionOverrides>
    ```

1. 変更内容を保存します。

## <a name="configure-the-task-pane"></a>作業ウィンドウを構成する

アイテムの複数選択は [、SelectedItemsChanged](/javascript/api/office/office.eventtype?view=outlook-js-preview&preserve-view=true) イベントに依存して、メッセージがいつ選択または選択解除されるかを判断します。 このイベントには作業ウィンドウの実装が必要です。

1. **./src/taskpane** フォルダーから、**taskpane.html** を開きます。

1. 要素で **\<script\>**、属性`"https://appsforoffice.microsoft.com/lib/beta/hosted/office.js"`を `src` . これは、コンテンツ配信ネットワーク (CDN) 上のベータ ライブラリを参照します。

    ```html
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/beta/hosted/office.js"></script>
    ```

1. **\<body\>** 要素で、要素全体 **\<main\>** を次のマークアップに置き換えます。

    ```html
    <main id="app-body" class="ms-welcome__main">
        <h2 class="ms-font-xl">Retrieve the subject line of multiple messages with one click!</h2>
        <ul id="selected-items"></ul>
        <div role="button" id="run" class="ms-welcome__action ms-Button ms-Button--hero ms-font-xl">
            <span class="ms-Button-label">Run</span>
        </div>
    </main>
    ```

1. 変更内容を保存します。

## <a name="implement-a-handler-for-the-selecteditemschanged-event"></a>SelectedItemsChanged イベントのハンドラーを実装する

イベントが発生したときにアドインに `SelectedItemsChanged` 通知するには、メソッドを使用してイベント ハンドラーを登録する `addHandlerAsync` 必要があります。

1. **./src/taskpane** フォルダーから、**taskpane.js** を開きます。

1. コールバック関数で `Office.onReady()` 、既存のコードを次のように置き換えます。

    ```javascript
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
        document.getElementById("run").onclick = run;
    
        // Register an event handler to identify when messages are selected.
        Office.context.mailbox.addHandlerAsync(Office.EventType.SelectedItemsChanged, run, asyncResult => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log(asyncResult.error.message);
            return;
          }
    
          console.log("Event handler added.");
        });
    }
    ```

## <a name="retrieve-the-subject-line-of-selected-messages"></a>選択したメッセージの件名行を取得する

イベント ハンドラーを登録したので、 [getSelectedItemsAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#outlook-office-mailbox-getselecteditemsasync-member(1)) メソッドを呼び出して、選択したメッセージの件名を取得し、作業ウィンドウに記録します。 このメソッドは `getSelectedItemsAsync` 、アイテム ID、アイテムの種類 (`Message` 現時点でサポートされている唯一の型)、アイテム モード (`Read` または `Compose`) など、他のメッセージ プロパティを取得するためにも使用できます。

1. **taskpane.js** 関数に`run`移動し、次のコードを挿入します。

    ```javascript
    // Clear list of previously selected messages, if any.
    const list = document.getElementById("selected-items");
    while (list.firstChild) {
        list.removeChild(list.firstChild);
    }

    // Retrieve the subject line of the selected messages and log it to a list in the task pane.
    Office.context.mailbox.getSelectedItemsAsync(asyncResult => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log(asyncResult.error.message);
            return;      
        }

        asyncResult.value.forEach(item => {
            const listItem = document.createElement("li");
            listItem.textContent = item.subject;
            list.appendChild(listItem);
        });
    });
    ```

1. 変更内容を保存します。

## <a name="try-it-out"></a>試してみる

1. ターミナルから、プロジェクトのルート ディレクトリで次のコードを実行します。 これにより、ローカル Web サーバーが起動し、アドインがサイドロードされます。

    ```command line
    npm start
    ```

    > [!TIP]
    > アドインが自動的にサイドロードされない場合は、 [テスト用の Sideload Outlook アドイン](sideload-outlook-add-ins-for-testing.md?tabs=windows#outlook-on-the-desktop) の手順に従って、Outlook で手動でサイドロードします。

1. Outlook で、閲覧ウィンドウが有効になっていることを確認します。 閲覧ウィンドウを有効にするには、「閲覧ウィンドウ [を使用して、メッセージをプレビューするように構成する」](https://support.microsoft.com/office/2fd687ed-7fc4-4ae3-8eab-9f9b8c6d53f0)を参照してください。

1. 受信トレイに移動し、 **Ctrl キーを押** しながらメッセージを選択することで、複数のメッセージを選択します。

1. リボンから [ **タスクウィンドウの表示** ] を選択します。

1. 作業ウィンドウで [ **実行** ] を選択し、選択したメッセージの件名行の一覧を表示します。

    :::image type="content" source="../images/outlook-multi-select.png" alt-text="選択した複数のメッセージから取得された件名行のサンプル リスト。":::

## <a name="item-multi-select-behavior-and-limitations"></a>項目の複数選択の動作と制限事項

アイテムの複数選択は、読み取りモードと作成モードの両方で Exchange メールボックス内のメッセージのみをサポートします。 Outlook アドインは、次の条件が満たされている場合にのみ、複数のメッセージに対してアクティブ化されます。

- メッセージは、一度に 1 つの Exchange メールボックスから選択する必要があります。 Exchange 以外のメールボックスはサポートされていません。
- メッセージは、一度に 1 つのメールボックス フォルダーから選択する必要があります。 複数のメッセージが異なるフォルダーに配置されている場合、会話ビューが有効でない限り、アドインはアクティブ化されません。 詳細については、「 [会話内の複数選択」を](#multi-select-in-conversations)参照してください。
- アドインは、イベントを検出するために作業ウィンドウを実装する `SelectedItemsChanged` 必要があります。
- Outlook の [閲覧ウィンドウ](https://support.microsoft.com/office/2fd687ed-7fc4-4ae3-8eab-9f9b8c6d53f0) を有効にする必要があります。
- 一度に最大 100 個のメッセージを選択できます。

> [!NOTE]
> 会議の招待と応答は、予定ではなくメッセージと見なされるため、選択に含めることができます。

### <a name="multi-select-in-conversations"></a>会話内の複数選択

アイテムの複数選択は、メールボックスまたは特定のフォルダーで有効になっているかどうかに関係なく [、Conversations ビュー](https://support.microsoft.com/office/0eeec76c-f59b-4834-98e6-05cfdfa9fb07) をサポートします。 次の表では、会話が展開または折りたたまれている場合、会話ヘッダーが選択されたとき、および会話メッセージが現在表示されているフォルダーとは別のフォルダーにある場合に想定される動作について説明します。

|Selection|展開された会話ビュー|折りたたまれた会話ビュー|
|------|------|------|
|**会話ヘッダーが選択されている**|会話ヘッダーが選択されている唯一のアイテムの場合、複数選択をサポートするアドインはアクティブ化されません。 ただし、他の非ヘッダー メッセージも選択されている場合、アドインは選択したヘッダーではなく、そのメッセージに対してのみアクティブ化されます。|メッセージの選択には、最新のメッセージ (つまり、会話スタックの最初のメッセージ) が含まれます。<br><br>会話内の最新のメッセージが現在表示されているフォルダーから別のフォルダーにある場合は、現在のフォルダーにあるスタック内の後続のメッセージが選択に含まれます。|
|**選択した会話メッセージは、現在表示されているメッセージと同じフォルダーにあります**|選択したすべての会話メッセージが選択に含まれます。|該当なし。 折りたたまれた会話ビューでは、会話ヘッダーのみを選択できます。|
|**選択した会話メッセージは、現在表示されているフォルダーとは異なるフォルダーに配置されます** |選択したすべての会話メッセージが選択に含まれます。|該当なし。 折りたたまれた会話ビューでは、会話ヘッダーのみを選択できます。|

## <a name="next-steps"></a>次の手順

選択した複数のメッセージを操作するアドインを有効にしたので、アドインの機能を拡張し、ユーザー エクスペリエンスをさらに向上させることができます。 [Exchange Web Services (EWS)](web-services.md) や [Microsoft Graph](/graph/overview) などのサービスで、選択したメッセージのアイテム ID を使用して、より複雑な操作を実行します。

## <a name="see-also"></a>関連項目

- [Outlook アドインのマニフェスト](manifests.md)
- [Outlook アドインから Web サービスを呼び出す](web-services.md)
- [Microsoft Graph の概要](/graph/overview)
