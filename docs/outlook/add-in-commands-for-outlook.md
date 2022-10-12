---
title: Outlook アドイン コマンド
description: Outlook アドイン コマンドは、ボタンやドロップダウン メニューを追加することにより、リボンから特定のアドイン操作を開始する方法を提供します。
ms.date: 10/11/2022
ms.localizationpriority: high
ms.openlocfilehash: d029fd4acc1a32c912c73d6e5f468b9c217b9262
ms.sourcegitcommit: 787fbe4d4a5462ff6679ad7fd00748bf07391610
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/12/2022
ms.locfileid: "68546460"
---
# <a name="add-in-commands-for-outlook"></a>Outlook のアドイン コマンド

Outlook add-in commands provide ways to initiate specific add-in actions from the ribbon by adding buttons or drop-down menus. This lets users access add-ins in a simple, intuitive, and unobtrusive way. Because they offer increased functionality in a seamless manner, you can use add-in commands to create more engaging solutions.

> [!NOTE]
> アドイン コマンドは、Windows 用 Outlook 2013 以降、Mac 用 Outlook 2016 以降、iOS 用 Outlook、Android 用 Outlook、Exchange 2016 以降の Outlook on the web、Microsoft 365 の Outlook on the web および Outlook.com でのみ使用できます。
>
> Outlook 2013 でのアドイン コマンドのサポートには、次の 3 つの更新プログラムが必要です。
> - [2016 年 3 月 8 日にリリースされた Outlook 用セキュリティ更新プログラム](https://support.microsoft.com/kb/3114829)
> - [2016 年 3 月 8 日にリリースされた Office 用セキュリティ更新プログラム (KB3114816)](https://support.microsoft.com/topic/3d3eb171-78c2-0e61-62a2-85723bc4bcc0)
> - [2016 年 3 月 8 日にリリースされた Office 用セキュリティ更新プログラム (KB3114828)](https://support.microsoft.com/topic/54437016-d1e0-7aac-dbb7-4ecfbd57f5f0)
>
> Exchange 2016 のアドイン コマンドのサポートでは、[累積的な更新プログラム 5](https://support.microsoft.com/topic/d67d7693-96a4-fb6e-b60b-e64984e267bd) が必要です。

> [!TIP]
> アドインで XML マニフェストを使用している場合、アドイン コマンドは [、ItemHasAttachment、ItemHasKnownEntity、または ItemHasRegularExpressionMatch ルール](activation-rules.md) を使用してアクティブ化するアイテムの種類を制限しないアドインでのみ使用できます。 ただし、 [コンテキスト アドイン](contextual-outlook-add-ins.md) は、現在選択されているアイテムがメッセージか予定かに応じて異なるコマンドを表示し、読み取りまたは作成のシナリオに表示するかを選択できます。 可能な場合はアドイン コマンドを使用するのが[ベスト プラクティス](../concepts/add-in-development-best-practices.md)です。

## <a name="create-the-ui-for-the-add-in-command"></a>アドイン コマンドの UI を作成する

アドイン コマンドは、アドイン マニフェストで宣言されます。 マークアップはマニフェストの種類によって異なります。

# <a name="xml-manifest"></a>[XML マニフェスト](#tab/xmlmanifest)

アドイン コマンドは [、VersionOverrides 要素](/javascript/api/manifest/versionoverrides)で宣言されます。 この要素は、下位互換性を確保する XML マニフェスト スキーマ v1.1 への追加です。 **\<VersionOverrides\>** をサポートしていないクライアントでは、既存のアドインは引き続きアドイン コマンドなしで機能します。

その **\<VersionOverrides\>** マニフェスト エントリでは、アプリケーション、リボンに追加するコントロールのタイプ、テキスト、アイコン、および関連する関数など、アドインの多くのものを指定します。

When an add-in needs to provide status updates, such as progress indicators or error messages, it must do so through the [notification APIs](/javascript/api/outlook/office.notificationmessages). The processing for the notifications must also be defined in a separate HTML file that is specified in the `FunctionFile` node of the manifest.

# <a name="teams-manifest-developer-preview"></a>[Teams マニフェスト (開発者プレビュー)](#tab/jsonmanifest)

アドイン コマンドは、"extensions.runtimes" プロパティと "extensions.ribbons" プロパティで宣言されます。 これらのプロパティは、アプリケーション、リボンに追加するコントロールの種類、テキスト、アイコン、関連する関数など、アドインの多くのものを指定します。

アドインから進捗状況の更新 (進行状況インジケータやエラー メッセージなど) を提供する必要がある場合は、 [通知 API](/javascript/api/outlook/office.notificationmessages) を介して行う必要があります。 通知の処理は、マニフェストの "runtimes.code.page" プロパティで指定された別の HTML ファイルでも定義する必要があります。

---
### <a name="icons"></a>アイコン

Developers should define icons for all required sizes so that the add-in commands will adjust smoothly along with the ribbon. The required icon sizes are 80 x 80 pixels, 32 x 32 pixels, and 16 x 16 pixels for desktop, and 48 x 48 pixels, 32 x 32 pixels, and 25 x 25 pixels for mobile.

## <a name="how-do-add-in-commands-appear"></a>アドイン コマンドの表示方法

アドイン コマンドは、リボンにボタンまたはドロップダウン メニューの項目として表示されます。 ユーザーがアドインをインストールすると、そのコマンドはボタンのグループとして UI に表示されます。 これは、リボンの既定のタブまたはカスタム タブ上のいずれかに表示されます。メッセージの場合、既定のタブは **[ホーム]** または **[メッセージ]** タブのいずれかです。予定表の場合、既定は **[会議]**、**[個別の会議]**、**[定期的な会議]**、または **[予定]** です。モジュール拡張機能の場合、既定のタブはカスタム タブです。既定タブでは、それぞれのアドインは 1 つのリボン グループを持つことができ、1 つのリボン グループに含まれるコマンドの数は 6 個までです。 カスタム タブには、アドインのグループを 10 個まで含めることができ、1 つのグループにコマンドが 6 個まで表示されます。 アドインに使用できるカスタム タブは 1 つのみに制限されています。

リボンがいっぱいになると、アドイン コマンドがオーバーフロー メニューに表示されます。 通常、アドインのアドイン コマンドはグループ化されています。

![リボンのアドイン コマンド ボタン。](../images/commands-normal.png)

![リボンとオーバーフロー メニューのアドイン コマンド ボタン。](../images/commands-collapsed.png)

When an add-in command is added to an add-in, the add-in name is removed from the app bar. Only the add-in command button on the ribbon remains.

### <a name="modern-outlook-on-the-web"></a>モダン Outlook on the web

Outlook on the web では、アドイン名はオーバーフロー メニューに表示されます。 アドインに複数のアドイン コマンドがある場合、アドイン メニューを展開して、アドイン名のラベルが付いたボタンのグループを表示できます。

![アドイン コマンド ボタンが見つかるオーバーフロー メニュー。](../images/commands-overflow-menu-web.png)

![アドイン コマンド ボタンを表示しているオーバーフロー メニュー。](../images/commands-overflow-menu-expand-web.png)

## <a name="what-are-the-types-of-add-in-commands"></a>アドイン コマンド タイプは何ですか?

アドイン コマンドのための UI は、リボン ボタンまたはドロップダウン メニューの項目で構成されます。 コマンドがトリガーするアクションの種類に基づいて、2 種類のアドイン コマンドがあります。

- **作業ウィンドウ コマンド**: ボタンまたはメニュー項目によって、アドインの作業ウィンドウが開きます。 この種のアドイン コマンドをマニフェスト内のマークアップと共に追加します。 コマンドの "分離コード" は Office に指定されます。
- **関数コマンド**: ボタンまたはメニュー項目は任意の JavaScript を実行します。 ほとんどの場合、このコードは Office JavaScript ライブラリで API を呼び出しますが、そうする必要はありません。 この種類のアドインでは、通常、ボタンまたはメニュー項目自体以外の UI は表示されません。 関数コマンドについては、次の点に注意してください。

   - トリガーされる関数は [displayDialogAsync](/javascript/api/office/office.ui#office-office-ui-displaydialogasync-member(1)) メソッドを呼び出してダイアログを表示できます。これは、エラーの表示、進行状況の表示、またはユーザーからの入力を求める適切な方法です。
   - 関数コマンドを実行するランタイムは、 [ブラウザーベースの](../testing/runtimes.md#browser-runtime)完全なランタイムです。 HTML をレンダリングし、インターネットに呼び出してデータを送信または取得できます。

### <a name="run-a-function-command"></a>関数コマンドを実行する

Use an add-in command button that executes a JavaScript function for scenarios where the user doesn't need to make any additional selections to initiate the action. This can be for actions such as track, remind me, or print, or scenarios when the user wants more in-depth information from a service.

モジュール拡張機能では、メイン ユーザー インターフェイスのコンテンツを操作する JavaScript 関数をアドイン コマンド ボタンで実行できます。

![Outlook リボンの機能を実行するボタン。](../images/commands-uiless-button-1.png)

### <a name="launch-a-task-pane"></a>作業ウィンドウの起動

Use an add-in command button to launch a task pane for scenarios where a user needs to interact with an add-in for a longer period of time. For example, the add-in requires changes to settings or the completion of many fields.

The default width of the vertical task pane is 320 px. The vertical task pane can be resized in both the Outlook Explorer and inspector. The pane can be resized in the same way the to-do pane and list view resize.

![Outlook リボンの作業ウィンドウを開くボタン。](../images/commands-task-pane-button-1.png)

<br/>

This screenshot shows an example of a vertical task pane. The pane opens with the name of the add-in command in the top left corner. Users can use the **X** button in the upper-right corner of the pane to close the add-in when they are finished using it. By default, this pane will not persist across messages. Add-ins can [support pinning](pinnable-taskpane.md) for the task pane and receive events when a new message is selected. All UI elements rendered in the task pane, aside from the add-in name and the close button, are provided by the add-in.

If a user chooses another add-in command that opens a task pane, the task pane is replaced with the recently used command. If a user chooses an add-in command button that executes a function, or drop-down menu while the task pane is open, the action will be completed and the task pane will remain open.

### <a name="drop-down-menu"></a>ドロップダウン メニュー

ドロップダウン メニュー アドイン コマンドでは、項目の静的リストを定義します。 メニューには、機能を実行する項目や作業ウィンドウを開く項目を自由に組み合わせて含めることができます。 サブメニューはサポートされません。

![Outlook リボンのドロップダウン メニューを表示するボタン。](../images/commands-menu-button-1.png)

## <a name="where-do-add-in-commands-appear-in-the-ui"></a>UI でアドイン コマンドが表示される場所

アドイン コマンドは次の 4 つのシナリオでサポートされています。

### <a name="reading-a-message"></a>メッセージの閲覧

ユーザーが閲覧ウィンドウまたはポップアウト閲覧フォームの **メッセージ** タブでメッセージを閲覧している間、既定のタブに追加されたアドイン コマンドは **ホーム** タブに表示されます。

### <a name="composing-a-message"></a>メッセージの作成

ユーザーがメッセージを作成している間は、既定のタブに追加されたアドイン コマンドが **[メッセージ]** タブに表示されます。

### <a name="creating-or-viewing-an-appointment-or-meeting-as-the-organizer"></a>開催者として予定または会議を作成または表示する

When creating or viewing an appointment or meeting as the organizer, add-in commands added to the default tab appear on the **Meeting**, **Meeting Occurrence**, **Meeting Series**, or **Appointment** tabs on pop-out forms. However, if the user selects an item in the calendar but doesn't open the pop-out, the add-in's ribbon group won't be visible in the ribbon.

### <a name="viewing-a-meeting-as-an-attendee"></a>出席者として会議を表示する

When viewing a meeting as an attendee, add-in commands added to the default tab appear on the **Meeting**, **Meeting Occurrence**, or **Meeting Series** tabs on pop-out forms. However, if a user selects an item in the calendar but doesn't open the pop-out, the add-in's ribbon group won't be visible in the ribbon

### <a name="using-a-module-extension"></a>モジュール拡張機能の使用

モジュール拡張機能を使用すると、モジュールのカスタム タブにアドイン コマンドが表示されます。

## <a name="see-also"></a>関連項目

- [アドイン コマンド デモの Outlook アドイン](https://github.com/officedev/outlook-add-in-command-demo)
- [Excel、PowerPoint、Word のマニフェストにアドイン コマンドを作成する](../develop/create-addin-commands.md)
- [Outlook アドインで関数コマンドをデバッグする](debug-ui-less.md)
- [チュートリアル: メッセージ作成 Outlook アドインのビルド](../tutorials/outlook-tutorial.md)
