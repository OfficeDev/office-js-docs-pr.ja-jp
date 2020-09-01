---
title: Outlook アドイン コマンド
description: Outlook アドイン コマンドは、ボタンやドロップダウン メニューを追加することにより、リボンから特定のアドイン操作を開始する方法を提供します。
ms.date: 07/07/2020
localization_priority: Priority
ms.openlocfilehash: 598d6e055b72d517d4a6bcfb90e3968b466e3aa0
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47294011"
---
# <a name="add-in-commands-for-outlook"></a><span data-ttu-id="89332-103">Outlook のアドイン コマンド</span><span class="sxs-lookup"><span data-stu-id="89332-103">Add-in commands for Outlook</span></span>

<span data-ttu-id="89332-p101">Outlook アドイン コマンドを作成して、ボタンまたはドロップダウン メニューとしてリボンに追加すると、リボンから特定のアドイン操作を開始できるようになり、ユーザーが簡単、直観的、かつ自然にアドインにアクセスできるようになります。アドイン コマンドを採用すれば、シームレスに機能性が向上するので、より魅力的なソリューションを作成することができます。</span><span class="sxs-lookup"><span data-stu-id="89332-p101">Outlook add-in commands provide ways to initiate specific add-in actions from the ribbon by adding buttons or drop-down menus. This lets users access add-ins in a simple, intuitive, and unobtrusive way. Because they offer increased functionality in a seamless manner, you can use add-in commands to create more engaging solutions.</span></span>

> [!NOTE]
> <span data-ttu-id="89332-107">アドイン コマンドは、Windows 用 Outlook 2013 以降、Mac 用 Outlook 2016 以降、iOS 用 Outlook、Android 用 Outlook、Exchange 2016 以降の Outlook on the web、Microsoft 365 の Outlook on the web および Outlook.com でのみ使用できます。</span><span class="sxs-lookup"><span data-stu-id="89332-107">Add-in commands are available only in Outlook 2013 or later on Windows, Outlook 2016 or later on Mac, Outlook on iOS, Outlook on Android, Outlook on the web for Exchange 2016 or later, and Outlook on the web for Microsoft 365 and Outlook.com.</span></span>
>
> <span data-ttu-id="89332-108">Outlook 2013 でのアドイン コマンドのサポートには、次の 3 つの更新プログラムが必要です。</span><span class="sxs-lookup"><span data-stu-id="89332-108">Support for add-in commands in Outlook 2013 requires three updates:</span></span>
> - [<span data-ttu-id="89332-109">2016 年 3 月 8 日にリリースされた Outlook 用セキュリティ更新プログラム</span><span class="sxs-lookup"><span data-stu-id="89332-109">March 8, 2016 security update for Outlook</span></span>](https://support.microsoft.com/kb/3114829)
> - [<span data-ttu-id="89332-110">2016 年 3 月 8 日にリリースされた Office 用セキュリティ更新プログラム (KB3114816)</span><span class="sxs-lookup"><span data-stu-id="89332-110">March 8, 2016 security update for Office (KB3114816)</span></span>](https://support.microsoft.com/help/3114816/march-8,-2016,-update-for-office-2013-kb3114816)
> - [<span data-ttu-id="89332-111">2016 年 3 月 8 日にリリースされた Office 用セキュリティ更新プログラム (KB3114828)</span><span class="sxs-lookup"><span data-stu-id="89332-111">March 8, 2016 security update for Office (KB3114828)</span></span>](https://support.microsoft.com/help/3114828/march-8,-2016,-update-for-office-2013-kb3114828)
>
> <span data-ttu-id="89332-112">Exchange 2016 のアドイン コマンドのサポートでは、[累積的な更新プログラム 5](https://support.microsoft.com/help/4012106/cumulative-update-5-for-exchange-server-2016) が必要です。</span><span class="sxs-lookup"><span data-stu-id="89332-112">Support for add-in commands in Exchange 2016 requires [Cumulative Update 5](https://support.microsoft.com/help/4012106/cumulative-update-5-for-exchange-server-2016).</span></span>

<span data-ttu-id="89332-p102">アドイン コマンドは、アクティブ化するアイテムの種類を制限する [ItemHasAttachment、ItemHasKnownEntity、ItemHasRegularExpressionMatch ルール](activation-rules.md) を使用しないアドインに対してのみ使用できます。ただし、[コンテキスト アドイン](contextual-outlook-add-ins.md) は、現在選択されているアイテムがメッセージか予定かに応じて異なるコマンドを表示でき、閲覧シナリオまたは作成シナリオのどちらで表示するかを選択できます。可能な場合はアドイン コマンドを使用するのが [ベスト プラクティス](../concepts/add-in-development-best-practices.md) です。</span><span class="sxs-lookup"><span data-stu-id="89332-p102">Add-in commands are only available for add-ins that do not use [ItemHasAttachment, ItemHasKnownEntity, or ItemHasRegularExpressionMatch rules](activation-rules.md) to limit the types of items they activate on. However, [contextual add-ins](contextual-outlook-add-ins.md) can present different commands depending on whether the currently selected item is a message or appointment, and can choose to appear in read or compose scenarios. Using add-in commands if possible is a [best practice](../concepts/add-in-development-best-practices.md).</span></span>

## <a name="creating-the-add-in-command"></a><span data-ttu-id="89332-116">アドイン コマンドの作成</span><span class="sxs-lookup"><span data-stu-id="89332-116">Creating the add-in command</span></span>

<span data-ttu-id="89332-p103">アドイン コマンドは、[VersionOverrides](../reference/manifest/versionoverrides.md) 要素のアドイン マニフェストで宣言されます。この要素はマニフェスト スキーマ v1.1 に追加されたもので、下位互換性が保証されています。`VersionOverrides` をサポートしていないクライアントでも、既存のアドインは引き続きアドイン コマンドのないときと変わらずに機能します。</span><span class="sxs-lookup"><span data-stu-id="89332-p103">Add-in commands are declared in the add-in manifest in the [VersionOverrides element](../reference/manifest/versionoverrides.md). This element is an addition to the manifest schema v1.1 that ensures backward compatibility. In a client that doesn't support `VersionOverrides`, existing add-ins will continue to function as they did without add-in commands.</span></span>

<span data-ttu-id="89332-120">`VersionOverrides` マニフェスト エントリは、アドインについての多くの事柄 (アプリケーション、リボンに追加するコントロールの種類、テキスト、アイコン、関連する機能など) を指定します。</span><span class="sxs-lookup"><span data-stu-id="89332-120">The `VersionOverrides` manifest entries specify many things for the add-in, such as the application, types of controls to add to the ribbon, the text, the icons, and any associated functions.</span></span>

<span data-ttu-id="89332-p104">アドインが、進行状況のインジケーターやエラー メッセージなど、状態更新を提供しなければならない場合、それは[通知 API](/javascript/api/outlook/office.notificationmessages) を通して行う必要があります。通知の処理を、マニフェストの `FunctionFile` ノードで指定されている別の HTML ファイルに定義する必要もあります。</span><span class="sxs-lookup"><span data-stu-id="89332-p104">When an add-in needs to provide status updates, such as progress indicators or error messages, it must do so through the [notification APIs](/javascript/api/outlook/office.notificationmessages). The processing for the notifications must also be defined in a separate HTML file that is specified in the `FunctionFile` node of the manifest.</span></span>

<span data-ttu-id="89332-p105">アドイン コマンドがリボンに合わせて適正に配置されるように、開発者は必要なサイズのアイコンをすべて定義する必要があります。必要とされるアイコンのサイズは、デスクトップの場合には 80 x 80 ピクセル、32 x 32 ピクセル、16 x 16 ピクセルで、モバイルの場合には 48 × 48 ピクセル、32 x 32 ピクセル、25 x 25 ピクセルです。</span><span class="sxs-lookup"><span data-stu-id="89332-p105">Developers should define icons for all required sizes so that the add-in commands will adjust smoothly along with the ribbon. The required icon sizes are 80 x 80 pixels, 32 x 32 pixels, and 16 x 16 pixels for desktop, and 48 x 48 pixels, 32 x 32 pixels, and 25 x 25 pixels for mobile.</span></span>

## <a name="how-do-add-in-commands-appear"></a><span data-ttu-id="89332-125">アドイン コマンドの表示方法</span><span class="sxs-lookup"><span data-stu-id="89332-125">How do add-in commands appear?</span></span>

<span data-ttu-id="89332-p106">アドイン コマンドは、リボン上にボタンとして表示されます。ユーザーがアドインをインストールすると、アドインのコマンドはボタン グループとして UI に表示されます。これは、リボンの既定のタブまたはカスタム タブのいずれかに表示されます。メッセージの場合、既定のタブは **[ホーム]** タブまたは **[メッセージ]** タブのいずれかです。予定表の場合、既定のタブは **[会議]** タブ、**[個別の会議]** タブ、**[定期的な会議]** タブ、または **[予定]** タブです。モジュール拡張機能の場合、既定のタブはカスタム タブです。既定タブでは、それぞれのアドインは 1 つのリボン グループを持つことができ、1 つのリボン グループに含まれるコマンドの数は 6 個までです。カスタム タブには、アドインのグループを 10 個まで含めることができ、1 つのグループにコマンドが 6 個まで表示されます。アドインに使用できるカスタム タブは 1 つに制限されています。</span><span class="sxs-lookup"><span data-stu-id="89332-p106">An add-in command appears on the ribbon as a button. When a user installs an add-in, its commands appear in the UI as a group of buttons. This can either be on the ribbon's default tab or on a custom tab. For messages, the default is either the **Home** or **Message** tab. For the calendar, the default is the **Meeting**, **Meeting Occurrence**, **Meeting Series**, or **Appointment** tab. For module extensions, the default is a custom tab. On the default tab, each add-in can have one ribbon group with up to 6 commands. On custom tabs, the add-in can have up to 10 groups, each with 6 commands. Add-ins are limited to only one custom tab.</span></span>

<span data-ttu-id="89332-131">リボンがいっぱいになると、アドイン コマンドがオーバーフロー メニューに表示されます。</span><span class="sxs-lookup"><span data-stu-id="89332-131">As the ribbon gets more crowded, add-in commands will be displayed in the overflow menu.</span></span> <span data-ttu-id="89332-132">通常、アドインのアドイン コマンドはグループ化されています。</span><span class="sxs-lookup"><span data-stu-id="89332-132">The add-in commands for an add-in are usually grouped together.</span></span>

![リボンのアドイン コマンド ボタン](../images/commands-normal.png)

![リボンとオーバーフロー メニューのアドイン コマンド ボタン](../images/commands-collapsed.png)

<span data-ttu-id="89332-p108">アドインにアドイン コマンドが追加されると、アドイン名は、アプリ バーから削除されます。リボン上のアドイン コマンド ボタンだけが残ります。</span><span class="sxs-lookup"><span data-stu-id="89332-p108">When an add-in command is added to an add-in, the add-in name is removed from the app bar. Only the add-in command button on the ribbon remains.</span></span>

### <a name="modern-outlook-on-the-web"></a><span data-ttu-id="89332-137">モダン Outlook on the web</span><span class="sxs-lookup"><span data-stu-id="89332-137">Modern Outlook on the web</span></span>

<span data-ttu-id="89332-138">Outlook on the web では、アドイン名はオーバーフロー メニューに表示されます。</span><span class="sxs-lookup"><span data-stu-id="89332-138">In Outlook on the web, the add-in name is displayed in an overflow menu.</span></span> <span data-ttu-id="89332-139">アドインに複数のアドイン コマンドがある場合、アドイン メニューを展開して、アドイン名のラベルが付いたボタンのグループを表示できます。</span><span class="sxs-lookup"><span data-stu-id="89332-139">If the add-in has multiple add-in commands, you can expand the add-in menu to see the group of buttons labeled with the add-in name.</span></span>

![アドイン コマンド ボタンが見つかるオーバーフロー メニュー](../images/commands-overflow-menu-web.png)

![アドイン コマンド ボタンを表示しているオーバーフローメニュー](../images/commands-overflow-menu-expand-web.png)

## <a name="what-ux-shapes-exist-for-add-in-commands"></a><span data-ttu-id="89332-142">アドイン コマンドの UX シェイプの目的</span><span class="sxs-lookup"><span data-stu-id="89332-142">What UX shapes exist for add-in commands?</span></span>

<span data-ttu-id="89332-p110">アドイン コマンドの UX シェイプは、さまざまな機能を実行できるボタンを含む Office アプリケーションのリボン タブで構成されています。現時点では、次の 3 つの UI シェイプがサポートされています:</span><span class="sxs-lookup"><span data-stu-id="89332-p110">The UX shape for an add-in command consists of a ribbon tab in the Office application that contains buttons that can perform various functions. Currently, three UI shapes are supported:</span></span>

- <span data-ttu-id="89332-145">JavaScript 関数を実行するボタン</span><span class="sxs-lookup"><span data-stu-id="89332-145">A button that executes a JavaScript function</span></span>
- <span data-ttu-id="89332-146">作業ウィンドウを起動するボタン</span><span class="sxs-lookup"><span data-stu-id="89332-146">A button that launches a task pane</span></span>
- <span data-ttu-id="89332-147">他の 2 種類のボタンについて 1 つ以上を選択肢とするドロップダウン メニューを表示するボタン</span><span class="sxs-lookup"><span data-stu-id="89332-147">A button that shows a drop-down menu with one or more buttons of the other two types</span></span>

### <a name="executing-a-javascript-function"></a><span data-ttu-id="89332-148">JavaScript 関数の実行</span><span class="sxs-lookup"><span data-stu-id="89332-148">Executing a JavaScript function</span></span>

<span data-ttu-id="89332-p111">JavaScript 関数を実行するアドイン コマンド ボタンは、操作を開始するためにユーザーが追加の選択をする必要のないシナリオで使用します。追跡や通知、印刷などの操作が該当します。また、ユーザーがサービスからより詳細な情報を取得するシナリオでも使用します。</span><span class="sxs-lookup"><span data-stu-id="89332-p111">Use an add-in command button that executes a JavaScript function for scenarios where the user doesn't need to make any additional selections to initiate the action. This can be for actions such as track, remind me, or print, or scenarios when the user wants more in-depth information from a service.</span></span>

<span data-ttu-id="89332-151">モジュール拡張機能では、メイン ユーザー インターフェイスのコンテンツを操作する JavaScript 関数をアドイン コマンド ボタンで実行できます。</span><span class="sxs-lookup"><span data-stu-id="89332-151">In module extensions, the add-in command button can execute JavaScript functions that interact with the content in the main user interface.</span></span>

![Outlook リボンの機能を実行するボタン。](../images/commands-uiless-button-1.png)

### <a name="launching-a-task-pane"></a><span data-ttu-id="89332-153">作業ウィンドウの起動</span><span class="sxs-lookup"><span data-stu-id="89332-153">Launching a task pane</span></span>

<span data-ttu-id="89332-p112">作業ウィンドウを起動するアドイン コマンド ボタンは、ユーザーが長時間アドインとの対話式操作を行う必要があるシナリオで使用します。たとえば、アドインでは設定の変更や多数のフィールドへの入力が必要になることがあります。</span><span class="sxs-lookup"><span data-stu-id="89332-p112">Use an add-in command button to launch a task pane for scenarios where a user needs to interact with an add-in for a longer period of time. For example, the add-in requires changes to settings or the completion of many fields.</span></span>

<span data-ttu-id="89332-p113">垂直作業ウィンドウの既定の幅は 320 px です。垂直作業ウィンドウのサイズは、Outlook エクスプローラーとインスペクターの両方で変更できます。このウィンドウのサイズは、To Do ウィンドウやリスト ビューのサイズを変更するときと同じ方法で変更することができます。</span><span class="sxs-lookup"><span data-stu-id="89332-p113">The default width of the vertical task pane is 320 px. The vertical task pane can be resized in both the Outlook Explorer and inspector. The pane can be resized in the same way the to-do pane and list view resize.</span></span>

![Outlook リボンの作業ウィンドウを開くボタン。](../images/commands-task-pane-button-1.png)

<br/>

<span data-ttu-id="89332-p114">このスクリーンショットは、垂直作業ウィンドウの例を示しています。左上隅にアドイン コマンドの名前が付いたウィンドウが開いています。ユーザーは、このウィンドウの使用後に、右上隅の **[X]** ボタンを使用してウィンドウを閉じることができます。既定では、このウィンドウはメッセージを越えて存続しません。アドインは、作業ウィンドウの[ピン留めをサポートする](pinnable-taskpane.md)ことができます。また、新しいメッセージが選択されたときには、イベントを受信できます。作業ウィンドウに表示されるすべての UI 要素は、アドインによって提示されます (アドイン名と閉じるボタンを除く)。</span><span class="sxs-lookup"><span data-stu-id="89332-p114">This screenshot shows an example of a vertical task pane. The pane opens with the name of the add-in command in the top left corner. Users can use the **X** button in the upper-right corner of the pane to close the add-in when they are finished using it. By default, this pane will not persist across messages. Add-ins can [support pinning](pinnable-taskpane.md) for the task pane and receive events when a new message is selected. All UI elements rendered in the task pane, aside from the add-in name and the close button, are provided by the add-in.</span></span>

<span data-ttu-id="89332-p115">ユーザーが作業ウィンドウを開く別のアドイン コマンドを選択すると、作業ウィンドウは直近に使用されたコマンドに置き換えられます。作業ウィンドウが開いているときにユーザーが、関数を実行するアドイン コマンド ボタンまたはドロップダウン メニューをクリックすると、操作が完了して、作業ウィンドウは開いたままになります。</span><span class="sxs-lookup"><span data-stu-id="89332-p115">If a user chooses another add-in command that opens a task pane, the task pane is replaced with the recently used command. If a user chooses an add-in command button that executes a function, or drop-down menu while the task pane is open, the action will be completed and the task pane will remain open.</span></span>

### <a name="drop-down-menu"></a><span data-ttu-id="89332-168">ドロップダウン メニュー</span><span class="sxs-lookup"><span data-stu-id="89332-168">Drop-down menu</span></span>

<span data-ttu-id="89332-p116">ドロップダウン メニュー アドイン コマンドでは、ボタンの静的リストを定義します。メニューには、機能を実行するボタンや作業ウィンドウを開くボタンを自由に組み合わせて含めることができます。サブメニューはサポートされません。</span><span class="sxs-lookup"><span data-stu-id="89332-p116">A drop-down menu add-in command defines a static list of buttons. The buttons within the menu can be any mix of buttons that execute a function or buttons that open a task pane. Submenus are not supported.</span></span>

![Outlook リボンのドロップダウン メニューを表示するボタン。](../images/commands-menu-button-1.png)

## <a name="where-do-add-in-commands-appear-in-the-ui"></a><span data-ttu-id="89332-173">UI でアドイン コマンドが表示される場所</span><span class="sxs-lookup"><span data-stu-id="89332-173">Where do add-in commands appear in the UI?</span></span>

<span data-ttu-id="89332-174">アドイン コマンドは次の 4 つのシナリオでサポートされています。</span><span class="sxs-lookup"><span data-stu-id="89332-174">Add-in commands are supported for four scenarios:</span></span>

### <a name="reading-a-message"></a><span data-ttu-id="89332-175">メッセージの閲覧</span><span class="sxs-lookup"><span data-stu-id="89332-175">Reading a message</span></span>

<span data-ttu-id="89332-176">ユーザーが閲覧ウィンドウまたはポップアウト閲覧フォームの **メッセージ** タブでメッセージを閲覧している間、既定のタブに追加されたアドイン コマンドは **ホーム** タブに表示されます。</span><span class="sxs-lookup"><span data-stu-id="89332-176">When the user is reading a message in the reading pane or in the **Message** tab for a pop-out read form, add-in commands added to the default tab appear on the **Home** tab.</span></span>

### <a name="composing-a-message"></a><span data-ttu-id="89332-177">メッセージの作成</span><span class="sxs-lookup"><span data-stu-id="89332-177">Composing a message</span></span>

<span data-ttu-id="89332-178">ユーザーがメッセージを作成している間は、既定のタブに追加されたアドイン コマンドが **[メッセージ]** タブに表示されます。</span><span class="sxs-lookup"><span data-stu-id="89332-178">When the user is composing a message, add-in commands added to the default tab appear on the **Message** tab.</span></span>

### <a name="creating-or-viewing-an-appointment-or-meeting-as-the-organizer"></a><span data-ttu-id="89332-179">開催者として予定または会議を作成または表示する</span><span class="sxs-lookup"><span data-stu-id="89332-179">Creating or viewing an appointment or meeting as the organizer</span></span>

<span data-ttu-id="89332-p117">開催者として予定または会議を作成または表示する場合、既定のタブに追加されたアドイン コマンドは、ポップアウト フォームの **[会議]**、**[個別の会議]**、**[定期的な会議]**、または **[予定]** のタブに表示されます。ただし、ユーザーが予定表のアイテムを選択してもポップ アウトを開かなければ、そのアドインのリボン グループはリボンに表示されません。</span><span class="sxs-lookup"><span data-stu-id="89332-p117">When creating or viewing an appointment or meeting as the organizer, add-in commands added to the default tab appear on the **Meeting**, **Meeting Occurrence**, **Meeting Series**, or **Appointment** tabs on pop-out forms. However, if the user selects an item in the calendar but doesn't open the pop-out, the add-in's ribbon group won't be visible in the ribbon.</span></span>

### <a name="viewing-a-meeting-as-an-attendee"></a><span data-ttu-id="89332-182">出席者として会議を表示する</span><span class="sxs-lookup"><span data-stu-id="89332-182">Viewing a meeting as an attendee</span></span>

<span data-ttu-id="89332-p118">出席者として会議を表示する場合、既定のタブに追加されたアドイン コマンドは、ポップアウト フォームの **[会議]**、**[個別の会議]**、または **[定期的な会議]** のタブに表示されます。ただし、ユーザーが予定表のアイテムを選択してもポップ アウトを開かなければ、そのアドインのリボン グループはリボンに表示されません。</span><span class="sxs-lookup"><span data-stu-id="89332-p118">When viewing a meeting as an attendee, add-in commands added to the default tab appear on the **Meeting**, **Meeting Occurrence**, or **Meeting Series** tabs on pop-out forms. However, if a user selects an item in the calendar but doesn't open the pop-out, the add-in's ribbon group won't be visible in the ribbon</span></span>

### <a name="using-a-module-extension"></a><span data-ttu-id="89332-185">モジュール拡張機能の使用</span><span class="sxs-lookup"><span data-stu-id="89332-185">Using a module extension</span></span>

<span data-ttu-id="89332-186">モジュール拡張機能を使用すると、モジュールのカスタム タブにアドイン コマンドが表示されます。</span><span class="sxs-lookup"><span data-stu-id="89332-186">When using a module extension, add-in commands appear on the extension's custom tab.</span></span>

## <a name="see-also"></a><span data-ttu-id="89332-187">関連項目</span><span class="sxs-lookup"><span data-stu-id="89332-187">See also</span></span>

- [<span data-ttu-id="89332-188">アドイン コマンド デモの Outlook アドイン</span><span class="sxs-lookup"><span data-stu-id="89332-188">Add-in command demo Outlook add-in</span></span>](https://github.com/officedev/outlook-add-in-command-demo)
- [<span data-ttu-id="89332-189">Excel、PowerPoint、Word のマニフェストにアドイン コマンドを作成する</span><span class="sxs-lookup"><span data-stu-id="89332-189">Create add-in commands in your manifest for Excel, PowerPoint, and Word</span></span>](../develop/create-addin-commands.md)
