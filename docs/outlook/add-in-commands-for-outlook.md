---
title: Outlook アドイン コマンド
description: Outlook アドイン コマンドは、ボタンやドロップダウン メニューを追加することにより、リボンから特定のアドイン操作を開始する方法を提供します。
ms.date: 07/07/2020
localization_priority: Priority
ms.openlocfilehash: 7705c168077d2a704ff16b05bfb82416cd7f4154
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094030"
---
# <a name="add-in-commands-for-outlook"></a>Outlook のアドイン コマンド

Outlook add-in commands provide ways to initiate specific add-in actions from the ribbon by adding buttons or drop-down menus. This lets users access add-ins in a simple, intuitive, and unobtrusive way. Because they offer increased functionality in a seamless manner, you can use add-in commands to create more engaging solutions.

> [!NOTE]
> アドイン コマンドは、Windows 用 Outlook 2013 以降、Mac 用 Outlook 2016 以降、iOS 用 Outlook、Android 用 Outlook、Exchange 2016 以降の Outlook on the web、Microsoft 365 の Outlook on the web および Outlook.com でのみ使用できます。
>
> Outlook 2013 でのアドイン コマンドのサポートには、次の 3 つの更新プログラムが必要です。
> - [2016 年 3 月 8 日にリリースされた Outlook 用セキュリティ更新プログラム](https://support.microsoft.com/kb/3114829)
> - [2016 年 3 月 8 日にリリースされた Office 用セキュリティ更新プログラム (KB3114816)](https://support.microsoft.com/help/3114816/march-8,-2016,-update-for-office-2013-kb3114816)
> - [2016 年 3 月 8 日にリリースされた Office 用セキュリティ更新プログラム (KB3114828)](https://support.microsoft.com/help/3114828/march-8,-2016,-update-for-office-2013-kb3114828)
>
> Exchange 2016 のアドイン コマンドのサポートでは、[累積的な更新プログラム 5](https://support.microsoft.com/help/4012106/cumulative-update-5-for-exchange-server-2016) が必要です。

Add-in commands are only available for add-ins that do not use [ItemHasAttachment, ItemHasKnownEntity, or ItemHasRegularExpressionMatch rules](activation-rules.md) to limit the types of items they activate on. However, [contextual add-ins](contextual-outlook-add-ins.md) can present different commands depending on whether the currently selected item is a message or appointment, and can choose to appear in read or compose scenarios. Using add-in commands if possible is a [best practice](../concepts/add-in-development-best-practices.md).

## <a name="creating-the-add-in-command"></a>アドイン コマンドの作成

Add-in commands are declared in the add-in manifest in the [VersionOverrides element](../reference/manifest/versionoverrides.md). This element is an addition to the manifest schema v1.1 that ensures backward compatibility. In a client that doesn't support `VersionOverrides`, existing add-ins will continue to function as they did without add-in commands.

`VersionOverrides` マニフェスト エントリは、アドインについての多くの事柄 (ホスト、リボンに追加するコントロールの種類、テキスト、アイコン、関連する機能など) を指定します。

When an add-in needs to provide status updates, such as progress indicators or error messages, it must do so through the [notification APIs](/javascript/api/outlook/office.notificationmessages). The processing for the notifications must also be defined in a separate HTML file that is specified in the `FunctionFile` node of the manifest.

Developers should define icons for all required sizes so that the add-in commands will adjust smoothly along with the ribbon. The required icon sizes are 80 x 80 pixels, 32 x 32 pixels, and 16 x 16 pixels for desktop, and 48 x 48 pixels, 32 x 32 pixels, and 25 x 25 pixels for mobile.

## <a name="how-do-add-in-commands-appear"></a>アドイン コマンドの表示方法

An add-in command appears on the ribbon as a button. When a user installs an add-in, its commands appear in the UI as a group of buttons. This can either be on the ribbon's default tab or on a custom tab. For messages, the default is either the **Home** or **Message** tab. For the calendar, the default is the **Meeting**, **Meeting Occurrence**, **Meeting Series**, or **Appointment** tab. For module extensions, the default is a custom tab. On the default tab, each add-in can have one ribbon group with up to 6 commands. On custom tabs, the add-in can have up to 10 groups, each with 6 commands. Add-ins are limited to only one custom tab.

リボンがいっぱいになると、アドイン コマンドがオーバーフロー メニューに表示されます。 通常、アドインのアドイン コマンドはグループ化されています。

![リボンのアドイン コマンド ボタン](../images/commands-normal.png)

![リボンとオーバーフロー メニューのアドイン コマンド ボタン](../images/commands-collapsed.png)

When an add-in command is added to an add-in, the add-in name is removed from the app bar. Only the add-in command button on the ribbon remains.

### <a name="modern-outlook-on-the-web"></a>モダン Outlook on the web

Outlook on the web では、アドイン名はオーバーフロー メニューに表示されます。 アドインに複数のアドイン コマンドがある場合、アドイン メニューを展開して、アドイン名のラベルが付いたボタンのグループを表示できます。

![アドイン コマンド ボタンが見つかるオーバーフロー メニュー](../images/commands-overflow-menu-web.png)

![アドイン コマンド ボタンを表示しているオーバーフローメニュー](../images/commands-overflow-menu-expand-web.png)

## <a name="what-ux-shapes-exist-for-add-in-commands"></a>アドイン コマンドの UX シェイプの目的

The UX shape for an add-in command consists of a ribbon tab in the host application that contains buttons that can perform various functions. Currently, three UI shapes are supported:

- JavaScript 関数を実行するボタン
- 作業ウィンドウを起動するボタン
- 他の 2 種類のボタンについて 1 つ以上を選択肢とするドロップダウン メニューを表示するボタン

### <a name="executing-a-javascript-function"></a>JavaScript 関数の実行

Use an add-in command button that executes a JavaScript function for scenarios where the user doesn't need to make any additional selections to initiate the action. This can be for actions such as track, remind me, or print, or scenarios when the user wants more in-depth information from a service.

モジュール拡張機能では、メイン ユーザー インターフェイスのコンテンツを操作する JavaScript 関数をアドイン コマンド ボタンで実行できます。

![Outlook リボンの機能を実行するボタン。](../images/commands-uiless-button-1.png)

### <a name="launching-a-task-pane"></a>作業ウィンドウの起動

Use an add-in command button to launch a task pane for scenarios where a user needs to interact with an add-in for a longer period of time. For example, the add-in requires changes to settings or the completion of many fields.

The default width of the vertical task pane is 320 px. The vertical task pane can be resized in both the Outlook Explorer and inspector. The pane can be resized in the same way the to-do pane and list view resize.

![Outlook リボンの作業ウィンドウを開くボタン。](../images/commands-task-pane-button-1.png)

<br/>

This screenshot shows an example of a vertical task pane. The pane opens with the name of the add-in command in the top left corner. Users can use the **X** button in the upper-right corner of the pane to close the add-in when they are finished using it. By default, this pane will not persist across messages. Add-ins can [support pinning](pinnable-taskpane.md) for the task pane and receive events when a new message is selected. All UI elements rendered in the task pane, aside from the add-in name and the close button, are provided by the add-in.

If a user chooses another add-in command that opens a task pane, the task pane is replaced with the recently used command. If a user chooses an add-in command button that executes a function, or drop-down menu while the task pane is open, the action will be completed and the task pane will remain open.

### <a name="drop-down-menu"></a>ドロップダウン メニュー

A drop-down menu add-in command defines a static list of buttons. The buttons within the menu can be any mix of buttons that execute a function or buttons that open a task pane. Submenus are not supported.

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
