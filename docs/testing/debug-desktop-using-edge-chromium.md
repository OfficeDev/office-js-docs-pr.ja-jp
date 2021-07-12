---
title: Microsoft Edge WebView2 (Chromium ベース) を使用した Windows 上のアドインをデバッグする
description: VS Code で拡張機能 Debugger for Microsoft Edge を使用し、Microsoft Edge WebView2 (Chromium ベース) を使用した Office アドインをデバッグする方法について説明します。
ms.date: 01/29/2021
localization_priority: Priority
ms.openlocfilehash: 6a62718147fbb5d2e8a6819066425737d853cbf0
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53350177"
---
# <a name="debug-add-ins-on-windows-using-edge-chromium-webview2"></a>Edge Chromium WebView2 を使用して Windows でアドインをデバッグする

Windows 上で動作する Office アドインは、VS Code の拡張機能 Debugger for Microsoft Edge を使用することで、Edge Chromium WebView2 ランタイムに対してデバッグを行うことができます。

## <a name="prerequisites"></a>前提条件

- [Visual Studio Code](https://code.visualstudio.com/) (管理者として実行する必要があります)
- [Node.js (バージョン 10 以上)](https://nodejs.org/)
- Windows 10
- [Microsoft Edge Chromium は Windows Insider に提供しています](https://www.microsoftedgeinsider.com/)

## <a name="install-and-use-the-debugger"></a>デバッガーをインストールして使用する

1. [Office アドイン用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)を使用してプロジェクトを作成してください。これを行うには、「[Outlook アドインのクイック スタート](../quickstarts/outlook-quickstart.md)」などのクイック スタート ガイドのいずれかをご利用ください。

    > [!TIP]
    > Yeoman ジェネレーター ベースのアドインを使用していない場合は、レジストリ キーを調整する必要があります。 プロジェクトのルート フォルダーで、コマンドラインを使用して以下を実行します: `office-add-in-debugging start <your manifest path>`。

1. VS Code でプロジェクトを開きます。 VS Code で **Ctrl キー + Shift キー + X キー** を選択して、拡張機能バーを開きます。 「Debugger for Microsoft Edge」で拡張機能を検索し、これをインストールします。

1. プロジェクトの **.vscode** フォルダーで、**launch.json** ファイルを開きます。 構成セクションに以下のコードを追加します。

      ```JSON
        {
          "name": "Debug Office Add-in (Edge Chromium)",
          "type": "edge",
          "request": "attach",
          "useWebView": "advanced",
          "port": 9229,
          "timeout": 600000,
          "webRoot": "${workspaceRoot}",
        },
      ```

1. 次に、**[表示]、[デバッグ]** の順に選択するか、**Ctrl キー + Shift キー + D キー** を入力してデバッグ ビューに切り替えます。

1. デバッグ オプションから、**Excel Desktop (Edge Chromium)** などのホスト アプリケーション用に Edge Chromium オプションを選択します。 **F5** キーを選択するか、メニューから **[デバッグ]、[デバッグの開始]** の順に選択してデバッグを開始します。

1. これで、Excel などのホスト アプリケーションでアドインを使用する準備ができました。 **[作業ウィンドウの表示]** を選択するか、その他のアドイン コマンドを実行します。 ダイアログ ボックスが表示され、以下が表示されます。

    > WebView は読み込み時に停止します。
    > WebView をデバッグするには、拡張機能 Microsoft Debugger for Edge を使用して VS Code を WebView のインスタンスにアタッチし、[OK] をクリックして続行します。 今後このダイアログが表示されないようにするには、[キャンセル] をクリックします。

    **[OK]** を選択します。

    > [!NOTE]
    > **[キャンセル]** を選択すると、このアドインのインスタンスの実行中はダイアログが表示されなくなります。 ただし、アドインを再起動すると、ダイアログはもう一度表示されます。

1. これで、プロジェクトのコードにブレークポイントを設定し、デバッグを実行できるようになりました。

## <a name="see-also"></a>関連項目

- [Office アドインのテストとデバッグ](test-debug-office-add-ins.md)
- [Visual Studio Code 用 Microsoft Office アドイン デバッガー拡張機能](debug-with-vs-extension.md)
- [作業ウィンドウからデバッガーをアタッチする](attach-debugger-from-task-pane.md)