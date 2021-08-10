---
title: Visual Studio Code 用 Microsoft Office アドイン デバッガー拡張機能
description: アドイン デバッガー Visual Studio Code拡張機能Microsoft Office使用して、アドインのOfficeデバッグします。
ms.date: 02/01/2021
localization_priority: Normal
ms.openlocfilehash: d027e5937fa3a58623ce9e798fc683e5459e73b8b72606c0a006e465c9c1360c
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57088469"
---
# <a name="microsoft-office-add-in-debugger-extension-for-visual-studio-code"></a>Visual Studio Code 用 Microsoft Office アドイン デバッガー拡張機能

Visual Studio Code の Microsoft Office アドイン デバッガー拡張機能を使用すると、Office アドインを元の webView (EdgeHTML) ランタイムで Microsoft Edge に対してデバッグできます。 WebView2 に対するデバッグMicrosoft Edge (Chromiumベース) については、この[記事を参照してください。](./debug-desktop-using-edge-chromium.md)

このデバッグ モードは動的で、コードの実行中にブレークポイントを設定できます。 デバッガーが接続されている間、コード内の変更をすぐに確認できます。すべてデバッグ セッションを失う必要はありません。 コードの変更も保持されます。そのため、コードに対する複数の変更の結果を確認できます。 次の図は、この拡張機能の動作を示しています。

![Officeアドイン デバッガー拡張機能は、アドインのセクションExcelデバッグします。](../images/vs-debugger-extension-for-office-addins.jpg)

## <a name="prerequisites"></a>前提条件

- [Visual Studio Code](https://code.visualstudio.com/) (管理者として実行する必要があります)
- [Node.js (バージョン 10 以上)](https://nodejs.org/)
- Windows 10
- [Microsoft Edge](https://www.microsoft.com/edge)

これらの手順では、コマンド ラインの使用経験、基本的な JavaScript の理解、および Yo Office ジェネレーターを使用する前に Office アドイン プロジェクトを作成したと仮定します。 前にこれを行ったことがない場合は、次のようなチュートリアルの 1 つを参照Excel Office[検討してください](../tutorials/excel-tutorial.md)。

## <a name="install-and-use-the-debugger"></a>デバッガーをインストールして使用する

1. アドイン プロジェクトを作成する必要がある場合は[、Yo Officeを使用して作成します](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator)。 コマンド ライン内のプロンプトに従って、プロジェクトをセットアップします。 ニーズに合わせて任意の言語または種類のプロジェクトを選択できます。

    > [!NOTE]
    > プロジェクトが既に存在する場合は、手順 1 をスキップして、手順 2 に進みます。

1. 管理者としてコマンド プロンプトを開きます。
   ![コマンド プロンプト のオプション ([管理者として実行] を含む) Windows 10。](../images/run-as-administrator-vs-code.jpg)

1. プロジェクト ディレクトリに移動します。

1. 次のコマンドを実行して、プロジェクトを管理者Visual Studio Code開きます。

    ```command&nbsp;line
    code .
    ```

  ファイルVisual Studio Code開いた後、手動でプロジェクト フォルダーに移動します。

  > [!TIP]
  > 管理者としてVisual Studio Codeを開く場合は、管理者として実行オプションを選択し、Visual Studio Codeで管理者を検索した後Windows。

1. VS Code で **Ctrl キー + Shift キー + X キー** を選択して、拡張機能バーを開きます。 "Microsoft Office アドイン デバッガー" 拡張機能を検索してインストールします。

1. プロジェクトの .vscode フォルダーで、**launch.json** ファイルを開きます。 セクションに次のコードを追加 `configurations` します。

    ```JSON
    {
      "type": "office-addin",
      "request": "attach",
      "name": "Attach to Office Add-ins",
      "port": 9222,
      "trace": "verbose",
      "url": "https://localhost:3000/taskpane.html?_host_Info=HOST$Win32$16.01$en-US$$$$0",
      "webRoot": "${workspaceFolder}",
      "timeout": 45000
    }
    ```

1. コピーした JSON のセクションで、"url" セクションを探します。 この URL では、大文字の HOST テキストを、アドインをホストしているアプリケーションに置き換Office必要があります。 たとえば、Office アドインが Excel 用の場合、URL 値は https://localhost:3000/taskpane.html?_host_Info= <strong>"Excel</strong>$Win 32$16.01$en-US$ \$ \$ \$ 0" になります。

1. コマンド プロンプトを開き、プロジェクトのルート フォルダーに移動します。 コマンドを実行 `npm start` して開発サーバーを起動します。 アドインがクライアントに読み込まれるOffice作業ウィンドウを開きます。

1. [デバッグ] **Visual Studio Codeし、[** デバッグの表示] を>、Ctrl + Shift **+ D** と入力してデバッグ ビューに切り替えます。

1. [デバッグ] オプションで、[**アドインに接続Office選択します**。**[F5]** を選択するか、メニューから **[デバッグ - >デバッグ** の開始] を選択してデバッグを開始します。

1. プロジェクトの作業ウィンドウ ファイルにブレークポイントを設定します。 コード行の横にホバー Visual Studio Code表示される赤い円を選択すると、ブレークポイントを設定できます。

    ![赤い円は、次のコード行にVisual Studio Code。](../images/set-breakpoint.jpg)

1. アドインを実行します。 ブレークポイントがヒットし、ローカル変数を検査できます。

## <a name="see-also"></a>関連項目

- [Office アドインのテストとデバッグ](test-debug-office-add-ins.md)

- [Windows 10 で開発者ツールを使用してアドインをデバッグする](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

- [Microsoft Edge WebView2 (Chromium ベース) を使用した Windows 上のアドインをデバッグする](debug-desktop-using-edge-chromium.md)
