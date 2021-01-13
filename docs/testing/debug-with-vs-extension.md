---
title: Visual Studio Code 用 Microsoft Office アドイン デバッガー拡張機能
description: アドイン デバッガー Visual Studioコード拡張機能Microsoft Office使用して、アドインのOfficeデバッグします。
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: 83791d5d60238288e3059809b8b8c02b1f4f768f
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/13/2021
ms.locfileid: "49840112"
---
# <a name="microsoft-office-add-in-debugger-extension-for-visual-studio-code"></a>Visual Studio Code 用 Microsoft Office アドイン デバッガー拡張機能

コードMicrosoft Officeアドイン デバッガー拡張機能Visual Studioでは、Edge ランタイムに対して Office アドインをデバッグできます。

このデバッグ モードは動的で、コードの実行中にブレークポイントを設定できます。 デバッガーがアタッチされている間は、デバッグ セッションを失わずに、コードの変更をすぐに確認できます。 コードの変更も保持されます。そのため、コードに対する複数の変更の結果を確認できます。 次の図は、この拡張機能の動作を示しています。

![Officeアドインのセクションをデバッグするアドイン デバッガー拡張機能](../images/vs-debugger-extension-for-office-addins.jpg)

## <a name="prerequisites"></a>前提条件

- [Visual Studioコード](https://code.visualstudio.com/) (管理者として実行する必要があります)
- [Node.js (バージョン 10 以上)](https://nodejs.org/)
- Windows 10
- [Microsoft Edge](https://www.microsoft.com/edge)

これらの手順は、コマンド ラインの使用経験、基本的な JavaScript の理解、Yo Office ジェネレーターを使用する前に Office アドイン プロジェクトを作成した経験を前提にしています。 まだこれを行っていない場合は、次の [Excel](../tutorials/excel-tutorial.md)やアドインのチュートリアルOfficeチュートリアルのいずれかを参照してください。

## <a name="install-and-use-the-debugger"></a>デバッガーをインストールして使用する

1. アドイン プロジェクトを作成する必要がある場合は、Yo Office ジェネレーターを使用 [して作成します](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator)。 コマンド ライン内のプロンプトに従って、プロジェクトを設定します。 必要に応じて、任意の言語または種類のプロジェクトを選択できます。

> [!NOTE]
> プロジェクトが既に存在する場合は、手順 1 をスキップして手順 2 に進みます。

2. 管理者としてコマンド プロンプトを開きます。
   ![Windows 10 のコマンド プロンプト オプション ("管理者として実行" を含む)](../images/run-as-administrator-vs-code.jpg)

3. プロジェクト ディレクトリに移動します。

4. 次のコマンドを実行して、管理者として Visual Studioコードでプロジェクトを開きます。

```command&nbsp;line
code .
```

コードVisual Studio開いた後、手動でプロジェクト フォルダーに移動します。

> [!TIP]
> 管理者として Visual Studio コードを開く場合は、Windowsでコードを検索した後、Visual Studio コードを開く際に管理者として実行オプションを選択します。

5. VS Code 内で **Ctrl + Shift + X** キーを押して拡張機能バーを開きます。 "Microsoft Office アドイン デバッガー" 拡張機能を検索してインストールします。

6. プロジェクトの .vscode フォルダーで、ファイルlaunch.js **開** きます。 セクションに次のコードを追加 `configurations` します。

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

7. コピーした JSON のセクションで、"url" セクションを探します。 この URL では、大文字の HOST テキストを、アドインをホストしているアプリケーションに置Officeがあります。 たとえば、Office アドインが Excel 用の場合、URL 値は https://localhost:3000/taskpane.html?_host_Info= <strong>"Excel</strong>$Win 32$16.01$en-US$ \$ \$ \$ 0" になります。

8. コマンド プロンプトを開き、プロジェクトのルート フォルダーに移動します。 コマンドを実行 `npm start` して開発サーバーを起動します。 アドインがクライアントに読み込Office作業ウィンドウを開きます。

9. コードにVisual Studioし、[デバッグ] の[>] を選択するか **、Ctrl + Shift + D** キーを押してデバッグ ビューに切り替えます。

10. [デバッグ] オプションで、[アドイン **にアタッチOffice選択します**。F5 **キーを押** するか、[デバッグ] **->メニュー** から [デバッグの開始] を選択してデバッグを開始します。

11. プロジェクトの作業ウィンドウ ファイルにブレークポイントを設定します。 VS Code でブレークポイントを設定するには、コード行の横にカーソルを合わせると、表示される赤い円を選択します。

![VS Code のコード行に赤い円が表示される](../images/set-breakpoint.jpg)

12. アドインを実行します。 ブレークポイントにヒットしたと表示され、ローカル変数を検査できます。

## <a name="see-also"></a>関連項目

* [Office アドインのテストとデバッグ](test-debug-office-add-ins.md)

* [Windows 10 で開発者ツールを使用してアドインをデバッグする](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

* [作業ウィンドウからデバッガーをアタッチする](attach-debugger-from-task-pane.md)