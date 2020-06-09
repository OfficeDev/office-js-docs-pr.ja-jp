---
title: Visual Studio Code の Microsoft Office アドインデバッガーの拡張機能
description: Office アドインをデバッグするには、Visual Studio Code extension Microsoft Office アドインデバッガーを使用します。
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: 1bd3814eba6da2339e7865d720b8a4c792b9310e
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611212"
---
# <a name="microsoft-office-add-in-debugger-extension-for-visual-studio-code"></a>Visual Studio Code の Microsoft Office アドインデバッガーの拡張機能

Visual Studio コード用の Microsoft Office アドインデバッガー拡張機能を使用すると、エッジランタイムに対して Office アドインをデバッグできます。

このデバッグモードは動的なので、コードの実行中にブレークポイントを設定できます。 デバッグセッションを失わずに、デバッガーがアタッチされている間は、コード内の変更をすぐに表示できます。 コードの変更も引き続き行われるため、コードに対する複数の変更の結果を確認できます。 次の図は、この拡張機能の動作を示しています。

![Office Addin デバッガー拡張機能 Excel アドインのセクションをデバッグする](../images/vs-debugger-extension-for-office-addins.jpg)

## <a name="prerequisites"></a>前提条件

- [Visual Studio Code](https://code.visualstudio.com/) (管理者として実行する必要があります)
- [Node.js (バージョン10以降)](https://nodejs.org/)
- Windows 10
- [Microsoft Edge](https://www.microsoft.com/edge)

これらの手順では、コマンドラインを使用して基本的な JavaScript を理解し、Yo Office ジェネレーターを使用する前に Office アドインプロジェクトを作成していることを前提としています。 これを実行していない場合は、この[Excel Office アドインのチュートリアル](../tutorials/excel-tutorial.md)のように、チュートリアルの1つにアクセスすることを検討してください。

## <a name="install-and-use-the-debugger"></a>デバッガーをインストールして使用する

1. アドインプロジェクトを作成する必要がある場合は、 [Yo Office ジェネレーターを使用して](https://docs.microsoft.com/office/dev/add-ins/quickstarts/excel-quickstart-jquery?tabs=yeomangenerator)プロジェクトを作成します。 コマンドライン内のプロンプトに従って、プロジェクトを設定します。 必要に応じて、任意の言語やプロジェクトの種類を選択できます。

> !ことプロジェクトが既に存在する場合は、手順1をスキップし、手順2に進みます。

2. 管理者としてコマンドプロンプトを開きます。
   ![Windows 10 の "管理者として実行" を含むコマンドプロンプトオプション](../images/run-as-administrator-vs-code.jpg)

3. プロジェクトディレクトリに移動します。

4. 次のコマンドを実行して、Visual Studio Code で管理者としてプロジェクトを開きます。

```command&nbsp;line
code .
```

Visual Studio コードが開いたら、プロジェクトフォルダーに手動で移動します。

> [!TIP]
> Visual Studio Code を管理者として開くには、Visual Studio Code を Windows で検索した後、そのコードを開くときに [**管理者として実行**] オプションを選択します。

5. VS コード内で、 **CTRL + SHIFT + X**を選択して [拡張バー] を開きます。 「Microsoft Office アドインデバッガー」拡張機能を検索してインストールします。

6. プロジェクトの vscode フォルダーで、**起動. json**ファイルを開きます。 次のコードをセクションに追加し `configurations` ます。

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

7. 先ほどコピーした JSON のセクションで、[url] セクションを見つけます。 この URL では、大文字のホストテキストを Office アドインのホストアプリケーションに置き換える必要があります。 たとえば、Office アドインが excel 用の場合、URL の値は " https://localhost:3000/taskpane.html?_host_Info= <strong>excel</strong>$Win 32 $ 16.01 $ en-us $ \$ \$ \$ 0" になります。

8. コマンドプロンプトを開き、自分がプロジェクトのルートフォルダーにあることを確認します。 コマンドを実行し `npm start` て、開発サーバーを起動します。 アドインが Office クライアントに読み込まれたら、作業ウィンドウを開きます。

9. Visual Studio Code に戻り、[**表示] > [デバッグ**] を選択するか、 **CTRL + SHIFT + D キー**を押してデバッグビューに切り替えます。

10. デバッグオプションで、[ **Office アドインにアタッチ**] を選択します。**F5 キーを押す**か、メニューからデバッグを**開始し >** デバッグを開始してデバッグを開始します。

11. プロジェクトの作業ウィンドウファイルにブレークポイントを設定します。 VS コードでブレークポイントを設定するには、コード行の横にあるカーソルを使用して、表示される赤い円を選択します。

![VS Code のコード行に赤い円が表示される](../images/set-breakpoint.jpg)

12. アドインを実行します。 ブレークポイントにヒットしたことが表示され、ローカル変数を調べることができます。

## <a name="see-also"></a>関連項目

* [Office アドインのテストとデバッグ](test-debug-office-add-ins.md)

* [Windows 10 で開発者ツールを使用してアドインをデバッグする](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

* [作業ウィンドウからデバッガーをアタッチする](attach-debugger-from-task-pane.md)
