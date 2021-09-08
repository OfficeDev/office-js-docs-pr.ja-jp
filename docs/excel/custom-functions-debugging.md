---
title: UI レスのカスタム関数のデバッグ
description: 作業ウィンドウを使用しないExcel関数をデバッグする方法について説明します。
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: 1ee0e6e88b3ada88749278740d68f76c4a7368f6
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938712"
---
# <a name="ui-less-custom-functions-debugging"></a>UI レスのカスタム関数のデバッグ

この記事では、作業ウィンドウまたは他のユーザー インターフェイス要素 (UI レスのカスタム関数) を使用しないカスタム関数のデバッグのみについて説明します。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

オンWindows:

- [ExcelデスクトップおよびVisual Studio Code (VS Code) デバッガー](#use-the-vs-code-debugger-for-excel-desktop)
- [Excel on the webとVS Codeデバッガー](#use-the-vs-code-debugger-for-excel-in-microsoft-edge)
- [Excel on the webブラウザー ツール](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [コマンド ライン](#use-the-command-line-tools-to-debug)

Mac の場合:

- [Excel on the webブラウザー ツール](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [コマンド ライン](#use-the-command-line-tools-to-debug)

> [!NOTE]
> わかりやすくするために、この記事では、Visual Studio Code を使用してタスクを編集、実行し、場合によってはデバッグ ビューを使用するコンテキストでのデバッグを示します。 別のエディターまたはコマンド ライン ツールを使用している場合は[](#commands-for-building-and-running-your-add-in)、この記事の最後にあるコマンド ラインの手順を参照してください。

## <a name="requirements"></a>要件

このデバッグ プロセスは、作業 **ウィンドウ** や他の UI 要素を使用しない UI レスのカスタム関数でのみ機能します。 UI レスのカスタム関数を作成するには、「Excel のカスタム関数を作成する」チュートリアルの手順に従い[、Office](../tutorials/excel-tutorial-create-custom-functions.md)アドイン用の[Yeoman](https://www.npmjs.com/package/generator-office)ジェネレーターによってインストールされている作業ウィンドウと UI 要素をすべて削除します。

このデバッグ プロセスは、共有ランタイムを使用するカスタム関数プロジェクトと [互換性がない点に注意してください](../develop/configure-your-add-in-to-use-a-shared-runtime.md)。

## <a name="use-the-vs-code-debugger-for-excel-desktop"></a>デスクトップにVS CodeデバッガーをExcelする

デスクトップ上VS Code UI レスカスタム関数をデバッグするには、Office Excelを使用します。

> [!NOTE]
> Mac のデスクトップ デバッグは使用できませんが、ブラウザー ツールとコマンド ラインを使用してデバッグ[Excel on the web)。](#use-the-command-line-tools-to-debug)

### <a name="run-your-add-in-from-vs-code"></a>アドインを実行するには、次のVS Code

1. カスタム関数ルート プロジェクト フォルダーを開きます[。VS Code。](https://code.visualstudio.com/)
1. [ **ターミナル の実行>タスクを選択し** 、ウォッチを入力または選択 **します**。 これにより、ファイルの変更が監視され、再構築されます。
1. [ **ターミナル の実行>タスクを選択し** 、Dev Server を **入力または選択します**。

### <a name="start-the-vs-code-debugger"></a>デバッガーのVS Codeする

1. [ **ファイルの表示>実行] を** 選択するか **、Ctrl + Shift + D** と入力してデバッグ ビューに切り替えます。
1. [実行] ドロップダウン メニューから、[デスクトップ **(Excel関数) を選択します**。
1. デバッグ **を開始するには、[F5]** を選択します ( **または>から** [デバッグの開始] を選択します。 新しいExcelブックが開き、アドインが既にサイドロードされ、すぐに使用できます。

### <a name="start-debugging"></a>デバッグを開始する

1. このVS Code、ソース コード スクリプト ファイル (functions.jsまたは **functions.ts) を開きます**。
2. [カスタム関数のソース](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) コードでブレークポイントを設定します。
3. ブックのExcel、カスタム関数を使用する数式を入力します。

この時点で、ブレークポイントを設定したコード行で実行が停止します。 これで、コードをステップ実行し、ウォッチを設定し、必要なデバッグVS Codeを使用できます。

## <a name="use-the-vs-code-debugger-for-excel-in-microsoft-edge"></a>アプリケーション内のVS CodeデバッガーをExcel使用Microsoft Edge

このコマンドを使用VS Code、UI レスのカスタム関数を、Excelブラウザー Microsoft Edgeできます。 この機能をVS CodeするにはMicrosoft Edge拡張機能用の[デバッガーをインストールMicrosoft Edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge)があります。

### <a name="run-your-add-in-from-vs-code"></a>アドインを実行するには、次のVS Code

1. カスタム関数ルート プロジェクト フォルダーを開きます[。VS Code。](https://code.visualstudio.com/)
2. [ **ターミナル の実行>タスクを選択し** 、ウォッチを入力または選択 **します**。 これにより、ファイルの変更が監視され、再構築されます。
3. [ **ターミナル の実行>タスクを選択し** 、Dev Server を **入力または選択します**。

### <a name="start-the-vs-code-debugger"></a>デバッガーのVS Codeする

1. [ **ファイルの表示>実行] を** 選択するか **、Ctrl + Shift + D** と入力してデバッグ ビューに切り替えます。
1. [デバッグ] オプションで、[オンライン] **Office (エッジ Chromium) を選択します**。
1. ブラウザー Excel開Microsoft Edge新しいブックを作成します。
1. リボン **で [共有** ] を選択し、この新しいブックの URL のリンクをコピーします。
1. デバッグ **を開始するには、[F5]** **(または>[** デバッグの開始] を選択します。 ドキュメントの URL を求めるプロンプトが表示されます。
1. ブックの URL に貼り付け、Enter キーを押します。

### <a name="sideload-your-add-in"></a>アドインのサイドロード

1. リボンの **[挿入**] タブを選択し、[アドイン] セクションで、[アドイン] Office **を選択します**。
2. [アドイン **Office]** ダイアログで **、[MY ADD-INS]** タブを選択し、[自分のアドインの管理] を選択し、[マイ アップロード] をクリック **します**。
  
    ![右上Officeの [アドインの管理] というドロップダウンが表示された [Office アドイン] ダイアログボックスと、その下に [アップロード マイ アドイン] というオプションが表示されます。](../images/office-add-ins-my-account.png)

3. **アドイン** マニフェスト ファイルを参照し、[次へ] を **アップロード。**
  
    ![[参照]、[アップロード]、[キャンセル] のボタンがある [アドインのアップロード] ダイアログ。](../images/upload-add-in.png)

### <a name="set-breakpoints"></a>ブレークポイントの設定

1. このVS Code、ソース コード スクリプト ファイル (functions.jsまたは **functions.ts) を開きます**。
2. [カスタム関数のソース](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) コードでブレークポイントを設定します。
3. ブックのExcel、カスタム関数を使用する数式を入力します。

## <a name="use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web"></a>ブラウザー開発者ツールを使用して、ブラウザーのカスタム関数をデバッグExcel on the web

ブラウザー開発者ツールを使用して、UI レスのカスタム関数をデバッグExcel on the web。 次の手順は、Windows macOS の両方で機能します。

### <a name="run-your-add-in-from-visual-studio-code"></a>アドインをアプリから実行Visual Studio Code

1. カスタム関数ルート プロジェクト フォルダーを Visual Studio Code [(VS Code) で開きます](https://code.visualstudio.com/)。
2. [ **ターミナル の実行>タスクを選択し** 、ウォッチを入力または選択 **します**。 これにより、ファイルの変更が監視され、再構築されます。
3. [ **ターミナル の実行>タスクを選択し** 、Dev Server を **入力または選択します**。

### <a name="sideload-your-add-in"></a>アドインのサイドロード

1. [ファイル[Office on the web] を開きます](https://office.live.com/)。
2. 新しいブックを開Excelします。
3. リボンの **[挿入**] タブを開き、[アドイン] セクションで、[アドイン] Office **を選択します**。
4. [アドイン **Office]** ダイアログで **、[MY ADD-INS]** タブを選択し、[自分のアドインの管理] を選択し、[マイ アップロード] をクリック **します**。
  
    ![右上Officeの [アドインの管理] というドロップダウンが表示された [Office アドイン] ダイアログボックスと、その下に [アップロード マイ アドイン] というオプションが表示されます。](../images/office-add-ins-my-account.png)

5. アドイン マニフェスト ファイルを **参照** して、**[アップロード]** を選択します。
  
    ![[参照]、[アップロード]、[キャンセル] のボタンがある [アドインのアップロード] ダイアログ。](../images/upload-add-in.png)

> [!NOTE]
> ドキュメントにサイドロードすると、ドキュメントを開くごとにサイドロードされたままです。

### <a name="start-debugging"></a>デバッグを開始する

1. ブラウザーで開発者ツールを開きます。 Chrome およびほとんどのブラウザーの場合、F12 は開発者ツールを開きます。
2. 開発者ツールで **、Cmd + P** または Ctrl + **P** (functions.jsまたは **functions.ts)** を **使用して** ソース コード スクリプト ファイルを開きます。
3. [カスタム関数のソース](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) コードでブレークポイントを設定します。 

コードを変更する必要がある場合は、編集を行い、VS Code保存できます。 ブラウザーを更新して、読み込まれた変更を確認します。

## <a name="use-the-command-line-tools-to-debug"></a>コマンド ライン ツールを使用してデバッグする

アドインを使用していないVS Code、コマンド ライン (bash、PowerShell など) を使用してアドインを実行できます。 ブラウザー開発者ツールを使用して、ブラウザーのコードをデバッグする必要Excel on the web。 コマンド ラインを使用してデスクトップ バージョンExcelデバッグすることはできません。

1. コマンド ラインから実行して `npm run watch` 、コードの変更が発生した場合の監視と再構築を行います。
2. 2 番目のコマンド ライン ウィンドウを開きます (最初のウィンドウはウォッチの実行中にブロックされます)。

3. デスクトップ バージョンのデスクトップ バージョンでアドインを起動する場合はExcelコマンドを実行します。
  
    `npm run start:desktop`
  
    または、アドインを次のコマンドで起動する場合Excel on the web実行します。
  
    `npm run start:web`
  
    このExcel on the webアドインをサイドロードする必要があります。 「アドインを [サイドロードする」の手順に従って](#sideload-your-add-in) 、アドインをサイドロードします。 次に、次のセクションに進み、デバッグを開始します。
  
4. ブラウザーで開発者ツールを開きます。 Chrome およびほとんどのブラウザーの場合、F12 は開発者ツールを開きます。
5. 開発者ツールで、ソース コード スクリプト ファイル(functions.jsまたは **functions.ts) を開きます**。 カスタム関数コードは、ファイルの末尾近くに位置している可能性があります。
6. カスタム関数のソース コードで、コード行を選択してブレークポイントを適用します。

コードを変更する必要がある場合は、編集を行い、Visual Studio保存できます。 ブラウザーを更新して、読み込まれた変更を確認します。

### <a name="commands-for-building-and-running-your-add-in"></a>アドインを構築および実行するコマンド

使用可能なビルド タスクは複数あります。

- `npm run watch`: 開発用のビルドと、ソース ファイルの保存時に自動的に再構築する
- `npm run build-dev`: 一度開発用にビルドする
- `npm run build`: 実稼働用のビルド
- `npm run dev-server`: 開発に使用する Web サーバーを実行します。

次のタスクを使用して、デスクトップまたはオンラインでデバッグを開始できます。

- `npm run start:desktop`: デスクトップExcelを開始し、アドインをサイドロードします。
- `npm run start:web`: アドインExcel on the webを開始し、サイドロードします。
- `npm run stop`: デバッグExcel停止します。

## <a name="next-steps"></a>次のステップ

UI レス [のカスタム関数の認証方法について説明します](custom-functions-authentication.md)。

## <a name="see-also"></a>関連項目

* [カスタム関数のトラブルシューティング](custom-functions-troubleshooting.md)
* [Excel のカスタム関数でのエラー処理 ](custom-functions-errors.md)
* [Excel でカスタム関数を作成する](custom-functions-overview.md)
