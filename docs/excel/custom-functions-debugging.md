---
title: UI レスカスタム関数のデバッグ
description: 作業ウィンドウを使用しない Excel カスタム関数をデバッグする方法について説明します。
ms.date: 05/31/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1268aa07d160970fda12f8fccbe88e0427b246fc
ms.sourcegitcommit: 81f6018ac9731ff73e36d30f5ff10df21504c093
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/04/2022
ms.locfileid: "65891929"
---
# <a name="ui-less-custom-functions-debugging"></a>UI レスカスタム関数のデバッグ

この記事では、作業ウィンドウまたはその他のユーザー インターフェイス要素 (UI レスカスタム関数) を使用しないカスタム関数 *のデバッグについて説明* します。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

Windows の場合:

- [Excel Desktop および Visual Studio Code (VS Code) デバッガー](#use-the-vs-code-debugger-for-excel-desktop)
- [Excel on the Web および VS Code デバッガー](#use-the-vs-code-debugger-for-excel-in-microsoft-edge)
- [Excel on the Web ツールとブラウザー ツール](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [コマンド ライン](#use-the-command-line-tools-to-debug)

Mac の場合:

- [Excel on the Web ツールとブラウザー ツール](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [コマンド ライン](#use-the-command-line-tools-to-debug)

> [!NOTE]
> わかりやすくするために、この記事では、Visual Studio Code を使用してタスクを編集、実行し、場合によってはデバッグ ビューを使用するコンテキストでのデバッグについて説明します。 別のエディターまたはコマンド ライン ツールを使用している場合は、この記事の最後にある [コマンド ラインの手順](#commands-for-building-and-running-your-add-in) を参照してください。

## <a name="requirements"></a>要件

このデバッグ プロセスは、作業ウィンドウやその他の UI 要素を使用しない UI レスのカスタム関数 **でのみ** 機能します。 UI レスのカスタム関数を作成するには、「 [Excel でカスタム関数を作成](../tutorials/excel-tutorial-create-custom-functions.md) する」チュートリアルの手順に従って、 [Office アドイン用 Yeoman ジェネレーター](../develop/yeoman-generator-overview.md)によってインストールされているすべての作業ウィンドウと UI 要素を削除します。

このデバッグ プロセスは、 [共有ランタイム](../develop/configure-your-add-in-to-use-a-shared-runtime.md)を使用するカスタム関数プロジェクトと互換性がありません。

このデバッグ プロセスは、Yeoman ジェネレーターの **マニフェストのみのオプションを含む Office アドイン プロジェクトで** 作成されたプロジェクトでは機能しません。 この記事の後半で説明するスクリプトは、そのオプションを使用してインストールされません。 このオプションで作成されたアドインをデバッグするには、必要に応じて、これらの記事のいずれかの手順を参照してください。 

- [Microsoft Edge (Chromium ベース)で開発者ツールを使用してアドインをデバッグする](../testing/debug-add-ins-using-devtools-edge-chromium.md)
- [Internet Explorer で開発者ツールを使用してアドインをデバッグする](../testing/debug-add-ins-using-f12-tools-ie.md)
- [Mac で Office アドインをデバッグする](../testing/debug-office-add-ins-on-ipad-and-mac.md)

## <a name="use-the-vs-code-debugger-for-excel-desktop"></a>Excel Desktop で VS Code デバッガーを使用する

VS Code を使用すると、デスクトップ上の Office Excel で UI レスのカスタム関数をデバッグできます。

> [!NOTE]
> Mac 用のデスクトップ デバッグは利用できませんが、 [ブラウザーツールとコマンド ラインを使用して Excel on the web をデバッグ](#use-the-command-line-tools-to-debug)できます)。

### <a name="run-your-add-in-from-vs-code"></a>VS Code からアドインを実行する

1. [VS Code](https://code.visualstudio.com/) でカスタム関数のルート プロジェクト フォルダーを開きます。
1. **[ターミナル >実行タスク**] **を選択し**、ウォッチを入力または選択します。 これにより、ファイルの変更が監視され、再構築されます。
1. **[ターミナル >実行タスク**] を選択し、**開発サーバー** を入力または選択します。

### <a name="start-the-vs-code-debugger"></a>VS Code デバッガーを開始する

1. **[ビュー>実行**] を選択するか **、Ctrl + Shift + D キー** を押してデバッグ ビューに切り替えます。
1. [実行] ドロップダウン メニューから、 **Excel Desktop (カスタム関数)** を選択します。
1. **F5** を選択します (または、メニューから **[Run -> Start Debugging**]\(デバッグの開始\) を選択してデバッグを開始します。 新しい Excel ブックが開き、アドインが既にサイドロードされ、使用できるようになります。

### <a name="start-debugging"></a>デバッグの開始

1. VS Code で、ソース コード スクリプト ファイル (**functions.js** または **functions.ts**) を開きます。
2. カスタム関数のソース コードに[ブレークポイントを設定](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints)します。
3. Excel ブックで、カスタム関数を使用する数式を入力します。

この時点で、ブレークポイントを設定したコード行で実行が停止します。 これで、コードをステップ実行し、ウォッチを設定し、必要な VS Code デバッグ機能を使用できます。

## <a name="use-the-vs-code-debugger-for-excel-in-microsoft-edge"></a>Microsoft Edge で Excel 用の VS Code デバッガーを使用する

VS Code を使用して、Microsoft Edge ブラウザーで Excel で UI レスカスタム関数をデバッグできます。 Microsoft Edge で VS Code を使用するには、 [Visual Studio Code 用の Microsoft Edge DevTools 拡張機能をインストールする](/microsoft-edge/visual-studio-code/microsoft-edge-devtools-extension)必要があります。

### <a name="run-your-add-in-from-vs-code"></a>VS Code からアドインを実行する

1. [VS Code](https://code.visualstudio.com/) でカスタム関数のルート プロジェクト フォルダーを開きます。
2. **[ターミナル >実行タスク**] **を選択し**、ウォッチを入力または選択します。 これにより、ファイルの変更が監視され、再構築されます。
3. **[ターミナル >実行タスク**] を選択し、**開発サーバー** を入力または選択します。

### <a name="start-the-vs-code-debugger"></a>VS Code デバッガーを開始する

1. **[ビュー>実行**] を選択するか **、Ctrl + Shift + D キー** を押してデバッグ ビューに切り替えます。
1. [デバッグ] オプションから、[ **Office Online (Edge Chromium)]** を選択します。
1. Microsoft Edge ブラウザーで Excel を開き、新しいブックを作成します。
1. リボンで **[共有** ] を選択し、この新しいブックの URL のリンクをコピーします。
1. **F5** を選択します (または、メニューから **[Run > Start Debugging**]\(デバッグの開始\) を選択して、デバッグを開始します。 ドキュメントの URL を求めるプロンプトが表示されます。
1. ブックの URL を貼り付けて Enter キーを押します。

### <a name="sideload-your-add-in"></a>アドインのサイドロード

1. リボンの [ **挿入** ] タブを選択し、[ **アドイン** ] セクションで [ **Office アドイン**] を選択します。
2. **[Office アドイン**] ダイアログで、[**MY ADD-INS**] タブを選択し、[**マイ アドインの管理**] を選択して、[**マイ アドインのアップロード**] を選択します。
  
    ![[Office アドイン] ダイアログで、右上にドロップダウンが [アドインの管理] と表示され、その下に [自分のアドインのアップロード] オプションが表示されます。](../images/office-add-ins-my-account.png)

3. アドイン マニフェスト ファイルを **参照** し、[**アップロード**] を選択します。
  
    ![[参照]、[アップロード]、[キャンセル] のボタンがある [アドインのアップロード] ダイアログ。](../images/upload-add-in.png)

### <a name="set-breakpoints"></a>ブレークポイントを設定する

1. VS Code で、ソース コード スクリプト ファイル (**functions.js** または **functions.ts**) を開きます。
2. カスタム関数のソース コードに[ブレークポイントを設定](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints)します。
3. Excel ブックで、カスタム関数を使用する数式を入力します。

## <a name="use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web"></a>ブラウザー開発者ツールを使用して Excel on the Web でカスタム関数をデバッグする

ブラウザー開発者ツールを使用して、Web 上の Excel で UI レスのカスタム関数をデバッグできます。 次の手順は、Windows と macOS の両方で機能します。

### <a name="run-your-add-in-from-visual-studio-code"></a>Visual Studio Code からアドインを実行する

1. [Visual Studio Code (VS Code)](https://code.visualstudio.com/) でカスタム関数のルート プロジェクト フォルダーを開きます。
2. **[ターミナル >実行タスク**] **を選択し**、ウォッチを入力または選択します。 これにより、ファイルの変更が監視され、再構築されます。
3. **[ターミナル >実行タスク**] を選択し、**開発サーバー** を入力または選択します。

### <a name="sideload-your-add-in"></a>アドインのサイドロード

1. [Web で Office を開きます](https://office.live.com/)。
2. 新しい Excel ブックを開きます。
3. リボンの [ **挿入** ] タブを開き、[アドイン] セクション **で** [ **Office アドイン**] を選択します。
4. **[Office アドイン**] ダイアログで、[**MY ADD-INS**] タブを選択し、[**マイ アドインの管理**] を選択して、[**マイ アドインのアップロード**] を選択します。
  
    ![[Office アドイン] ダイアログで、右上にドロップダウンが [アドインの管理] と表示され、その下に [自分のアドインのアップロード] オプションが表示されます。](../images/office-add-ins-my-account.png)

5. アドイン マニフェスト ファイルを **参照** して、**[アップロード]** を選択します。
  
    ![[参照]、[アップロード]、[キャンセル] のボタンがある [アドインのアップロード] ダイアログ。](../images/upload-add-in.png)

> [!NOTE]
> ドキュメントにサイドロードすると、ドキュメントを開くたびにサイドロードされたままになります。

### <a name="start-debugging"></a>デバッグの開始

1. ブラウザーで開発者ツールを開きます。 Chrome およびほとんどのブラウザーの場合、F12 は開発者ツールを開きます。
2. 開発者ツールで、 **Cmd + P** または **Ctrl + P** (**functions.js** または **functions.ts**) を使用してソース コード スクリプト ファイルを開きます。
3. カスタム関数のソース コードに[ブレークポイントを設定](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints)します。 

コードを変更する必要がある場合は、VS Code で編集を行い、変更を保存できます。 ブラウザーを更新して、読み込まれた変更を確認します。

## <a name="use-the-command-line-tools-to-debug"></a>コマンド ライン ツールを使用してデバッグする

VS Code を使用していない場合は、コマンド ライン (bash、PowerShell など) を使用してアドインを実行できます。 ブラウザー開発者ツールを使用して、Web 上の Excel でコードをデバッグする必要があります。 コマンド ラインを使用してデスクトップ バージョンの Excel をデバッグすることはできません。

1. コマンド ラインから、コードの変更が発生したときに監視と再構築を実行 `npm run watch` します。
2. 2 番目のコマンド ライン ウィンドウを開きます (最初のコマンド ライン ウィンドウは、ウォッチの実行中にブロックされます)。

3. デスクトップ バージョンの Excel でアドインを開始する場合は、次のコマンドを実行します。
  
    `npm run start:desktop`
  
    または、Web 上の Excel でアドインを開始する場合は、次のコマンドを実行します。
  
    `npm run start:web -- --document {url}` (ここで `{url}` 、OneDrive または SharePoint 上の Excel ファイルの URL)
  
    アドインがドキュメントにサイドロードされない場合は、「 [アドインをサイドロード](#sideload-your-add-in) する」の手順に従ってアドインをサイドロードします。 次に、次のセクションに進み、デバッグを開始します。
  
4. ブラウザーで開発者ツールを開きます。 Chrome およびほとんどのブラウザーの場合、F12 は開発者ツールを開きます。
5. 開発者ツールで、ソース コード スクリプト ファイル (**functions.js** または **functions.ts**) を開きます。 カスタム関数コードは、ファイルの末尾近くに配置できます。
6. カスタム関数のソース コードで、コード行を選択してブレークポイントを適用します。

コードを変更する必要がある場合は、Visual Studio で編集を行い、変更を保存できます。 ブラウザーを更新して、読み込まれた変更を確認します。

### <a name="commands-for-building-and-running-your-add-in"></a>アドインをビルドして実行するためのコマンド

使用可能なビルド タスクがいくつかあります。

- `npm run watch`: 開発用にビルドされ、ソース ファイルが保存されたときに自動的にリビルドされます
- `npm run build-dev`: 開発用に 1 回ビルドする
- `npm run build`: 運用環境用のビルド
- `npm run dev-server`: 開発に使用される Web サーバーを実行します

次のタスクを使用して、デスクトップまたはオンラインでデバッグを開始できます。

- `npm run start:desktop`: デスクトップ上で Excel を起動し、アドインをサイドロードします。
- `npm run start:web -- --document {url}` (OneDrive または SharePoint 上の Excel ファイルの URL を指定 `{url}` します): Web 上で Excel を起動し、アドインをサイドロードします。
- `npm run stop`: Excel とデバッグを停止します。

## <a name="next-steps"></a>次の手順

[UI レスカスタム関数の認証方法](custom-functions-authentication.md)について説明します。

## <a name="see-also"></a>関連項目

* [カスタム関数のトラブルシューティング](custom-functions-troubleshooting.md)
* [Excel のカスタム関数でのエラー処理 ](custom-functions-errors.md)
* [Excel でカスタム関数を作成する](custom-functions-overview.md)
