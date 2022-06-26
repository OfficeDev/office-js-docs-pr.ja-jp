---
title: カスタム関数のデバッグ
description: 共有ランタイムを使用しないExcelカスタム関数をデバッグする方法について説明します。
ms.date: 06/15/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1c53d73a0356d4f5f9af9bebbb6c34b99dbeb395
ms.sourcegitcommit: d8fbe472b35c758753e5d2e4b905a5973e4f7b52
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2022
ms.locfileid: "66229688"
---
# <a name="custom-functions-debugging"></a>カスタム関数のデバッグ

この記事では、**[共有ランタイム](../develop/configure-your-add-in-to-use-a-shared-runtime.md)を使用しない** カスタム関数のデバッグについて説明します。 共有ランタイムを使用するカスタム関数アドインをデバッグするには、「共有 [JavaScript ランタイムを使用するようにOffice アドインを構成する:デバッグ](../develop/configure-your-add-in-to-use-a-shared-runtime.md#debug)」を参照してください。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

## <a name="requirements"></a>要件

このデバッグ プロセスは、 **共有ランタイムを使用しない** カスタム関数に対してのみ機能します。 共有ランタイムを使用しないカスタム関数は、Office アドイン [用の Yeoman ジェネレーター](../develop/yeoman-generator-overview.md)を使用して作成されたExcel **カスタム関数アドイン プロジェクト** です。

このデバッグ プロセスは、Yeoman ジェネレーターの **マニフェストのみのオプションを含むOffice アドイン プロジェクトで** 作成されたプロジェクトでは機能しません。 この記事の後半で説明するスクリプトは、そのオプションを使用してインストールされません。 このオプションで作成されたアドインをデバッグするには、必要に応じて、これらの記事のいずれかの手順を参照してください。

- [Microsoft Edge (Chromium ベース)で開発者ツールを使用してアドインをデバッグする](../testing/debug-add-ins-using-devtools-edge-chromium.md)
- [Internet Explorer で開発者ツールを使用してアドインをデバッグする](../testing/debug-add-ins-using-f12-tools-ie.md)
- [Mac で Office アドインをデバッグする](../testing/debug-office-add-ins-on-ipad-and-mac.md)

次のアンカー リンクを使用して、デバッグ シナリオに関連するこの記事のセクションにアクセスします。

Windows:

- [Excel Desktop and Visual Studio Code (VS Code) デバッガー](#use-the-vs-code-debugger-for-excel-desktop)
- [デバッガーのExcel on the webとVS Code](#use-the-vs-code-debugger-for-excel-in-microsoft-edge)
- [Excel on the webツールとブラウザー ツール](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [コマンド ライン](#use-the-command-line-tools-to-debug)

Mac の場合:

- [Excel on the webツールとブラウザー ツール](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [コマンド ライン](#use-the-command-line-tools-to-debug)

> [!NOTE]
> わかりやすくするために、この記事では、Visual Studio Code を使用してタスクを編集、実行し、場合によってはデバッグ ビューを使用するコンテキストでのデバッグについて説明します。 別のエディターまたはコマンド ライン ツールを使用している場合は、この記事の最後にある [コマンド ラインの手順](#commands-for-building-and-running-your-add-in) を参照してください。

## <a name="use-the-vs-code-debugger-for-excel-desktop"></a>Excel Desktop でVS Code デバッガーを使用する

VS Codeを使用して、デスクトップ上のOffice Excelで共有ランタイムを使用しないカスタム関数をデバッグできます。

> [!IMPORTANT]
> 次のデバッグ手順には既知の問題があります。 この手順は、スクリプトの種類として **TypeScript** が選択された Yeoman ジェネレーターの **Excel Custom Functions アドイン プロジェクト** オプションを使用してインストールされたプロジェクトでは機能しますが、スクリプトの種類として **JavaScript** がインストールされているプロジェクトでは、この手順は機能しません。 詳細については、 [OfficeDev/office-js-docs-pr の問題 #3355](https://github.com/OfficeDev/office-js-docs-pr/issues/3355) を参照してください。

> [!NOTE]
> Mac 用のデスクトップ デバッグは利用できませんが、[ブラウザーツールとコマンド ラインを使用してExcel on the webをデバッグ](#use-the-command-line-tools-to-debug)できます。

### <a name="run-your-add-in-from-vs-code"></a>VS Codeからアドインを実行する

1. VS Codeでカスタム関数のルート プロジェクト フォルダー[を](https://code.visualstudio.com/)開きます。
1. **[ターミナル >実行タスク**] **を選択し**、ウォッチを入力または選択します。 これにより、ファイルの変更が監視され、再構築されます。
1. **[ターミナル >実行タスク**] を選択し、**開発サーバー** を入力または選択します。

### <a name="start-the-vs-code-debugger"></a>VS Code デバッガーを起動する

1. **[ビュー>実行**] を選択するか **、Ctrl + Shift + D キー** を押してデバッグ ビューに切り替えます。
1. **[実行とデバッグ**] ドロップダウン メニューから、[**デスクトップ (カスタム関数)] Excel** 選択します。

    :::image type="content" source="../images/custom-functions-run-and-debug-menu.jpg" alt-text="[実行とデバッグ] ドロップダウン メニュー Excelデスクトップ (カスタム関数) を示すスクリーンショット。":::

1. **F5** を選択します (または、メニューから **[Run -> Start Debugging**]\(デバッグの開始\) を選択してデバッグを開始します。 新しいExcel ブックが開き、アドインが既にサイドロードされ、使用できるようになります。

### <a name="start-debugging"></a>デバッグの開始

1. VS Codeで、ソース コード スクリプト ファイル (**functions.js** または **functions.ts**) を開きます。
2. カスタム関数のソース コードに[ブレークポイントを設定](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints)します。
3. Excel ブックに、カスタム関数を使用する数式を入力します。

この時点で、ブレークポイントを設定したコード行で実行が停止します。 これで、コードをステップ実行し、ウォッチを設定し、必要なVS Codeデバッグ機能を使用できます。

## <a name="use-the-vs-code-debugger-for-excel-in-microsoft-edge"></a>Microsoft EdgeでExcelにVS Code デバッガーを使用する

VS Codeを使用すると、Microsoft Edge ブラウザーのExcelで共有ランタイムを使用しないカスタム関数をデバッグできます。 Microsoft EdgeでVS Codeを使用するには、[Visual Studio Code 用の Microsoft Edge DevTools 拡張機能をインストールする](/microsoft-edge/visual-studio-code/microsoft-edge-devtools-extension)必要があります。

### <a name="run-your-add-in-from-vs-code"></a>VS Codeからアドインを実行する

1. VS Codeでカスタム関数のルート プロジェクト フォルダー[を](https://code.visualstudio.com/)開きます。
1. **[ターミナル >実行タスク**] **を選択し**、ウォッチを入力または選択します。 これにより、ファイルの変更が監視され、再構築されます。
1. **[ターミナル >実行タスク**] を選択し、**開発サーバー** を入力または選択します。

### <a name="start-the-vs-code-debugger"></a>VS Code デバッガーを起動する

1. **[ビュー>実行**] を選択するか **、Ctrl + Shift + D キー** を押してデバッグ ビューに切り替えます。
1. [デバッグ] オプションから[**Office オンライン (エッジ Chromium)]** を選択します。
1. Microsoft Edge ブラウザーでExcelを開き、新しいブックを作成します。
1. リボンで **[共有** ] を選択し、この新しいブックの URL のリンクをコピーします。
1. **F5** を選択します (または、メニューから **[Run > Start Debugging**]\(デバッグの開始\) を選択して、デバッグを開始します。 ドキュメントの URL を求めるプロンプトが表示されます。
1. ブックの URL を貼り付けて Enter キーを押します。

### <a name="sideload-your-add-in"></a>アドインのサイドロード

1. リボンの [**挿入**] タブを選択し、[**アドイン**] セクションで [**アドインOffice選択します**。
2. **[Office アドイン**] ダイアログで、[**MY アドイン**] タブを選択し、[**マイ アドインの管理**] を選択して、[**マイ アドイン] をアップロード** します。
  
    ![Office アドイン ダイアログで、右上にドロップダウンが [アドインの管理] と表示され、その下に [自分のアドインをアップロード] オプションが表示されます。](../images/office-add-ins-my-account.png)

3. アドイン マニフェスト ファイルを **参照** し、**アップロード** を選択します。
  
    ![[参照]、[アップロード]、[キャンセル] のボタンがある [アドインのアップロード] ダイアログ。](../images/upload-add-in.png)

### <a name="set-breakpoints"></a>ブレークポイントを設定する

1. VS Codeで、ソース コード スクリプト ファイル (**functions.js** または **functions.ts**) を開きます。
2. カスタム関数のソース コードに[ブレークポイントを設定](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints)します。
3. Excel ブックに、カスタム関数を使用する数式を入力します。

## <a name="use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web"></a>ブラウザー開発者ツールを使用して、Excel on the webでカスタム関数をデバッグする

ブラウザー開発者ツールを使用すると、Excel on the webで共有ランタイムを使用しないカスタム関数をデバッグできます。 次の手順は、WindowsとmacOSの両方で機能します。

### <a name="run-your-add-in-from-visual-studio-code"></a>Visual Studio Code からアドインを実行する

1. [Visual Studio Code (VS Code)](https://code.visualstudio.com/) でカスタム関数のルート プロジェクト フォルダーを開きます。
2. **[ターミナル >実行タスク**] **を選択し**、ウォッチを入力または選択します。 これにより、ファイルの変更が監視され、再構築されます。
3. **[ターミナル >実行タスク**] を選択し、**開発サーバー** を入力または選択します。

### <a name="sideload-your-add-in"></a>アドインのサイドロード

1. [Office on the web](https://office.live.com/)を開きます。
2. 新しいExcel ブックを開きます。
3. リボンの [**挿入**] タブを開き、[アドイン] セクション **で** **[アドインOffice** 選択します。
4. **[Office アドイン**] ダイアログで、[**MY アドイン**] タブを選択し、[**マイ アドインの管理**] を選択して、[**マイ アドイン] をアップロード** します。
  
    ![Office アドイン ダイアログで、右上にドロップダウンが [アドインの管理] と表示され、その下に [自分のアドインをアップロード] オプションが表示されます。](../images/office-add-ins-my-account.png)

5. アドイン マニフェスト ファイルを **参照** して、**[アップロード]** を選択します。
  
    ![[参照]、[アップロード]、[キャンセル] のボタンがある [アドインのアップロード] ダイアログ。](../images/upload-add-in.png)

> [!NOTE]
> ドキュメントにサイドロードすると、ドキュメントを開くたびにサイドロードされたままになります。

### <a name="start-debugging"></a>デバッグの開始

1. ブラウザーで開発者ツールを開きます。 Chrome およびほとんどのブラウザーの場合、F12 は開発者ツールを開きます。
2. 開発者ツールで、 **Cmd + P** または **Ctrl + P** (**functions.js** または **functions.ts**) を使用してソース コード スクリプト ファイルを開きます。
3. カスタム関数のソース コードに[ブレークポイントを設定](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints)します。 

コードを変更する必要がある場合は、VS Codeで編集を行い、変更を保存できます。 ブラウザーを更新して、読み込まれた変更を確認します。

## <a name="use-the-command-line-tools-to-debug"></a>コマンド ライン ツールを使用してデバッグする

VS Codeを使用していない場合は、コマンド ライン (bash、PowerShell など) を使用してアドインを実行できます。 ブラウザー開発者ツールを使用して、Excel on the webでコードをデバッグする必要があります。 コマンド ラインを使用してデスクトップ バージョンのExcelをデバッグすることはできません。

1. コマンド ラインから、コードの変更が発生したときに監視と再構築を実行 `npm run watch` します。
2. 2 番目のコマンド ライン ウィンドウを開きます (最初のコマンド ライン ウィンドウは、ウォッチの実行中にブロックされます)。

3. デスクトップ バージョンのExcelでアドインを起動する場合は、次のコマンドを実行します。
  
    `npm run start:desktop`
  
    または、アドインを起動する場合は、次のコマンドを実行Excel on the web。
  
    `npm run start:web -- --document {url}`(ここで`{url}`、OneDriveまたはSharePoint上のExcel ファイルの URL)
  
    アドインがドキュメントにサイドロードされない場合は、「 [アドインをサイドロード](#sideload-your-add-in) する」の手順に従ってアドインをサイドロードします。 次に、次のセクションに進み、デバッグを開始します。
  
4. ブラウザーで開発者ツールを開きます。 Chrome およびほとんどのブラウザーの場合、F12 は開発者ツールを開きます。
5. 開発者ツールで、ソース コード スクリプト ファイル (**functions.js** または **functions.ts**) を開きます。 カスタム関数コードは、ファイルの末尾近くに配置できます。
6. カスタム関数のソース コードで、コード行を選択してブレークポイントを適用します。

コードを変更する必要がある場合は、Visual Studioで編集を行い、変更を保存できます。 ブラウザーを更新して、読み込まれた変更を確認します。

### <a name="commands-for-building-and-running-your-add-in"></a>アドインをビルドして実行するためのコマンド

使用可能なビルド タスクがいくつかあります。

- `npm run watch`: 開発用にビルドされ、ソース ファイルが保存されたときに自動的にリビルドされます
- `npm run build-dev`: 開発用に 1 回ビルドする
- `npm run build`: 運用環境用のビルド
- `npm run dev-server`: 開発に使用される Web サーバーを実行します

次のタスクを使用して、デスクトップまたはオンラインでデバッグを開始できます。

- `npm run start:desktop`: デスクトップでExcelを開始し、アドインをサイドロードします。
- `npm run start:web -- --document {url}`(ここで`{url}`、OneDriveまたはSharePoint上のExcel ファイルの URL): Excel on the webを開始し、アドインをサイドロードします。
- `npm run stop`: Excelとデバッグを停止します。

## <a name="next-steps"></a>次のステップ

[共有ランタイムを使用しないカスタム関数の認証](custom-functions-authentication.md)について説明します。

## <a name="see-also"></a>関連項目

* [カスタム関数のトラブルシューティング](custom-functions-troubleshooting.md)
* [Excel のカスタム関数でのエラー処理 ](custom-functions-errors.md)
* [Excel でカスタム関数を作成する](custom-functions-overview.md)
