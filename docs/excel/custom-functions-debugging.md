---
ms.date: 07/10/2020
description: 作業ウィンドウを使用しない Excel カスタム関数をデバッグする方法について説明します。
title: UI レスのカスタム関数のデバッグ
localization_priority: Normal
ms.openlocfilehash: 00065a465a22f83891dfb207943102b079e96a0f
ms.sourcegitcommit: 7482ab6bc258d98acb9ba9b35c7dd3b5cc5bed21
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/24/2021
ms.locfileid: "51178077"
---
# <a name="ui-less-custom-functions-debugging"></a>UI レスのカスタム関数のデバッグ

作業ウィンドウまたは他のユーザー インターフェイス要素 (UI レスのカスタム関数) を使用しないカスタム関数のデバッグは、使用しているプラットフォームに応じて複数の方法で実行できます。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

Windows の場合:
- [Excel Desktop および Visual Studio コード (VS Code) デバッガー](#use-the-vs-code-debugger-for-excel-desktop)
- [Excel on the web and VS Code Debugger](#use-the-vs-code-debugger-for-excel-in-microsoft-edge)
- [Excel on the web and browser tools](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [コマンド ライン](#use-the-command-line-tools-to-debug)

Mac の場合:
- [Excel on the web and browser tools](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [コマンド ライン](#use-the-command-line-tools-to-debug)

> [!NOTE]
> わかりやすくするために、この記事では、Visual Studio Code を使用してタスクを編集、実行し、場合によってはデバッグ ビューを使用するコンテキストでのデバッグを示します。 別のエディターまたはコマンド ライン ツールを使用している場合は[](#commands-for-building-and-running-your-add-in)、この記事の最後にあるコマンド ラインの手順を参照してください。

## <a name="requirements"></a>要件

デバッグを開始する前に [、Yeoman](https://github.com/OfficeDev/generator-office) ジェネレーターを使用Officeカスタム関数プロジェクトを作成する必要があります。 カスタム関数プロジェクトを作成する方法のガイダンスについては、カスタム関数の [チュートリアルを参照してください](../tutorials/excel-tutorial-create-custom-functions.md)。

## <a name="use-the-vs-code-debugger-for-excel-desktop"></a>Excel Desktop で VS Code デバッガーを使用する

VS Code を使用すると、デスクトップ上の Excel で UI レスのOfficeをデバッグできます。

> [!NOTE]
> Mac のデスクトップ デバッグは使用できませんが、ブラウザー ツールとコマンド ラインを使用して Web 上の [Excel をデバッグできます](#use-the-command-line-tools-to-debug))。

### <a name="run-your-add-in-from-vs-code"></a>VS Code からアドインを実行する

1. VS Code でカスタム関数ルート プロジェクト フォルダー [を開きます](https://code.visualstudio.com/)。
2. [ **ターミナル の実行>タスクを選択し** 、ウォッチを入力または選択 **します**。 これにより、ファイルの変更が監視され、再構築されます。
3. [ **ターミナル の実行>タスクを選択し** 、Dev Server を **入力または選択します**。

### <a name="start-the-vs-code-debugger"></a>VS Code デバッガーを起動する

4. [ **ファイルの表示>実行] を** 選択するか **、Ctrl + Shift + D** と入力してデバッグ ビューに切り替えます。
5. [実行] ドロップダウン メニューから **、[Excel デスクトップ ] (エッジ クロム) を選択します**。
6. デバッグ **を開始するには、[F5]** を選択します ( **または>から** [デバッグの開始] を選択します。 アドインが既にサイドロードされ、すぐに使用できる状態で、新しい Excel ブックが開きます。

### <a name="start-debugging"></a>デバッグを開始する

1. VS Code で、ソース コード スクリプトファイル (functions.js **または functions.ts) を開きます**。
2. [カスタム関数のソース](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) コードでブレークポイントを設定します。
3. Excel ブックに、カスタム関数を使用する数式を入力します。

この時点で、ブレークポイントを設定したコード行で実行が停止します。 これで、コードをステップ実行し、ウォッチを設定し、必要な VS Code デバッグ機能を使用できます。

## <a name="use-the-vs-code-debugger-for-excel-in-microsoft-edge"></a>Microsoft Edge で Excel の VS Code デバッガーを使用する

VS Code を使用すると、Microsoft Edge ブラウザーの Excel で UI レスのカスタム関数をデバッグできます。 Microsoft Edge で VS Code を使用するには、デバッガー for [Microsoft Edge 拡張機能をインストールする必要](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge) があります。

### <a name="run-your-add-in-from-vs-code"></a>VS Code からアドインを実行する

1. VS Code でカスタム関数ルート プロジェクト フォルダー [を開きます](https://code.visualstudio.com/)。
2. [ **ターミナル の実行>タスクを選択し** 、ウォッチを入力または選択 **します**。 これにより、ファイルの変更が監視され、再構築されます。
3. [ **ターミナル の実行>タスクを選択し** 、Dev Server を **入力または選択します**。

### <a name="start-the-vs-code-debugger"></a>VS Code デバッガーを起動する

4. [ **ファイルの表示>実行] を** 選択するか **、Ctrl + Shift + D** と入力してデバッグ ビューに切り替えます。
5. [デバッグ] オプションで、[オンライン] **Office (Edge Chromium) を選択します**。
6. Microsoft Edge ブラウザーで Excel を開き、新しいブックを作成します。
7. リボン **で [共有** ] を選択し、この新しいブックの URL のリンクをコピーします。
8. デバッグ **を開始するには、[F5]** **(または>[** デバッグの開始] を選択します。 ドキュメントの URL を求めるプロンプトが表示されます。
9. ブックの URL に貼り付け、Enter キーを押します。

### <a name="sideload-your-add-in"></a>アドインのサイドロード

1. リボンの **[挿入**] タブを選択し、[アドイン] セクションで、[アドイン] Office **選択します**。
2. [アドイン **Office]** ダイアログで **、[MY ADD-INS]** タブを選択し、[マイアドインの管理] を選択し、[自分のアドインのアップロード]**を選択します**。
    
    ![右上に [個人用アドインの管理] というドロップダウンがあり、その下に [マイ アドインのアップロード] オプションのドロップダウンがある [Office アドイン] ダイアログ](../images/office-add-ins-my-account.png)

3. **アドイン** マニフェスト ファイルを参照し、[アップロード] を **選択します**。
    
    ![[参照]、[アップロード]、[キャンセル] のボタンがある [アドインのアップロード] ダイアログ。](../images/upload-add-in.png)


### <a name="set-breakpoints"></a>ブレークポイントの設定
1. VS Code で、ソース コード スクリプトファイル (functions.js **または functions.ts) を開きます**。
2. [カスタム関数のソース](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) コードでブレークポイントを設定します。
3. Excel ブックに、カスタム関数を使用する数式を入力します。

## <a name="use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web"></a>ブラウザー開発者ツールを使用して、Web 上の Excel でカスタム関数をデバッグする

ブラウザー開発者ツールを使用して、Web 上の Excel で UI レスのカスタム関数をデバッグできます。 次の手順は、Windows と macOS の両方で機能します。

### <a name="run-your-add-in-from-visual-studio-code"></a>コードからアドインをVisual Studioする

1. カスタム関数ルート プロジェクト フォルダーを [コード] [(VS Code) Visual Studio開きます](https://code.visualstudio.com/)。
2. [ **ターミナル の実行>タスクを選択し** 、ウォッチを入力または選択 **します**。 これにより、ファイルの変更が監視され、再構築されます。
3. [ **ターミナル の実行>タスクを選択し** 、Dev Server を **入力または選択します**。

### <a name="sideload-your-add-in"></a>アドインのサイドロード

1. Web [Officeを開きます](https://office.live.com/)。
2. 新しい Excel ブックを開きます。
3. リボンの **[挿入**] タブを開き、[アドイン] セクションで、[アドイン] Office **選択します**。
4. [アドイン **Office]** ダイアログで **、[MY ADD-INS]** タブを選択し、[マイアドインの管理] を選択し、[自分のアドインのアップロード]**を選択します**。
    
    ![右上に [個人用アドインの管理] というドロップダウンがあり、その下に [マイ アドインのアップロード] オプションのドロップダウンがある [Office アドイン] ダイアログ](../images/office-add-ins-my-account.png)

5. アドイン マニフェスト ファイルを **参照** して、**[アップロード]** を選択します。
    
    ![[参照]、[アップロード]、[キャンセル] のボタンがある [アドインのアップロード] ダイアログ。](../images/upload-add-in.png)

> [!NOTE]
> ドキュメントにサイドロードすると、ドキュメントを開くごとにサイドロードされたままです。

### <a name="start-debugging"></a>デバッグを開始する

1. ブラウザーで開発者ツールを開きます。 Chrome およびほとんどのブラウザーの場合、F12 は開発者ツールを開きます。
2. 開発者ツールで **、Cmd + P** または Ctrl + **P** (functions.jsまたは **functions.ts)** を **使用して** ソース コード スクリプト ファイルを開きます。
3. [カスタム関数のソース](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) コードでブレークポイントを設定します。 

コードを変更する必要がある場合は、VS Code で編集を行い、変更を保存できます。 ブラウザーを更新して、読み込まれた変更を確認します。

## <a name="use-the-command-line-tools-to-debug"></a>コマンド ライン ツールを使用してデバッグする

VS Code を使用していない場合は、コマンド ライン (bash、PowerShell など) を使用してアドインを実行できます。 ブラウザー開発者ツールを使用して、Web 上の Excel でコードをデバッグする必要があります。 コマンド ラインを使用してデスクトップ バージョンの Excel をデバッグすることはできません。

1. コマンド ラインから実行して `npm run watch` 、コードの変更が発生した場合の監視と再構築を行います。
2. 2 番目のコマンド ライン ウィンドウを開きます (最初のウィンドウはウォッチの実行中にブロックされます)。

3. デスクトップ バージョンの Excel でアドインを起動する場合は、次のコマンドを実行します。
    
    `npm run start:desktop`
    
    または、Web 上の Excel でアドインを起動する場合は、次のコマンドを実行します。
    
    `npm run start:web`
    
    Web 上の Excel の場合は、アドインをサイドロードする必要があります。 「アドインを [サイドロードする」の手順に従って](#sideload-your-add-in) 、アドインをサイドロードします。 次に、次のセクションに進み、デバッグを開始します。
    
4. ブラウザーで開発者ツールを開きます。 Chrome およびほとんどのブラウザーの場合、F12 は開発者ツールを開きます。
5. 開発者ツールで、ソース コード スクリプト ファイル(functions.jsまたは **functions.ts) を開きます**。 カスタム関数コードは、ファイルの末尾近くに位置している可能性があります。
6. カスタム関数のソース コードで、コード行を選択してブレークポイントを適用します。

コードを変更する必要がある場合は、編集を行い、Visual Studio保存できます。 ブラウザーを更新して、読み込まれた変更を確認します。

### <a name="commands-for-building-and-running-your-add-in"></a>アドインを構築および実行するコマンド

使用できるビルド タスクは複数あります。
- `npm run watch`: 開発用のビルドと、ソース ファイルの保存時に自動的に再構築する
- `npm run build-dev`: 一度開発用にビルドする
- `npm run build`: 実稼働用のビルド
- `npm run dev-server`: 開発に使用する Web サーバーを実行します。

次のタスクを使用して、デスクトップまたはオンラインでデバッグを開始できます。
- `npm run start:desktop`: デスクトップ上で Excel を起動し、アドインをサイドロードします。
- `npm run start:web`: Web 上で Excel を起動し、アドインをサイドロードします。
- `npm run stop`: Excel とデバッグを停止します。

## <a name="next-steps"></a>次の手順
UI レス [のカスタム関数の認証方法について説明します](custom-functions-authentication.md)。

## <a name="see-also"></a>関連項目

* [カスタム関数のトラブルシューティング](custom-functions-troubleshooting.md)
* [Excel のカスタム関数でのエラー処理 ](custom-functions-errors.md)
* [Excel でカスタム関数を作成する](custom-functions-overview.md)
