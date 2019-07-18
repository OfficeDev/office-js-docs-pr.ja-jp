---
ms.date: 07/10/2019
description: Excel でカスタム関数をデバッグします。
title: カスタム関数のデバッグ
localization_priority: Normal
ms.openlocfilehash: 987df4fc638b94b7a5002c99aee6e36642f4e4a4
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771457"
---
# <a name="custom-functions-debugging"></a>カスタム関数のデバッグ

カスタム関数のデバッグは、使用しているプラットフォームによっては複数の方法で実行できます。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Windows の場合:
- [Excel デスクトップと Visual Studio Code (VS コード) デバッガー](#use-the-vs-code-debugger-for-excel-desktop)
- [Excel on the web および VS コードデバッガー](#use-the-vs-code-debugger-for-excel-in-microsoft-edge)
- [Excel on the web およびブラウザーツール](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [コマンドライン](#use-the-command-line-tools-to-debug)

On Mac:
- [Excel on the web およびブラウザーツール](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [コマンドライン](#use-the-command-line-tools-to-debug)

> [!NOTE]
> 簡単にするために、この記事では、Visual Studio Code を使用した編集、タスクの実行、および場合によってはデバッグビューを使用するためのデバッグについて説明します。 別のエディターまたはコマンドラインツールを使用している場合は、この記事の最後にある[コマンドラインの手順](#commands-for-building-and-running-your-add-in)を参照してください。

## <a name="requirements"></a>要件

デバッグを開始する前に、 [Office アドイン用の [ごみ箱] ジェネレーター](https://github.com/OfficeDev/generator-office)を使用して、カスタム関数プロジェクトを作成する必要があります。 カスタム関数プロジェクトを作成する方法のガイダンスについては、「[カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)」を参照してください。

## <a name="use-the-vs-code-debugger-for-excel-desktop"></a>Excel デスクトップ用の VS コードデバッガーを使用する

VS コードを使用して、デスクトップ上の Office Excel でカスタム関数をデバッグできます。

> [!NOTE]
> Mac 用のデスクトップデバッグは使用できませんが、[ブラウザーツールおよびコマンドラインを使用して、web 上で Excel をデバッグすることによって](#use-the-command-line-tools-to-debug)実現できます。

### <a name="run-your-add-in-from-vs-code"></a>VS コードからアドインを実行する

1. [VS Code](https://code.visualstudio.com/)でカスタム関数ルートプロジェクトフォルダーを開きます。
2. [**ターミナル > タスクの実行**] を選択して、[**ウォッチ**] を入力または選択します。 これにより、ファイルの変更が監視され、再構築されます。
3. [**ターミナル > タスクの実行**] を選択し、[**開発サーバー**] を入力または選択します。

### <a name="start-the-vs-code-debugger"></a>VS コードデバッガーを開始する

4. [**表示 > デバッグ**] を選択するか、 **Ctrl + Shift + D キー**を押してデバッグビューに切り替えます。
5. デバッグオプションで、[ **Excel デスクトップ**] を選択します。
6. **F5 キーを押し**て (または、[デバッグ] **-> メニューからデバッグ開始**)、デバッグを開始します。 アドインが既にサイドロードで使用できる状態で、新しい Excel ブックが開きます。

### <a name="start-debugging"></a>デバッグを開始する

1. VS Code で、ソースコードスクリプトファイル (**node.js**または**関数 ts**) を開きます。
2. カスタム関数のソースコードに[ブレークポイントを設定](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints)します。
3. Excel ブックで、カスタム関数を使用する数式を入力します。

この時点で、ブレークポイントを設定したコード行では、この時点で実行が停止します。 コードをステップ実行し、ウォッチポイントを設定して、必要な VS コードデバッグ機能を使用できるようになりました。

## <a name="use-the-vs-code-debugger-for-excel-in-microsoft-edge"></a>Microsoft Edge で Excel の VS コードデバッガーを使用する

VS コードを使用して、Microsoft Edge ブラウザー上の Excel でカスタム関数をデバッグできます。 Microsoft Edge で VS コードを使用するには、 [Microsoft edge 拡張機能用のデバッガー](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge)をインストールする必要があります。

### <a name="run-your-add-in-from-vs-code"></a>VS コードからアドインを実行する

1. [VS Code](https://code.visualstudio.com/)でカスタム関数ルートプロジェクトフォルダーを開きます。
2. [**ターミナル > タスクの実行**] を選択して、[**ウォッチ**] を入力または選択します。 これにより、ファイルの変更が監視され、再構築されます。
3. [**ターミナル > タスクの実行**] を選択し、[**開発サーバー**] を入力または選択します。

### <a name="start-the-vs-code-debugger"></a>VS コードデバッガーを開始する

4. [**表示 > デバッグ**] を選択するか、 **Ctrl + Shift + D キー**を押してデバッグビューに切り替えます。
5. [デバッグオプション] で、[ **Office Online (Microsoft Edge)**] を選択します。
6. Microsoft Edge ブラウザーで Excel を開き、新しいブックを作成します。
7. リボンの [**共有**] を選択し、この新しいブックの URL のリンクをコピーします。
8. **F5 キーを押し**ます (または、[ **> デバッグ**] を選択して、メニューからデバッグを開始します)。デバッグを開始します。 ドキュメントの URL の入力を求めるプロンプトが表示されます。
9. ブックの URL を貼り付け、Enter キーを押します。

### <a name="sideload-your-add-in"></a>アドインのサイドロード

1. リボンの [**挿入**] タブを選択し、 **** [アドイン] セクションで、[ **Office アドイン**] を選択します。
2. **[Office アドイン]** ダイアログ ボックスで、**[個人用アドイン]** タブ、**[個人用アドインの管理]**、**[個人用アドインのアップロード]** の順に選択します。
    
    ![右上に [個人用アドインの管理] というドロップダウンがあり、その下に [マイ アドインのアップロード] オプションのドロップダウンがある [Office アドイン] ダイアログ](../images/office-add-ins-my-account.png)

3.  アドインマニフェストファイルを**参照**し、[**アップロード**] を選択します。
    
    ![[参照]、[アップロード]、[キャンセル] のボタンがある [アドインのアップロード] ダイアログ。](../images/upload-add-in.png)


### <a name="set-breakpoints"></a>ブレークポイントを設定する
1. VS Code で、ソースコードスクリプトファイル (**node.js**または**関数 ts**) を開きます。
2. カスタム関数のソースコードに[ブレークポイントを設定](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints)します。
3. Excel ブックで、カスタム関数を使用する数式を入力します。

## <a name="use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web"></a>ブラウザー開発者ツールを使用して、web 上の Excel でカスタム関数をデバッグする

ブラウザー開発者ツールを使用して、web 上の Excel でカスタム関数をデバッグできます。 次の手順は、Windows と macOS の両方で動作します。

### <a name="run-your-add-in-from-visual-studio-code"></a>Visual Studio Code からアドインを実行する

1. カスタム関数のルートプロジェクトフォルダーを[Visual Studio Code (VS コード)](https://code.visualstudio.com/)で開きます。
2. [**ターミナル > タスクの実行**] を選択して、[**ウォッチ**] を入力または選択します。 これにより、ファイルの変更が監視され、再構築されます。
3. [**ターミナル > タスクの実行**] を選択し、[**開発サーバー**] を入力または選択します。

### <a name="sideload-your-add-in"></a>アドインのサイドロード

1. [Web 上の Microsoft Office を](https://office.live.com/)開きます。
2. 新しい Excel ブックを開きます。
3. リボンの  **[挿入]** タブを開き、 **[アドイン]** セクションで、 **Office [アドイン]** を選択します。
4. **[Office アドイン]** ダイアログ ボックスで、**[個人用アドイン]** タブ、**[個人用アドインの管理]**、**[個人用アドインのアップロード]** の順に選択します。
    
    ![右上に [個人用アドインの管理] というドロップダウンがあり、その下に [マイ アドインのアップロード] オプションのドロップダウンがある [Office アドイン] ダイアログ](../images/office-add-ins-my-account.png)

5.  アドイン マニフェスト ファイルを**参照**して、**[アップロード]** を選択します。
    
    ![[参照]、[アップロード]、[キャンセル] のボタンがある [アドインのアップロード] ダイアログ。](../images/upload-add-in.png)

> [!NOTE]
> サイドロードしたドキュメントは、ドキュメントを開くたびにサイドロードされたままになります。

### <a name="start-debugging"></a>デバッグを開始する

1. 開発者ツールをブラウザーで開きます。 Chrome およびほとんどのブラウザー F12 では、開発者ツールが開きます。
2. 開発者ツールで、 **Cmd + p**または**Ctrl + p** (**node.js**または**functions**) を使用してソースコードスクリプトファイルを開きます。
3. カスタム関数のソースコードに[ブレークポイントを設定](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints)します。 

コードを変更する必要がある場合は、VS コードで編集を行って変更を保存することができます。 ブラウザーを更新して、変更が読み込まれたことを確認します。

## <a name="use-the-command-line-tools-to-debug"></a>コマンドラインツールを使用してデバッグする

VS コードを使用していない場合は、コマンドライン (bash、PowerShell など) を使用してアドインを実行できます。 Web 上の Excel でコードをデバッグするには、ブラウザー開発者ツールを使用する必要があります。 コマンドラインを使用して、デスクトップ版の Excel をデバッグすることはできません。

1. コマンドラインからを実行`npm run watch`すると、コードの変更が発生したときにを監視し、再構築します。
2. 2番目のコマンドラインウィンドウを開きます (最初のウィンドウは、ウォッチの実行中にブロックされます)。

3. Excel のデスクトップバージョンでアドインを起動するには、次のコマンドを実行します。
    
    `npm run start:desktop`
    
    または、web 上の Excel でアドインを開始する場合は、次のコマンドを実行します。
    
    `npm run start:web`
    
    Excel on the web では、アドインをサイドロードする必要もあります。 「[サイドロード](#sideload-your-add-in)を使用してアドインをサイドロードする」の手順に従います。 その後、次のセクションに進み、デバッグを開始します。
    
4. 開発者ツールをブラウザーで開きます。 Chrome およびほとんどのブラウザー F12 では、開発者ツールが開きます。
5. [開発者ツール] で、ソースコードスクリプトファイル (**node.js**または**関数 ts**) を開きます。 カスタム関数のコードは、ファイルの末尾付近に配置されている場合があります。
6. カスタム関数のソースコードで、コードの行を選択してブレークポイントを適用します。

コードを変更する必要がある場合は、Visual Studio で編集を行って変更を保存することができます。 ブラウザーを更新して、変更が読み込まれたことを確認します。

### <a name="commands-for-building-and-running-your-add-in"></a>アドインをビルドして実行するためのコマンド

使用可能なビルドタスクはいくつかあります。
- `npm run watch`: ソースファイルの保存時に開発用のビルドを作成し、自動的に再構築します。
- `npm run build-dev`: 開発用ビルド
- `npm run build`: 運用のためのビルド
- `npm run dev-server`: 開発に使用する web サーバーを実行します。

次のタスクを使用して、デスクトップまたはオンラインでデバッグを開始できます。
- `npm run start:desktop`: デスクトップ上で Excel を起動し、アドインを読み込みます。
- `npm run start:web`: Web 上で Excel を起動し、アドインを読み込みます。
- `npm run stop`: Excel およびデバッグを停止します。

## <a name="next-steps"></a>次のステップ
[カスタム関数の認証方法](custom-functions-authentication.md)について説明します。 または、[カスタム関数の一意のアーキテクチャ](custom-functions-architecture.md)を確認します。

## <a name="see-also"></a>関連項目

* [カスタム関数のトラブルシューティング](custom-functions-troubleshooting.md)
* [カスタム関数をXLLユーザー定義関数と互換性のあるものにします](make-custom-functions-compatible-with-xll-udf.md)。
* [Excel でカスタム関数を作成する](custom-functions-overview.md)
