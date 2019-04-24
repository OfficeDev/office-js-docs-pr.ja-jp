---
ms.date: 03/13/2019
description: Excel でカスタム関数をデバッグします。
title: カスタム関数のデバッグ (プレビュー)
localization_priority: Normal
ms.openlocfilehash: 08563ef630ebc457219c4c622328b84d13e6acab
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32448760"
---
# <a name="custom-functions-debugging-preview"></a>カスタム関数のデバッグ (プレビュー)

カスタム関数のデバッグは、使用しているプラットフォームによっては複数の方法で実行できます。

Windows の場合:
- [Excel デスクトップと Visual Studio code (VS コード) デバッガー](#use-the-vs-code-debugger-for-excel-desktop)
- [Excel Online および VS コードデバッガー](#use-the-vs-code-debugger-for-excel-online-in-microsoft-edge)
- [Excel Online およびブラウザーツール](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-online)
- [コマンドライン](#use-the-command-line-tools-to-debug)

On Mac:
- [Excel Online およびブラウザーツール](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-online)
- [コマンドライン](#use-the-command-line-tools-to-debug)

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

> [!NOTE]
> 簡単にするために、この記事では、Visual Studio Code を使用した編集、タスクの実行、および場合によってはデバッグビューを使用するためのデバッグについて説明します。 別のエディターまたはコマンドラインツールを使用している場合は、この記事の最後にある[コマンドラインの手順](#use-the-command-line-tools-to-debug)を参照してください。

## <a name="requirements"></a>要件

デバッグを開始する前に、Yo Office ジェネレーターを使用してカスタム関数アドインプロジェクトを作成し、プロジェクトに対して信頼できる自己署名証明書があることを確認する必要があります。 プロジェクトを作成する手順については、「[カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)」を参照してください。 証明書の信頼の手順については、「[自己署名証明書を信頼できるルート証明書として追加する](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)」を参照してください。

## <a name="use-the-vs-code-debugger-for-excel-desktop"></a>Excel デスクトップ用の VS コードデバッガーを使用する

VS コードを使用して、デスクトップ上の Office Excel でカスタム関数をデバッグできます。

> [!NOTE]
> Mac のデスクトップデバッグは利用できませんが[、ブラウザーツールを使用して Excel Online をデバッグする](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-online)ことで実現できます。

### <a name="run-your-add-in-from-vs-code"></a>VS コードからアドインを実行する

1. [VS Code](https://code.visualstudio.com/)でカスタム関数ルートプロジェクトフォルダーを開きます。
2. [**ターミナル > 実行タスク**] を選択し、[**ウォッチ**] を入力または選択します。 これにより、ファイルの変更が監視され、再構築されます。
3. [**ターミナル > 実行タスク**] を選択し、[**開発サーバー**] を入力または選択します。 

### <a name="start-the-vs-code-debugger"></a>VS コードデバッガーを開始する

4. [ **view > Debug** ] を選択するか、 **Ctrl + Shift + D**を入力してデバッグビューに切り替えます。
5. デバッグオプションで、[ **Excel デスクトップ**] を選択します。
6. **F5 キー**を選択するか、またはメニューからデバッグ**開始 >** を選択してデバッグを開始します。 アドインが既にサイドロードで使用できる状態で、新しい Excel ブックが開きます。

### <a name="start-debugging"></a>デバッグを開始する

1. VS Code で、ソースコードスクリプトファイル (node.js または関数 ts) を開きます。
2. カスタム関数のソースコードに[ブレークポイントを設定](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints)します。
3. Excel ブックで、カスタム関数を使用する数式を入力します。

この時点で、ブレークポイントを設定したコード行では、この時点で実行が停止します。 コードをステップ実行し、ウォッチポイントを設定して、必要な VS コードデバッグ機能を使用できるようになりました。

## <a name="use-the-vs-code-debugger-for-excel-online-in-microsoft-edge"></a>Microsoft Edge で Excel Online 用の VS コードデバッガーを使用する

Microsoft Edge ブラウザーで excel Online のカスタム関数をデバッグするには、VS コードを使用できます。 microsoft edge で VS コードを使用するには、 [microsoft edge 拡張機能用のデバッガー](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge)をインストールする必要があります。

### <a name="run-your-add-in-from-vs-code"></a>VS コードからアドインを実行する

1. [VS Code](https://code.visualstudio.com/)でカスタム関数ルートプロジェクトフォルダーを開きます。
2. [**ターミナル > 実行タスク**] を選択し、[**ウォッチ**] を入力または選択します。 これにより、ファイルの変更が監視され、再構築されます。
3. [**ターミナル > 実行タスク**] を選択し、[**開発サーバー**] を入力または選択します。 

### <a name="start-the-vs-code-debugger"></a>VS コードデバッガーを開始する

4. [ **view > Debug** ] を選択するか、 **Ctrl + Shift + D**を入力してデバッグビューに切り替えます。
5. デバッグオプションで、[ **Office Online (エッジ)**] を選択します。
6. Microsoft Edge ブラウザーを使用して excel online を開き、excel online を開き、新しいブックを作成します。
7. リボンの [**共有**] を選択し、この新しいブックの URL のリンクをコピーします。
8. **F5 キー**を選択するか、[**デバッグ > 開始**] メニューからデバッグを開始してデバッグを開始します。 ドキュメントの URL の入力を求めるプロンプトが表示されます。
9. ブックの URL を貼り付け、enter キーを押します。

### <a name="sideload-your-add-in"></a>アドインのサイドロード   

1. リボンの [**挿入**] タブを選択し、 **** [アドイン] セクションで、[ **Office アドイン**] を選択します。
2. **[Office アドイン]** ダイアログ ボックスで、**[個人用アドイン]** タブ、**[個人用アドインの管理]**、**[個人用アドインのアップロード]** の順に選択します。
    
    ![右上に [個人用アドインの管理] というドロップダウンがあり、その下に [マイ アドインのアップロード] オプションのドロップダウンがある [Office アドイン] ダイアログ](../images/office-add-ins-my-account.png)

3.  アドインマニフェストファイルを**参照**し、[**アップロード**] を選択します。
    
    ![[参照]、[アップロード]、[キャンセル] のボタンがある [アドインのアップロード] ダイアログ。](../images/upload-add-in.png)


### <a name="set-breakpoints"></a>ブレークポイントを設定する
1. VS Code で、ソースコードスクリプトファイル (node.js または関数 ts) を開きます。
2. カスタム関数のソースコードに[ブレークポイントを設定](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints)します。
3. Excel ブックで、カスタム関数を使用する数式を入力します。

## <a name="use-the-browser-developer-tools-to-debug-custom-functions-in-excel-online"></a>ブラウザー開発者ツールを使用して Excel Online のカスタム関数をデバッグする

ブラウザー開発者ツールを使用して、Excel Online のカスタム関数をデバッグできます。 次の手順は、Windows と macOS の両方で動作します。

### <a name="run-your-add-in-from-visual-studio-code"></a>Visual Studio Code からアドインを実行する

1. カスタム関数のルートプロジェクトフォルダーを[Visual Studio Code (VS コード)](https://code.visualstudio.com/)で開きます。
2. [**ターミナル > 実行タスク**] を選択し、[**ウォッチ**] を入力または選択します。 これにより、ファイルの変更が監視され、再構築されます。
3. [**ターミナル > 実行タスク**] を選択し、[**開発サーバー**] を入力または選択します。 

### <a name="sideload-your-add-in"></a>アドインのサイドロード   

1. [Microsoft Office Online](https://office.live.com/) を開きます。
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
2. 開発者ツールで、 **Cmd + p**または**Ctrl + p** (node.js または functions) を使用してソースコードスクリプトファイルを開きます。
3. カスタム関数のソースコードに[ブレークポイントを設定](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints)します。 

コードを変更する必要がある場合は、VS コードで編集を行って変更を保存することができます。 ブラウザーを更新して、変更が読み込まれたことを確認します。

## <a name="use-the-command-line-tools-to-debug"></a>コマンドラインツールを使用してデバッグする

VS コードを使用していない場合は、コマンドライン (bash、PowerShell など) を使用してアドインを実行できます。 Excel Online でコードをデバッグするには、ブラウザー開発者ツールを使用する必要があります。 コマンドラインを使用して、デスクトップ版の Excel をデバッグすることはできません。

1. コマンドラインからを実行`npm run watch`すると、コードの変更が発生したときにを監視し、再構築します。
2. 2番目のコマンドラインウィンドウを開きます (最初のウィンドウは、ウォッチの実行中にブロックされます)。

3. Excel のデスクトップバージョンでアドインを起動するには、次のコマンドを実行します。
    
    `npm run start desktop`
    
    または、Excel Online でアドインを起動したい場合は、次のコマンドを実行します。
    
    `npm run start web`
    
    Excel Online の場合は、アドインをサイドロードする必要もあります。 「[サイドロード](#sideload-your-add-in)を使用してアドインをサイドロードする」の手順に従います。 その後、次のセクションに進み、デバッグを開始します。
    
4. 開発者ツールをブラウザーで開きます。 Chrome およびほとんどのブラウザー F12 では、開発者ツールが開きます。
5. [開発者ツール] で、ソースコードスクリプトファイル (node.js または関数 ts) を開きます。 カスタム関数のコードは、ファイルの末尾付近に配置されている場合があります。
6. カスタム関数のソースコードで、コードの行を選択してブレークポイントを適用します。

コードを変更する必要がある場合は、Visual Studio で編集を行って変更を保存することができます。 ブラウザーを更新して、変更が読み込まれたことを確認します。

### <a name="commands-for-building-and-running-your-add-in"></a>アドインをビルドして実行するためのコマンド

使用可能なビルドタスクはいくつかあります。
- `npm run watch`: ソースファイルの保存時に開発用のビルドを作成し、自動的に再構築します。
- `npm run build-dev`: 開発用ビルド
- `npm run build`: 運用のためのビルド
- `npm run dev-server`: 開発に使用する web サーバーを実行します。

次のタスクを使用して、デスクトップまたはオンラインでデバッグを開始できます。
- `npm run start desktop`: デスクトップ上で Excel を起動し、アドインを読み込みます。
- `npm run start web`: Excel Online を起動して、アドインを読み込みます。
- `npm run stop`: Excel およびデバッグを停止します。

## <a name="see-also"></a>関連項目

* [カスタム関数のメタデータ](custom-functions-json.md)
* [Excel カスタム関数のランタイム](custom-functions-runtime.md)
* [カスタム関数のベスト プラクティス](custom-functions-best-practices.md)
* [カスタム関数の変更ログ](custom-functions-changelog.md)
* [Excel カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)
