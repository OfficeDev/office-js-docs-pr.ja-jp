---
title: 共有ランタイムを使用するように Office アドインを構成する
description: 共有ランタイムを使用して、追加のリボン、作業ウィンドウ、およびカスタム関数機能をサポートするように Office アドインを構成します。
ms.date: 07/18/2022
ms.prod: non-product-specific
ms.localizationpriority: high
ms.openlocfilehash: e6b10cc2d342d95a8542146ecbd95d750322421f
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/24/2022
ms.locfileid: "67422937"
---
# <a name="configure-your-office-add-in-to-use-a-shared-runtime"></a>共有ランタイムを使用するように Office アドインを構成する

[!include[Shared runtime requirements](../includes/shared-runtime-requirements-note.md)]

Office アドインを構成して、そのすべてのコードを 1 つの [共有ランタイム](../testing/runtimes.md#shared-runtime)で実行できます。 これにより、アドイン間での調整が容易になり、アドインのすべての部分から DOM や CORS にアクセスできます。 また、ドキュメントを開いたときにコードを実行したり、リボン ボタンを有効または無効にするなどの追加機能も有効にできます。 共有ランタイムが使用できるようにアドインを構成するには、この記事の手順に従います。

## <a name="create-the-add-in-project"></a>アドイン プロジェクトの作成

新しいプロジェクトを開始する場合は、[Office アドイン用の Yeoman ジェネレーター](yeoman-generator-overview.md)を使用して、Excel、PowerPoint、または Word アドイン プロジェクトを作成します。

コマンド `yo office --projectType taskpane --name "my office add in" --host <host> --js true` を実行します。ここで、`<host>` は次のいずれかの値です。

- Excel
- PowerPoint
- Word

> [!IMPORTANT]
> `--name` の引数値は、スペースがない場合でも二重引用符で囲む必要があります。

**--projecttype**、**--name** と **--js** コマンド ライン オプションのためのさまざまなオプションを使用できます。 オプションの完全な一覧については、「[Office アドイン用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)」 を参照してください。

ジェネレーターはプロジェクトを作成し、サポートしているノード コンポーネントをインストールします。 この記事の手順を使用して、共有ランタイムを使用するために既存の Visual Studio プロジェクトを更新することもできます。 ただし、マニフェストの XML スキーマの更新が必要になる場合があります。 詳細については、「[Office アドインでの開発エラーのトラブルシューティング](../testing/troubleshoot-development-errors.md#manifest-schema-validation-errors-in-visual-studio-projects)」を参照してください。

## <a name="configure-the-manifest"></a>マニフェストを構成する

新規または既存のプロジェクトで共有ランタイムが使用できるように構成するには、次の手順を実行します。 これらの手順は、[Office アドイン用の Yeoman ジェネレーター](yeoman-generator-overview.md)を使用してプロジェクトを生成したことを前提としています。

1. Visual Studio Code を開始し、アドイン プロジェクトを開きます。
1. **manifest.xml** ファイルを開きます。
1. Excel または PowerPoint アドインの場合は、要件セクションを更新して、[shared runtime](/javascript/api/requirement-sets/common/shared-runtime-requirement-sets) を含めます。 `CustomFunctionsRuntime` 要件が存在する場合は、必ず削除してください。 XML は、次のようになります。

    ```xml
    <Hosts>
      <Host Name="Workbook"/>
    </Hosts>
    <Requirements>
      <Sets DefaultMinVersion="1.1">
        <Set Name="SharedRuntime" MinVersion="1.1"/>
      </Sets>
    </Requirements>
    <DefaultSettings>
    ```

    > [!NOTE]
    > Word アドインのマニフェストに `SharedRuntime` 要件セットを追加しないでください。 この時点で既知の問題であるアドインを読み込むと、エラーが発生します。

1. **\<VersionOverrides\>** セクションを検索して次の **\<Runtimes\>** セクションを追加します。 作業ウィンドウを閉じてもアドイン コードを実行できるように、有効期間は **長く** する必要があります。 `resid` 値は **Taskpane.Url** で、**manifest.xml** ファイルの下部付近の `<bt:Urls>` セクションで指定された **taskpane.html** ファイルの場所を参照します。

    > [!IMPORTANT]
    > **\<Runtimes\>** セクションは、次の XML で示される正確な順序で **\<Host\>** 要素の後に入力する必要があります。

   ```xml
   <VersionOverrides ...>
     <Hosts>
       <Host ...>
         <Runtimes>
           <Runtime resid="Taskpane.Url" lifetime="long" />
         </Runtimes>
       ...
       </Host>
   ```

1. カスタム関数を使用して Excel アドインを生成した場合は、**\<Page\>** 要素を見つけます。 次に、ソースの場所を **Functions.Page.Url** から **Taskpane.Url** に変更します。

   ```xml
   <AllFormFactors>
   ...
   <Page>
     <SourceLocation resid="Taskpane.Url"/>
   </Page>
   ...
   ```

1. **\<FunctionFile\>** タグを見つけて、`resid` を **Commands.Url** から **Taskpane.Url** に変更します。 アクション コマンドがない場合は、**\<FunctionFile\>** エントリがないため、この手順は省略できます。

    ```xml
    </GetStarted>
    ...
    <FunctionFile resid="Taskpane.Url"/>
    ...
    ```

1. **manifest.xml** ファイルを保存します。

## <a name="configure-the-webpackconfigjs-file"></a>webpack.config.js ファイルを構成する

**webpack.config.js** は、複数のランタイム ローダーをビルドします。 **taskpane.html** ファイルを使用して共有ランタイムのみを読み込むには、それを変更する必要があります。

1. Visual Studio Code を起動し、生成したアドイン プロジェクトを開きます。
1. **webpack.config.js** ファイルを開きます。
1. **webpack.config.js** ファイルに次の **functions.html** プラグイン コードが含まれている場合は、それを削除します。

    ```javascript
    new HtmlWebpackPlugin({
        filename: "functions.html",
        template: "./src/functions/functions.html",
        chunks: ["polyfill", "functions"]
      })
    ```

1. **webpack.config.js** ファイルに次の **commands.html** プラグイン コードが含まれている場合は、それを削除します。

    ```javascript
    new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"]
      })
    ```

1. プロジェクトで **関数** チャンクまたは **コマンド** チャンクのいずれかを使用した場合は、次に示すようにそれらをチャンク リストに追加します (次のコードは、プロジェクトで両方のチャンクを使用した場合のものです)。

    ```javascript
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane", "commands", "functions"]
      })
    ```

1. 変更を保存してプロジェクトを再ビルドします。

   ```command line
   npm run build
   ```

> [!NOTE]
> プロジェクトに **functions.html** ファイルまたは **commands.html** ファイルがある場合は、それらを削除できます。 **taskpane.html** は、先ほど行った Webpack 更新プログラムを使用して、**functions.js** と **commands.js** コードを共有ランタイムに読み込みます。

## <a name="test-your-office-add-in-changes"></a>Office アドインの変更をテストする

次の手順を使用して、共有ランタイムを正しく使用していることを確認できます。

1. **taskpane.js** ファイルを開きます。
1. ファイルのすべての内容を次のコードで置き換えます。 これにより、作業ウィンドウが開かれた回数が表示されます。 onVisibilityModeChanged イベントの追加は、共有ランタイムでのみサポートされます。

    ```javascript
    /*global document, Office*/

    let _count = 0;

    Office.onReady(() => {
      document.getElementById("sideload-msg").style.display = "none";
      document.getElementById("app-body").style.display = "flex";

      updateCount(); // Update count on first open.
      Office.addin.onVisibilityModeChanged(function (args) {
        if (args.visibilityMode === "Taskpane") {
          updateCount(); // Update count on subsequent opens.
        }
      });
    });

    function updateCount() {
      _count++;
      document.getElementById("run").textContent = "Task pane opened " + _count + " times.";
    }
    ```

1. 変更を保存してプロジェクトを実行します。

   ```command line
   npm start
   ```

作業ウィンドウを開くたびに、開かれた回数が増加します。 **_count** の値は失われません。共有ランタイムは作業ウィンドウを閉じてもコードを実行したままにするためです。

## <a name="runtime-lifetime"></a>ランタイムの有効期間

要素を **\<Runtime\>** 追加する場合は、値 `long` または `short`. この値を `long` に設定すると、ドキュメントを開くとアドインを起動したり、作業ウィンドウを閉じた後にコードを継続して実行したり、カスタム関数から CORS および DOM を使用したりできます。

> [!NOTE]
> 既定の有効期間の値は `short` ですが、Excel、PowerPoint、および Word アドインでは `long` を使うことをお勧めします。この例でランタイムを `short` に設定した場合、いずれかのリボン ボタンを押したときにアドインが起動しますが、リボン ハンドラーの実行が完了するとアドインが終了することがあります。 同様に、作業ウィンドウを開くとアドインが起動します。ただし、作業ウィンドウを閉じると、アドインが終了する場合があります。

```xml
<Runtimes>
  <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

> [!NOTE]
> アドインにマニフェスト内の **\<Runtimes\>** 要素 (共有ランタイムに必要) が含まれており、WebView2 で Microsoft Edge を使用するための条件 (Chromium ベース) が満たされている場合は、その WebView2 コントロールが使用されます。 使用条件が満たされていない場合は、Windows または Microsoft 365 のバージョンに関係なく、Internet Explorer 11 が使用されます。 詳細については、「[ランタイム](/javascript/api/manifest/runtimes)」および「[Office アドインで使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md)」を参照してください。

## <a name="about-the-shared-runtime"></a>共有ランタイムについて

Windows または Mac では、アドインはリボン ボタン、カスタム関数、作業ウィンドウのコードを個別のランタイム環境で実行します。 これにより、グローバル データを簡単に共有できない、カスタム関数からすべての CORS 機能にアクセスできないなどの制限が発生します。

ただし、同じランタイム (共有ランタイムとも呼ばれます) でコードを共有するように Office アドインを構成できます。 これにより、アドイン間での調整が容易になり、アドインのすべての部分から、作業ウィンドウの DOM や CORS にアクセスできます。

共有ランタイムを構成すると、次のシナリオが可能になります。

- Office アドインは、追加の UI 機能を使用できます。
  - [アドイン コマンドを有効または無効にする](../design/disable-add-in-commands.md)
  - [ドキュメントが開いたら、Office アドインでコードを実行する](run-code-on-document-open.md)
  - [Office アドインの作業ウィンドウを表示または非表示にする](show-hide-add-in.md)
- 以下は、Excel アドインでのみ使用できます。
  - [Office アドインにカスタム キーボード ショートカットを追加する (プレビュー)](../design/keyboard-shortcuts.md)
  - [Office アドインでカスタム コンテキスト タブを作成する (プレビュー)](../design/contextual-tabs.md)
  - カスタム関数で CORS がすべてサポートされます。
  - カスタム関数で、Office.js API を呼び出して、スプレッドシート ドキュメントのデータを読み取ることができます。

Windows 上の Office の場合、共有ランタイムは、「[Office アドインで使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md)」で説明されているように、使用条件が満たされている場合、WebView2 (Chromium ベース) で Microsoft Edge を使用します。それ以外の場合は、Internet Explorer 11 を使用します。 また、アドインのリボンに表示するボタンはすべて、同じ共有ランタイムで実行されます。 次の図は、カスタム関数、リボン UI、作業ウィンドウ コードがすべて同じランタイムで実行される方法を示しています。

![Excel の共有ブラウザー ランタイムで実行されているカスタム関数、作業ウィンドウ、およびリボン ボタンの図。](../images/custom-functions-in-browser-runtime.png)

### <a name="debug"></a>デバッグ

共有ランタイムを使用している場合、この時点では、Windows の Excel でカスタム関数をデバッグするために Visual Studio Code を使用することはできません。 代わりに開発者ツールを使用する必要があります。 詳細については、「[Internet Explorer 用の開発者ツールを使用してアドインをデバッグする](../testing/debug-add-ins-using-f12-tools-ie.md)」または「[Microsoft Edge (Chromium ベース）の開発者ツールを使用してアドインをデバッグする](../testing/debug-add-ins-using-devtools-edge-chromium.md)」を参照してください。

### <a name="multiple-task-panes"></a>複数の作業ウィンドウ

共有ランタイムを使用する予定がある場合は、複数の作業ウィンドウを使用するようにアドインを設計しないでください。 共有ランタイムは、1 つの作業ウィンドウのみサポートします。 `<TaskpaneID>` のない作業ウィンドウは、別の作業ウィンドウとして扱われますのでご注意ください。

## <a name="see-also"></a>関連項目

- [カスタム関数から Excel API を呼び出す](../excel/call-excel-apis-from-custom-function.md)
- [Office アドインにカスタム キーボード ショートカットを追加する (プレビュー)](../design/keyboard-shortcuts.md)
- [Office アドインでカスタム コンテキスト タブを作成する (プレビュー)](../design/contextual-tabs.md)
- [アドイン コマンドを有効または無効にする](../design/disable-add-in-commands.md)
- [ドキュメントが開いたら、Office アドインでコードを実行する](run-code-on-document-open.md)
- [Office アドインの作業ウィンドウを表示または非表示にする](show-hide-add-in.md)
- [チュートリアル: Excel カスタム関数と作業ウィンドウの間でデータとイベントを共有する](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [Office アドインのランタイム](../testing/runtimes.md)
