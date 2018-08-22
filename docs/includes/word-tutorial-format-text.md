チュートリアルのこの手順では、テキストのフォントを変更して、テキストに組み込みのスタイルやカスタム スタイルを使用します。

> [!NOTE]
> このページでは、Word アドインのチュートリアルの個々の手順について説明します。このページに検索エンジンの結果から、またはその他の直接リンクからアクセスした場合は、「[Word アドインのチュートリアル](../tutorials/word-tutorial.yml)」の紹介ページに移動し、チュートリアルを最初から始めてください。

## <a name="apply-a-built-in-style-to-text"></a>組み込みのスタイルをテキストに適用する

1. コード エディターでプロジェクトを開きます。 
2. index.html ファイルを開きます。
3. `insert-paragraph` ボタンを格納している `div` の直下に、次のマークアップを追加します。

    ```html
    <div class="padding">            
        <button class="ms-Button" id="apply-style">Apply Style</button>            
    </div>
    ```

4. app.js ファイルを開きます。

5. `insert-paragraph` ボタンにクリック ハンドラーを割り当てる行の直下に、次のコードを追加します。

    ```js
    $('#apply-style').click(applyStyle);
    ```

6. `insertParagraph` 関数の直下に、次の関数を追加します。

    ```js
    function applyStyle() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to style text.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ``` 

7. `TODO1` を次のコードに置き換えます。 このコードではスタイルを段落に適用していますが、スタイルはテキストの範囲にも適用できます。

    ```js
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.styleBuiltIn = Word.Style.intenseReference;
    ``` 

## <a name="apply-a-custom-style-to-text"></a>カスタム スタイルをテキストに適用する

1. index.html ファイルを開きます。
2. ボタンを格納している `div` の下に、次のマークアップを追加します。`apply-style`

    ```html
    <div class="padding">            
        <button class="ms-Button" id="apply-custom-style">Apply Custom Style</button>            
    </div>
    ```

3. app.js ファイルを開きます。

4. ボタンにクリック ハンドラーを割り当てる行の下に、次のコードを追加します。`apply-style`

    ```js
    $('#apply-custom-style').click(applyCustomStyle);
    ```

5. 関数の下に、次の関数を追加します。`applyStyle`

    ```js
    function applyCustomStyle() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to apply the custom style.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ``` 

7. `TODO1` を次のコードに置き換えます。 このコードでは、まだ存在していないカスタム スタイルを適用しています。 「[アドインをテストする](#test-the-add-in)」の手順で **MyCustomStyle** という名前のスタイルを作成します。

    ```js
    const lastParagraph = context.document.body.paragraphs.getLast();
    lastParagraph.style = "MyCustomStyle";
    ``` 

## <a name="change-the-font-of-text"></a>テキストのフォントを変更する

1. index.html ファイルを開きます。
2. ボタンを格納している `div` の下に、次のマークアップを追加します。`apply-custom-style`

    ```html
    <div class="padding">            
        <button class="ms-Button" id="change-font">Change Font</button>            
    </div>
    ```

3. app.js ファイルを開きます。

4. ボタンにクリック ハンドラーを割り当てる行の下に、次のコードを追加します。`apply-custom-style`

    ```js
    $('#change-font').click(changeFont);
    ```

5. 関数の下に、次の関数を追加します。`applyCustomStyle`

    ```js
    function changeFont() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to apply a different font.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ``` 

7. `TODO1` を次のコードに置き換えます。 このコードでは、`Paragraph.getNext` メソッドにチェーンされた `ParagraphCollection.getFirst` メソッドを使用して 2 番目の段落への参照を取得することに注意してください。

    ```js
    const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    secondParagraph.font.set({
            name: "Courier New",
            bold: true,
            size: 18
        });
    ``` 

## <a name="test-the-add-in"></a>アドインをテストする

1. Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトが前の段階のチュートリアルから開いたままになっている場合は、Ctrl + C を 2 回入力して実行中の Web サーバーを停止します。 それ以外の場合は、Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトを開いて、プロジェクトの **Start** フォルダーに移動します。

     > [!NOTE]
     > ブラウザー同期サーバーは、app.js ファイルなどのファイルに変更を加えるたびに作業ウィンドウ内のアドインを再読み込みしますが、JavaScript を再トランスパイルしないため、ビルド コマンドを繰り返し実行して、app.js への変更を反映させる必要があります。 これを行うには、プロンプトが表示されてビルド コマンドを入力できるようにするため、サーバー プロセスを強制終了する必要があります。 ビルド後に、サーバーを再起動します。 次の数ステップで、このプロセスを実行します。

2. コマンドを実行し、Office アドインを実行できるすべてのホストでサポートされている以前のバージョンの JavaScript に ES6 ソース コードをトランスパイルします。`npm run build`
3. `npm start` コマンドを実行して、ローカルホストで稼働する Web サーバーを起動します。   
4. 作業ウィンドウを再読み込みするために、そのウィンドウを閉じて、**[ホーム]** メニューの **[作業ウィンドウの表示]** を選択してアドインを再度開きます。
5. ドキュメントに 3 つ以上の段落があることを確認してください。 **[段落の挿入]** を 3 回選択できます。 *ドキュメントの最後に空白の段落がないことを慎重にチェックしてください。空白の段落がある場合は、それを削除します。*
6. Word で、MyCustomStyle という名前のカスタム スタイルを作成します。 このスタイルには、必要に応じて任意の書式を設定できます。
7. **[スタイルの適用]** ボタンを選択します。 最初の段落は、組み込みのスタイルである **Intense Reference** でスタイル設定されます。
8. **[カスタム スタイルの適用]** ボタンを選択します。 最後の段落は、選択したカスタム スタイルでスタイル設定されます。 (何も起こらないように思える場合、最後の段落が空白である可能性があります。 その場合は、最後の段落にテキストを追加します)。
9. **[フォントの変更]** ボタンを選択します。 2 番目の段落のフォントを、18 ポイントで太字の Courier New に変更します。

    ![Word のチュートリアル - スタイルとフォントを適用する](../images/word-tutorial-apply-styles-and-font.png)
