チュートリアルのこの手順では、ドキュメントに画像、HTML、テーブルを挿入する方法について説明します。

> [!NOTE]
> このページでは、Word アドインのチュートリアルの個々の手順について説明します。このページに検索エンジンの結果から、またはその他の直接リンクからアクセスした場合は、「[Word アドインのチュートリアル](../tutorials/word-tutorial.yml)」の紹介ページに移動し、チュートリアルを最初から始めてください。

## <a name="insert-an-image"></a>画像の挿入

1. コード エディターでプロジェクトを開きます。 
2. index.html ファイルを開きます。
3. `replace-text` ボタンを格納している `div` の下に、次のマークアップを追加します。

    ```html
    <div class="padding">            
        <button class="ms-Button" id="insert-image">Insert Image</button>            
    </div>
    ```

4. app.js ファイルを開きます。

5. ファイルの先頭近くにある、use-strict 行のすぐ下に次の行を追加します。 この行は、別のファイルから変数をインポートします。 この変数は、画像をエンコードする Base 64 文字列です。 エンコードされた文字列を表示するには、プロジェクトのルートにある base64Image.js ファイルを開きます。

    ```js
    import { base64Image } from "./base64Image";
    ``` 

5. `replace-text` ボタンにクリック ハンドラーを割り当てる行の下に、次のコードを追加します。

    ```js
    $('#insert-image').click(insertImage);
    ```

6. `replaceText` 関数の下に、次の関数を追加します。

    ```js
    function insertImage() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to insert an image.

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

7. `TODO1` を次のコードに置き換えます。 この行により、Base 64 でエンコードされた画像がドキュメントの末尾に挿入されることに注意してください。 (`Paragraph` オブジェクトにも `insertInlinePictureFromBase64` メソッドやその他の `insert*` メソッドがあります。 例については、次の insertHTML セクションを参照してください)。

    ```js
    context.document.body.insertInlinePictureFromBase64(base64Image, "End");
    ``` 

## <a name="insert-html"></a>HTML の挿入

1. index.html ファイルを開きます。
2. `insert-image` ボタンを格納している `div` の下に、次のマークアップを追加します。

    ```html
    <div class="padding">            
        <button class="ms-Button" id="insert-html">Insert HTML</button>            
    </div>
    ```

3. app.js ファイルを開きます。

4. `insert-image` ボタンにクリック ハンドラーを割り当てる行の下に、次のコードを追加します。

    ```js
    $('#insert-html').click(insertHTML);
    ```

5. `insertImage` 関数の下に、次の関数を追加します。

    ```js
    function insertHTML() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to insert a string of HTML.

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

6. `TODO1` を次のコードに置き換えます。次の点に注意してください。
   - 最初の行は、ドキュメントの末尾に空白の段落を追加します。 
   - 2 行目は、その段落の末尾に HTML の文字列を挿入します。具体的には、Verdana フォントで書式設定された段落と、Word 文書の既定のスタイルが設定された段落の 2 つの段落が挿入されます。 (`insertImage` メソッドで説明したように、`context.document.body` オブジェクトにも `insert*` メソッドがあります)。

    ```js
    const blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", "After");
    blankParagraph.insertHtml('<p style="font-family: verdana;">Inserted HTML.</p><p>Another paragraph</p>', "End");
    ``` 

## <a name="insert-table"></a>テーブルの挿入

1. index.html ファイルを開きます。
3. `insert-html` ボタンを格納している `div` の下に、次のマークアップを追加します。

    ```html
    <div class="padding">            
        <button class="ms-Button" id="insert-table">Insert Table</button>            
    </div>
    ```

4. app.js ファイルを開きます。

5. `insert-html` ボタンにクリック ハンドラーを割り当てる行の下に、次のコードを追加します。

    ```js
    $('#insert-table').click(insertTable);
    ```

6. `insertHTML` 関数の下に、次の関数を追加します。

    ```js
    function insertTable() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to get a reference to the paragraph
            //        that will proceed the table.

            // TODO2: Queue commands to create a table and populate it with data.

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

7. `TODO1` を次のコードに置き換えます。 この行は `ParapgraphCollection.getFirst` メソッドを使用して最初の段落への参照を取得し、次に `Paragraph.getNext` メソッドを使用して 2 番目の段落への参照を取得することに注意してください。

    ```js
    const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    ``` 

8. `TODO2` を次のコードに置き換えます。次の点に注意してください。
   - `insertTable` メソッドの最初の 2 つのパラメーターは、行と列の数を指定します。
   - 3 番目のパラメーターは、テーブルを挿入する場所を指定します (この例では段落の後)。
   - 4 番目のパラメーターは、テーブルのセルの値を設定する 2 次元配列です。
   - このテーブルには既定のスタイルがそのまま設定されますが、`insertTable` メソッドがさまざまなメンバーを持つ `Table` オブジェクトを返し、その一部がテーブルのスタイル設定に使用されます。

     ```js
    const tableData = [
            ["Name", "ID", "Birth City"],
            ["Bob", "434", "Chicago"],
            ["Sue", "719", "Havana"],
        ];
    secondParagraph.insertTable(3, 3, "After", tableData);
    ``` 

## <a name="test-the-add-in"></a>アドインをテストする


1. Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトが前の段階のチュートリアルから開いたままになっている場合は、Ctrl + C を 2 回入力して実行中の Web サーバーを停止します。 それ以外の場合は、Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトを開いて、プロジェクトの **Start** フォルダーに移動します。

     > [!NOTE]
     > ブラウザー同期サーバーは、app.js ファイルなどのファイルに変更を加えるたびに作業ウィンドウ内のアドインを再読み込みしますが、JavaScript を再トランスパイルしないため、ビルド コマンドを繰り返し実行して、app.js への変更を反映させる必要があります。 これを行うには、プロンプトが表示されてビルド コマンドを入力できるようにするため、サーバー プロセスを強制終了する必要があります。 ビルド後に、サーバーを再起動します。 次の数ステップで、このプロセスを実行します。

2. `npm run build` コマンドを実行し、Office アドインを実行できるすべてのホストでサポートされている以前のバージョンの JavaScript に ES6 ソース コードをトランスパイルします。
3. `npm start` コマンドを実行して、ローカルホストで稼働する Web サーバーを起動します。
4. 作業ウィンドウを再読み込みするために、そのウィンドウを閉じて、**[ホーム]** メニューの **[作業ウィンドウの表示]** を選択してアドインを再度開きます。
5. 作業ウィンドウで **[段落の挿入]** を少なくとも 3 回選択し、ドキュメントに段落がいくつかあることを確認します。
6. **[画像の挿入]** ボタンをクリックし、ドキュメントの末尾に画像が挿入されることに注意してください。
7. **[HTML の挿入]** ボタンをクリックし、ドキュメントの末尾に 2 つの段落が挿入され、最初の段落に Verdana フォントが設定されていることに注意してください。
8. **[テーブルの挿入]** ボタンをクリックし、2 番目の段落の後にテーブルが挿入されることに注意してください。

    ![Word のチュートリアル - 画像、HTML、テーブルの挿入](../images/word-tutorial-insert-image-html-table.png)
