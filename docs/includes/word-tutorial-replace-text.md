このチュートリアルの手順では、選択したテキスト範囲の内側や外側にテキストを追加したり、選択した範囲のテキストを置き換えたりします。 

> [!NOTE]
> このページでは、Word アドインのチュートリアルの個々の手順について説明します。このページに検索エンジンの結果から、またはその他の直接リンクからアクセスした場合は、「[Word アドインのチュートリアル](../tutorials/word-tutorial.yml)」の紹介ページに移動し、チュートリアルを最初から始めてください。

## <a name="add-text-inside-a-range"></a>範囲内にテキストを追加する

1. コード エディターでプロジェクトを開きます。 
2. index.html ファイルを開きます。
3. `change-font` ボタンを格納している `div` の下に、次のマークアップを追加します。

    ```html
    <div class="padding">            
        <button class="ms-Button" id="insert-text-into-range">Insert Abbreviation</button>            
    </div>
    ```

4. app.js ファイルを開きます。

5. `change-font` ボタンにクリック ハンドラーを割り当てる行の下に、次のコードを追加します。

    ```js
    $('#insert-text-into-range').click(insertTextIntoRange);
    ```

6. `changeFont` 関数の下に、次の関数を追加します。

    ```js
    function insertTextIntoRange() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to insert text into a selected range.

            // TODO2: Load the text of the range and sync so that the 
            //        current range text can be read.

            // TODO3: Queue commands to repeat the text of the original
            //        range at the end of the document.

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

7. `TODO1` を次のコードに置き換えます。次の点に注意してください。
   - このメソッドの目的は、テキストが Click-to-Run という範囲の末尾に (C2R) という省略形を挿入することです。 これは前提を単純化し、文字列は存在しており、ユーザーがその文字列を選択したものとしています。
   - `Range.insertText` メソッドの最初のパラメーターは、`Range` オブジェクトに挿入する文字列です。
   - 2 番目のパラメーターは、範囲内のどの位置にテキストを挿入するかを指定します。 End の他に、Start、Before、After、Replace が選択できます。 
   - End と After の違いは、End が既存の範囲の内部の末尾に新しいテキストを挿入するのに対し、After の場合は文字列の入った新しい範囲を作成し、既存の範囲の後にその新しい範囲が挿入されることです。 同様に、Start はテキストを既存の範囲の内部の先頭に挿入しますが、Before の場合は新しい範囲を挿入します。 Replace は、既存の範囲のテキストを最初のパラメーターで指定した文字列に置き換えます。
   - チュートリアルの前の段階で示したとおり、ボディ オブジェクトの insert* メソッドに Before オプションや After オプションはありません。 これは、文書の本文の外部にはコンテンツを挿入できないからです。

    ```js
    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText(" (C2R)", "End");
    ``` 

8. `TODO2` はスキップし、次のセクションに移ります。 `TODO3` を次のコードに置き換えます。 このコードは、このチュートリアルの最初の段階で作成したコードに似ていますが、文書の先頭ではなく末尾に新しい段落を挿入する点が異なります。 この新しい段落で、新しいテキストが元の範囲の一部になっていることが示されます。
 
    ```js
    doc.body.insertParagraph("Original range: " + originalRange.text,
                             "End");
    ``` 

## <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a>ドキュメントのプロパティを作業ウィンドウのスクリプト オブジェクトにフェッチするコードを追加する

このチュートリアルのシリーズで前述したすべての関数では、Office ドキュメントへの*書き込み*コマンドをキューに登録していました。 各関数は、キューに登録されたコマンドを実行対象のドキュメントに送信する `context.sync()` メソッドを呼び出すことで終了しています。 ただし、最後の手順で追加したコードでは、`originalRange.text` プロパティを呼び出しています。このことが、これまでに作成した関数とは大きく異なります。`originalRange` オブジェクトは、この作業ウィンドウのスクリプトに存在する単なるプロキシ オブジェクトなので、 ドキュメントの指定された範囲にある実際のテキストを認識できません。そのため、その `text` プロパティでは実際の値が保持できません。 まず、ドキュメントからその範囲のテキスト値をフェッチする必要があり、その値を使用して `originalRange.text` の値を設定します。 そのようにした場合にのみ、例外がスローされることなく `originalRange.text` を呼び出せるようになります。 このフェッチ処理には、3 つの手順があります。

   1. コードで読み取る必要があるプロパティをロードする (つまりフェッチする) コマンドをキューに登録します。
   2. コンテキスト オブジェクトの `sync` メソッドを呼び出します。このメソッドは、キューに登録されたコマンドを実行対象のドキュメントに送信して、要求された情報を返します。
   3. `sync` メソッドは非同期であるため、フェッチされたプロパティをコードで呼び出す前に、そのメソッドが完了していることを確認します。

こうした手順は、コードで Office ドキュメントから情報を*読み取る*必要がある場合には必ず完了する必要があります。

1. `TODO2` を次のコードに置き換えます。
  
    ```js
    originalRange.load("text");
    return context.sync()
        .then(function() {

                // TODO4: Move the doc.body.insertParagraph line here.
    
            }
        )
            // TODO5: Move the final call of context.sync here and ensure
            //        that it does not run until the insertParagraph has 
            //        been queued.
    ``` 

2. 分岐していない同一のコード パスに 2 つの `return` ステートメントを含めることはできないため、`Word.run` の最後にある最終行の `return context.sync();` を削除します。新しい最後の `context.sync` は、このチュートリアルの後の方で追加します。 
3. `doc.body.insertParagraph` 行を切り取り、`TODO4` の代わりに貼り付けます。 
4. `TODO5` を次のコードに置き換えます。次の点に注意してください。
   - `sync` メソッドを `then` 関数に渡すことで、`insertParagraph` ロジックがキューに登録されるまで、そのメソッドが実行されないようにします。
   - `then` メソッドは渡されたどんな関数でも呼び出します。`sync` が 2 回呼び出されないように、context.sync の末尾の "()" は省略します。

    ```js
    .then(context.sync);
    ```

作業が完了すると、関数の全体は次のようになります。

  
```js
function insertTextIntoRange() {
    Word.run(function (context) {
        
        const doc = context.document;
        const originalRange = doc.getSelection();
        originalRange.insertText(" (C2R)", "End");

        originalRange.load("text");
        return context.sync()
            .then(function() {        
                        doc.body.insertParagraph("Current text of original range: " + originalRange.text,
                                                "End");            
                }
            )
            .then(context.sync);
    })
    .catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
}
``` 

## <a name="add-text-between-ranges"></a>範囲間にテキストを追加する

1. index.html ファイルを開きます。
2. `insert-text-into-range` ボタンを格納している `div` の下に、次のマークアップを追加します。

    ```html
    <div class="padding">            
        <button class="ms-Button" id="insert-text-outside-range">Add Version Info</button>            
    </div>
    ```

3. app.js ファイルを開きます。

4. `insert-text-into-range` ボタンにクリック ハンドラーを割り当てる行の下に、次のコードを追加します。

    ```js
    $('#insert-text-outside-range').click(insertTextBeforeRange);
    ```

5. `insertTextIntoRange` 関数の下に、次の関数を追加します。

    ```js
    function insertTextBeforeRange() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to insert a new range before the 
            //        selected range.

            // TODO2: Load the text of the original range and sync so that the 
            //        range text can be read and inserted.

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
   - このメソッドの目的は、Office 365 というテキストから成る範囲の前に Office 2019 というテキストの範囲を追加することです。 これは前提を単純化し、文字列は存在しており、ユーザーがその文字列を選択したものとしています。
   - `Range.insertText` メソッドの最初のパラメーターは、追加する文字列です。
   - 2 番目のパラメーターは、範囲内のどの位置にテキストを挿入するかを指定します。 位置オプションの詳細については、`insertTextIntoRange` 関数に関する上記の説明を参照してください。

    ```js
    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText("Office 2019, ", "Before");
    ``` 

7. `TODO2` を次のコードに置き換えます。 
 
     ```js
    originalRange.load("text");
    return context.sync()
        .then(function() {

                // TODO3: Queue commands to insert the original range as a
                //        paragraph at the end of the document.
    
                }
            )

            // TODO4: Make a final call of context.sync here and ensure
            //        that it does not run until the insertParagraph has 
            //        been queued.
    ``` 

8. `TODO3` を次のコードに置き換えます。 この新しい段落で、新しいテキストが元の選択範囲の一部になって***いない***ことが示されます。 元の範囲には、依然として選択時のテキストのみが含まれています。
 
    ```js
    doc.body.insertParagraph("Current text of original range: " + originalRange.text,
                             "End");
    ``` 

9. `TODO4` を次のコードに置き換えます。

    ```js
    .then(context.sync);
    ```


## <a name="replace-the-text-of-a-range"></a>範囲のテキストを置き換える

1. index.html ファイルを開きます。
2. `insert-text-outside-range` ボタンを格納している `div` の下に、次のマークアップを追加します。

    ```html
    <div class="padding">            
        <button class="ms-Button" id="replace-text">Change Quantity Term</button>            
    </div>
    ```

3. app.js ファイルを開きます。

4. `insert-text-outside-range` ボタンにクリック ハンドラーを割り当てる行の下に、次のコードを追加します。

    ```js
    $('#replace-text').click(replaceText);
    ```

5. `insertTextBeforeRange` 関数の下に、次の関数を追加します。

    ```js
    function replaceText() {
        Word.run(function (context) {
             
            // TODO1: Queue commands to replace the text.

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

6. `TODO1` を次のコードに置き換えます。 このメソッドの目的は、several という文字列を many という文字列で置き換えることです。 これは前提を単純化し、文字列は存在しており、ユーザーがその文字列を選択したものとしています。

    ```js
    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText("many", "Replace"); 
    ``` 

## <a name="test-the-add-in"></a>アドインをテストする

1. Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトが前の段階のチュートリアルから開いたままになっている場合は、Ctrl-C を 2 回入力して実行中の Web サーバーを停止します。 それ以外の場合は、Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトを開いて、プロジェクトの **Start** フォルダーに移動します。

     > [!NOTE]
     > ブラウザー同期サーバーは、app.js ファイルなどのファイルに変更を加えるたびに作業ウィンドウ内のアドインを再読み込みしますが、JavaScript を再トランスパイルしないため、ビルド コマンドを繰り返し実行して、app.js への変更を反映させる必要があります。 これを行うには、プロンプトが表示されてビルド コマンドを入力できるようにするため、サーバー プロセスを強制終了する必要があります。 ビルド後に、サーバーを再起動します。 次の数ステップで、このプロセスを実行します。

2. `npm run build` コマンドを実行し、Office アドインを実行できるすべてのホストでサポートされている以前のバージョンの JavaScript に ES6 ソース コードをトランスパイルします。
3. `npm start` コマンドを実行して、ローカルホストで稼働する Web サーバーを起動します。
4. 作業ウィンドウを再読み込みするために、そのウィンドウを閉じて、**[ホーム]** メニューの **[作業ウィンドウの表示]** を選択してアドインを再度開きます。
5. 作業ウィンドウで **[段落の挿入]** を選択し、文書の先頭に段落があることを確認します。
6. 一部のテキストを選択します。 Click-to-Run という語句を選択します。 *選択範囲の前後にあるスペースは含めないように注意してください。*
7. **[ラベル (短縮形) の挿入]** ボタンを選択します。 (C2R) が追加されることに注意してください。 また、この新しい文字列は既存の範囲に追加されるため、文書の下部に新しい段落が追加され、拡張されたテキスト全体が含まれていることに注意してください。
8. 一部のテキストを選択します。 Office 365 という語句を選択します。 *選択範囲の前後にあるスペースは含めないように注意してください。*
9. **[バージョン情報の追加]** ボタンを選択します。 Office 2019 が、Office 2016 と Office 365 の間に挿入されることに注意してください。 また、この新しい文字列は元の範囲に追加されるのではなく新しい範囲になるため、文書の下部に新しい段落が追加されるものの、その段落には最初に選択したテキストのみが含まれることに注意してください。
10. 一部のテキストを選択します。 several という語句を選択します。 *選択範囲の前後にあるスペースは含めないように注意してください。*
11. **[数量の用語の変更]** ボタンを選択します。 選択したテキストが many に置き換えられることに注意してください。

    ![Word のチュートリアル - テキストの追加と置換](../images/word-tutorial-text-replace.png)
