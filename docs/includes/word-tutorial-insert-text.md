チュートリアルのこの手順では、ユーザーが現在使用している Word のバージョンをアドインがサポートしているかどうかをプログラムによってテストし、ドキュメントにパラグラフを挿入します。

> [!NOTE]
> このページでは、Word アドインのチュートリアルの個々の手順について説明します。このページに検索エンジンの結果から、またはその他の直接リンクからアクセスした場合は、「[Word アドインのチュートリアル](../tutorials/word-tutorial.yml)」の紹介ページに移動し、チュートリアルを最初から始めてください。

## <a name="code-the-add-in"></a>アドインのコードを作成する

1. コード エディターでプロジェクトを開きます。 
2. index.html ファイルを開きます。
3. を次のマークアップに置き換えます。`TODO1`

    ```html
    <button class="ms-Button" id="insert-paragraph">Insert Paragraph</button>
    ```

4. app.js ファイルを開きます。
5. を次のコードに置き換えます。`TODO1` このコードでは、ユーザーの Word のバージョンが、このチュートリアルのすべての段階で使用するすべての API を含んでいる Word.js のバージョンをサポートしているかどうかを調べます。 運用アドインでは、未サポートの API を呼び出す UI を非表示または無効化する条件ブロックの本体を使用してください。 これにより、ユーザーは、自分が使用している Word のバージョンでサポートされているアドインの部分を使用できるようになります。

    ```js
    if (!Office.context.requirements.isSetSupported('WordApi', 1.3)) {
        console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
    } 
    ```

6. を次のコードに置き換えます。`TODO2`

    ```js
    $('#insert-paragraph').click(insertParagraph);
    ```

7. を次のコードに置き換えます。`TODO3` 次の点に注意してください。
   - Word.js のビジネス ロジックは、`Word.run` に渡される関数に追加されます。 このロジックは、すぐには実行されません。 その代わりに、保留中のコマンドのキューに追加されます。
   - メソッドは、キューに登録されたすべてのコマンドを、実行するために Word に送信します。`context.sync`
   - の後に `catch` ブロックを続けます。`Word.run` これは、どのような場合にも当てはまるベスト プラクティスです。 

    ```js
    function insertParagraph() {
        Word.run(function (context) {
            
            // TODO4: Queue commands to insert a paragraph into the document.

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

8. |||UNTRANSLATED_CONTENT_START|||Replace `TODO4` with the following code. Note:|||UNTRANSLATED_CONTENT_END|||
   - メソッドの最初のパラメーターは、新しい段落のテキストです。`insertParagraph`
   - 2 番目のパラメーターは、段落を挿入する本文内の場所です。 親オブジェクトが本文の場合、段落の挿入に使用できるその他のオプションには、End と Replace があります。 

    ```js
    const docBody = context.document.body;
    docBody.insertParagraph("Office has several versions, including Office 2016, Office 365 Click-to-Run, and Office Online.",
                            "Start");   
    ``` 

## <a name="test-the-add-in"></a>アドインをテストする

1. Git bash ウィンドウまたは Node.JS 対応のシステム プロンプトを開いて、プロジェクトの **Start** フォルダーに移動します。
2. コマンドを実行し、Office アドインを実行できるすべてのホストでサポートされている以前のバージョンの JavaScript に ES6 ソース コードをトランスパイルします。`npm run build`
3. コマンドを実行して、ローカルホストで稼働する Web サーバーを起動します。`npm start`   
4. 次のいずれかの方法を使用して、アドインをサイドロードします。
    - Windows: [Windows で Office アドインをサイドロードする](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Word Online: [Office Online で Office アドインをサイドロードする](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)
    - iPad および Mac: [iPad と Mac で Office アドインをサイドロードする](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)
5. Word の **[ホーム]** メニューで、**[作業ウィンドウの表示]** を選択します。
6. 作業ウィンドウで、**[段落の挿入]** を選択します。
7. 段落に変更を加えます。 
8. **[段落の挿入]** をもう一度選択します。 メソッドはドキュメントの本文の開始位置に挿入を行うため、新しい段落は前の段落より上に追加されます。`insertParagraph`

    ![Word のチュートリアル - 段落の挿入](../images/word-tutorial-insert-paragraph.png)
