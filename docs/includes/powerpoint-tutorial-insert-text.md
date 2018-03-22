このチュートリアルの手順では、その日の [Bing](https://www.bing.com) 写真を含むタイトル スライドにテキストを追加します。

> [!NOTE]
> このページでは、PowerPoint アドインのチュートリアルの個々の手順について説明します。 このページに検索エンジンの結果から、または直接リンクからアクセスした場合は、「[PowerPoint アドインのチュートリアル](../tutorials/powerpoint-tutorial.yml)」の紹介ページに移動し、チュートリアルを最初から始めてください。

## <a name="add-text-to-a-slide"></a>スライドにテキストを追加する 

1. **Home.html** ファイルで `TODO3` を次のマークアップに置き換えます。 このマークアップにより、アドインの作業ウィンドウ内に表示される **[テキストの挿入]** ボタンを定義します。

    ```html
        <br /><br />
        <button class="ms-Button ms-Button--primary" id="insert-text">
            <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
            <span class="ms-Button-label">Insert Text</span>
            <span class="ms-Button-description">Inserts text into the slide.</span>
        </button>
    ```

2. **Home.js** ファイルで `TODO4` を次のコードに置き換え、**[テキストの挿入]** ボタンのイベント ハンドラーを割り当てます。

    ```js
    $('#insert-text').click(insertText);
    ```

3. **Home.js** ファイルで `TODO5` を次のコードに置き換え、**insertText** 関数を定義します。 この関数は、現在のスライドにテキストを挿入します。

    ```js
    function insertText() {
        Office.context.document.setSelectedDataAsync('Hello World!',
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }
    ```

## <a name="test-the-add-in"></a>アドインのテスト

1. Visual Studio を使用して、アドインをテストします。そのために、`F5` キーを押すか **[開始]** ボタンをクリックして、リボンに **[作業ウィンドウの表示]** アドイン ボタンが表示された PowerPoint を起動します。アドインは IIS 上でローカルにホストされます。

    ![[開始] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-start.png)

2. PowerPoint でリボンの **[作業ウィンドウの表示]** ボタンをクリックし、アドインの作業ウィンドウを開きます。

    ![[ホーム] リボンで [作業ウィンドウの表示] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-show-taskpane-button.png)

3. 作業ウィンドウで **[イメージの挿入]** ボタンをクリックしてその日の Bing 写真を現在のスライドに追加し、そのタイトルにテキスト ボックスが含まれるデザインをそのスライドに選択します。

    ![[イメージの挿入] ボタンが強調表示されている PowerPoint アドインのスクリーンショット](../images/powerpoint-tutorial-insert-image-slide-design.png)

4. タイトル スライドのテキスト ボックスにカーソルを置き、作業ウィンドウで **[テキストの挿入]** ボタンをクリックしてテキストをスライドに追加します。

    ![[テキストの挿入] ボタンが強調表示されている PowerPoint アドインのスクリーンショット](../images/powerpoint-tutorial-insert-text.png)


5. Visual Studio で `Shift + F5` を押すか **[停止]** ボタンを選択してアドインを停止します。 アドインが停止すると、PowerPoint は自動的に閉じます。

    ![[停止] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-stop.png)