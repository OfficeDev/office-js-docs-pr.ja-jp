このチュートリアルの手順では、ドキュメントのスライド間を移動します。

> [!NOTE]
> このページでは、PowerPoint アドインのチュートリアルの個々の手順について説明します。 このページに検索エンジンの結果から、または直接リンクからアクセスした場合は、「[PowerPoint アドインのチュートリアル](../tutorials/powerpoint-tutorial.yml)」の紹介ページに移動し、チュートリアルを最初から始めてください。

## <a name="navigate-between-slides-of-the-document"></a>ドキュメントのスライド間を移動する

1. **Home.html** ファイルで `TODO5` を次のマークアップに置き換えます。 このマークアップにより、アドインの作業ウィンドウ内に表示される 4 つのナビゲーション ボタンを定義します。

    ```html
    <br /><br />
    <button class="ms-Button ms-Button--primary" id="go-to-first-slide">
        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="ms-Button-label">Go to First Slide</span>
        <span class="ms-Button-description">Go to the first slide.</span>
    </button>
    <br /><br />
    <button class="ms-Button ms-Button--primary" id="go-to-next-slide">
        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="ms-Button-label">Go to Next Slide</span>
        <span class="ms-Button-description">Go to the next slide.</span>
    </button>
    <br /><br />
    <button class="ms-Button ms-Button--primary" id="go-to-previous-slide">
        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="ms-Button-label">Go to Previous Slide</span>
        <span class="ms-Button-description">Go to the previous slide.</span>
    </button>
    <br /><br />
    <button class="ms-Button ms-Button--primary" id="go-to-last-slide">
        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="ms-Button-label">Go to Last Slide</span>
        <span class="ms-Button-description">Go to the last slide.</span>
    </button>
    ```

2. **Home.js** ファイルで `TODO8` を次のコードに置き換え、4 つのナビゲーション ボタンのイベント ハンドラーを割り当てます。

    ```js
    $('#go-to-first-slide').click(goToFirstSlide);
    $('#go-to-next-slide').click(goToNextSlide);
    $('#go-to-previous-slide').click(goToPreviousSlide);
    $('#go-to-last-slide').click(goToLastSlide);
    ```

3. **Home.js** ファイルで `TODO9` を次のコードに置き換え、ナビゲーション関数を定義します。 これらの関数では `goToByIdAsync` 関数を使用して、ドキュメント内のその位置 (最初、最後、前、次) に基づいてスライドを選択します。

    ```js
    function goToFirstSlide() {
        Office.context.document.goToByIdAsync(Office.Index.First, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function goToLastSlide() {
        Office.context.document.goToByIdAsync(Office.Index.Last, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function goToPreviousSlide() {
        Office.context.document.goToByIdAsync(Office.Index.Previous, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function goToNextSlide() {
        Office.context.document.goToByIdAsync(Office.Index.Next, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
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


3. **[ホーム]** タブの **[新しいスライド]** ボタンを使用して、2 つの新しいスライドをドキュメントに追加します。 

4. 作業ウィンドウで **[最初のスライドに移動]** ボタンをクリックします。 ドキュメントの最初のスライドが選択され、表示されます。

    ![[最初のスライドに移動] ボタンが強調表示されている PowerPoint アドインのスクリーンショット](../images/powerpoint-tutorial-go-to-first-slide.png)

5. 作業ウィンドウで **[次のスライドに移動]** ボタンをクリックします。 ドキュメントの次のスライドが選択され、表示されます。

    ![[次のスライドに移動] ボタンが強調表示されている PowerPoint アドインのスクリーンショット](../images/powerpoint-tutorial-go-to-next-slide.png)

6. 作業ウィンドウで **[前のスライドに移動]** ボタンをクリックします。 ドキュメントの前のスライドが選択され、表示されます。

    ![[前のスライドに移動] ボタンが強調表示されている PowerPoint アドインのスクリーンショット](../images/powerpoint-tutorial-go-to-previous-slide.png)

7. 作業ウィンドウで **[最後のスライドに移動]** ボタンをクリックします。 ドキュメントの最後のスライドが選択され、表示されます。

    ![[最後のスライドに移動] ボタンが強調表示されている PowerPoint アドインのスクリーンショット](../images/powerpoint-tutorial-go-to-last-slide.png)

8. Visual Studio で `Shift + F5` を押すか **[停止]** ボタンを選択してアドインを停止します。 アドインが停止すると、PowerPoint は自動的に閉じます。

    ![[停止] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-stop.png)
