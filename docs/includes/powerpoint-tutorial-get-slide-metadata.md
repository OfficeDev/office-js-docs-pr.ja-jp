このチュートリアルの手順では、選択したスライドのメタデータを取得できます。

> [!NOTE]
> このページでは、PowerPoint アドインのチュートリアルの個々の手順について説明します。 このページに検索エンジンの結果から、または直接リンクからアクセスした場合は、「[PowerPoint アドインのチュートリアル](../tutorials/powerpoint-tutorial.yml)」の紹介ページに移動し、チュートリアルを最初から始めてください。

## <a name="get-slide-metadata"></a>スライドのメタデータの取得

1. **Home.html** ファイルで `TODO4` を次のマークアップに置き換えます。 このマークアップにより、アドインの作業ウィンドウ内に表示される **[Get Slide Metadata]** (スライドのメタデータの取得) ボタンを定義します。

    ```html
    <br /><br />
    <button class="ms-Button ms-Button--primary" id="get-slide-metadata">
        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="ms-Button-label">Get Slide Metadata</span>
        <span class="ms-Button-description">Gets metadata for the selected slide(s).</span>
    </button>
    ```

2. **Home.js** ファイルで `TODO6` を次のコードに置き換え、**[Get Slide Metadata]** (スライドのメタデータの取得) ボタンのイベント ハンドラーを割り当てます。

    ```js
    $('#get-slide-metadata').click(getSlideMetadata);
    ```

3. **Home.js** ファイルで `TODO7` を次のコードに置き換え、**getSlideMetadata** 関数を定義します。 この関数は選択したスライドのメタデータを取得し、それをアドインの作業ウィンドウ内のポップアップ ダイアログ ウィンドウに書き込みます。

    ```js
    function getSlideMetadata() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange,
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                } else {
                    showNotification("Metadata for selected slide(s):", JSON.stringify(asyncResult.value), null, 2);
                }
            }
        );
    }
    ```

## <a name="test-the-add-in"></a>アドインのテスト

1. Visual Studio を使用して、アドインをテストします。そのために、`F5` キーを押すか **[開始]** ボタンをクリックして、リボンに **[作業ウィンドウの表示]** アドイン ボタンが表示された PowerPoint を起動します。アドインは IIS 上でローカルにホストされます。

    ![[開始] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-start.png)

2. PowerPoint でリボンの **[作業ウィンドウの表示]** ボタンをクリックし、アドインの作業ウィンドウを開きます。

    ![[ホーム] リボンで [作業ウィンドウの表示] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-show-taskpane-button.png)

3. 作業ウィンドウで **[Get Slide Metadata]** (スライドのメタデータの取得) ボタンを選択し、選択したスライドのメタデータを取得します。 スライドのメタデータは作業ウィンドウの下部にあるポップアップ ダイアログ ウィンドウに書き込まれます。 この例では、JSON メタデータ内の `slides` 配列に、選択したスライドの `id`、`title`、および `index` を指定するオブジェクトが 1 つ含まれます。 スライドのメタデータを取得するときに複数のスライドが選択されている場合、JSON メタデータ内の `slides` 配列には、選択したスライドごとにオブジェクトが 1 つ含まれます。

    ![[Get Slide Metadata] (スライドのメタデータの取得) ボタンが強調表示されている PowerPoint アドインのスクリーンショット](../images/powerpoint-tutorial-get-slide-metadata.png)

4. Visual Studio で `Shift + F5` を押すか **[停止]** ボタンを選択してアドインを停止します。 アドインが停止すると、PowerPoint は自動的に閉じます。

    ![[停止] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-stop.png)
