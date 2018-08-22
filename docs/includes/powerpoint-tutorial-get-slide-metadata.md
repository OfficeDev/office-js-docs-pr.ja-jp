<span data-ttu-id="8dda7-101">このチュートリアルの手順では、選択したスライドのメタデータを取得できます。</span><span class="sxs-lookup"><span data-stu-id="8dda7-101">In this step of the tutorial, you'll retrieve metadata for the selected slide.</span></span>

> [!NOTE]
> <span data-ttu-id="8dda7-102">このページでは、PowerPoint アドインのチュートリアルの個々の手順について説明します。</span><span class="sxs-lookup"><span data-stu-id="8dda7-102">This page describes an individual step of the PowerPoint add-in tutorial.</span></span> <span data-ttu-id="8dda7-103">このページに検索エンジンの結果から、または直接リンクからアクセスした場合は、「[PowerPoint アドインのチュートリアル](../tutorials/powerpoint-tutorial.yml)」の紹介ページに移動し、チュートリアルを最初から始めてください。</span><span class="sxs-lookup"><span data-stu-id="8dda7-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [PowerPoint add-in tutorial](../tutorials/powerpoint-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="get-slide-metadata"></a><span data-ttu-id="8dda7-104">スライドのメタデータの取得</span><span class="sxs-lookup"><span data-stu-id="8dda7-104">Get slide metadata</span></span>

1. <span data-ttu-id="8dda7-105">**Home.html** ファイルで `TODO4` を次のマークアップに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="8dda7-105">In the **Home.html** file, replace `TODO4` with the following markup.</span></span> <span data-ttu-id="8dda7-106">このマークアップにより、アドインの作業ウィンドウ内に表示される **[Get Slide Metadata]** (スライドのメタデータの取得) ボタンを定義します。</span><span class="sxs-lookup"><span data-stu-id="8dda7-106">This markup defines the **Get Slide Metadata** button that will appear within the add-in's task pane.</span></span>

    ```html
    <br /><br />
    <button class="ms-Button ms-Button--primary" id="get-slide-metadata">
        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="ms-Button-label">Get Slide Metadata</span>
        <span class="ms-Button-description">Gets metadata for the selected slide(s).</span>
    </button>
    ```

2. <span data-ttu-id="8dda7-107">**Home.js** ファイルで `TODO6` を次のコードに置き換え、**[Get Slide Metadata]** (スライドのメタデータの取得) ボタンのイベント ハンドラーを割り当てます。</span><span class="sxs-lookup"><span data-stu-id="8dda7-107">In the **Home.js** file, replace `TODO6` with the following code to assign the event handler for the **Get Slide Metadata** button.</span></span>

    ```js
    $('#get-slide-metadata').click(getSlideMetadata);
    ```

3. <span data-ttu-id="8dda7-108">**Home.js** ファイルで `TODO7` を次のコードに置き換え、**getSlideMetadata** 関数を定義します。</span><span class="sxs-lookup"><span data-stu-id="8dda7-108">In the **Home.js** file, replace `TODO7` with the following code to define the **getSlideMetadata** function.</span></span> <span data-ttu-id="8dda7-109">この関数は選択したスライドのメタデータを取得し、それをアドインの作業ウィンドウ内のポップアップ ダイアログ ウィンドウに書き込みます。</span><span class="sxs-lookup"><span data-stu-id="8dda7-109">This function retrieves metadata for the selected slide(s) and writes it to a popup dialog window within the add-in task pane.</span></span>

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

## <a name="test-the-add-in"></a><span data-ttu-id="8dda7-110">アドインのテスト</span><span class="sxs-lookup"><span data-stu-id="8dda7-110">Test the add-in</span></span>

1. <span data-ttu-id="8dda7-p104">Visual Studio を使用して、アドインをテストします。そのために、`F5` キーを押すか **[開始]** ボタンをクリックして、リボンに **[作業ウィンドウの表示]** アドイン ボタンが表示された PowerPoint を起動します。アドインは IIS 上でローカルにホストされます。</span><span class="sxs-lookup"><span data-stu-id="8dda7-p104">Using Visual Studio, test the add-in by pressing `F5` or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

    ![[開始] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="8dda7-114">PowerPoint でリボンの **[作業ウィンドウの表示]** ボタンをクリックし、アドインの作業ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="8dda7-114">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![[ホーム] リボンで [作業ウィンドウの表示] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-show-taskpane-button.png)

3. <span data-ttu-id="8dda7-116">作業ウィンドウで **[Get Slide Metadata]** (スライドのメタデータの取得) ボタンを選択し、選択したスライドのメタデータを取得します。</span><span class="sxs-lookup"><span data-stu-id="8dda7-116">In the task pane, choose the **Get Slide Metadata** button to get the metadata for the selected slide.</span></span> <span data-ttu-id="8dda7-117">スライドのメタデータは作業ウィンドウの下部にあるポップアップ ダイアログ ウィンドウに書き込まれます。</span><span class="sxs-lookup"><span data-stu-id="8dda7-117">The slide metadata is written to the popup dialog window at the bottom of the task pane.</span></span> <span data-ttu-id="8dda7-118">この例では、JSON メタデータ内の `slides` 配列に、選択したスライドの `id`、`title`、および `index` を指定するオブジェクトが 1 つ含まれます。</span><span class="sxs-lookup"><span data-stu-id="8dda7-118">In this case, the `slides` array within the JSON metadata contains one object that specifies the `id`, `title`, and `index` of the selected slide.</span></span> <span data-ttu-id="8dda7-119">スライドのメタデータを取得するときに複数のスライドが選択されている場合、JSON メタデータ内の `slides` 配列には、選択したスライドごとにオブジェクトが 1 つ含まれます。</span><span class="sxs-lookup"><span data-stu-id="8dda7-119">If multiple slides had been selected when you retrieved slide metadata, the `slides` array within the JSON metadata would contain one object for each selected slide.</span></span>

    ![[Get Slide Metadata] (スライドのメタデータの取得) ボタンが強調表示されている PowerPoint アドインのスクリーンショット](../images/powerpoint-tutorial-get-slide-metadata.png)

4. <span data-ttu-id="8dda7-121">Visual Studio で `Shift + F5` を押すか **[停止]** ボタンを選択してアドインを停止します。</span><span class="sxs-lookup"><span data-stu-id="8dda7-121">In Visual Studio, stop the add-in by pressing `Shift + F5` or choosing the **Stop** button.</span></span> <span data-ttu-id="8dda7-122">アドインが停止すると、PowerPoint は自動的に閉じます。</span><span class="sxs-lookup"><span data-stu-id="8dda7-122">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![[停止] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-stop.png)
