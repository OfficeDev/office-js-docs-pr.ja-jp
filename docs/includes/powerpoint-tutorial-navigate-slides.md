<span data-ttu-id="0fa15-101">このチュートリアルの手順では、ドキュメントのスライド間を移動します。</span><span class="sxs-lookup"><span data-stu-id="0fa15-101">In this step of the tutorial, you'll navigate between the slides of a document.</span></span>

> [!NOTE]
> <span data-ttu-id="0fa15-102">このページでは、PowerPoint アドインのチュートリアルの個々の手順について説明します。</span><span class="sxs-lookup"><span data-stu-id="0fa15-102">This page describes an individual step of the PowerPoint add-in tutorial.</span></span> <span data-ttu-id="0fa15-103">このページに検索エンジンの結果から、または直接リンクからアクセスした場合は、「[PowerPoint アドインのチュートリアル](../tutorials/powerpoint-tutorial.yml)」の紹介ページに移動し、チュートリアルを最初から始めてください。</span><span class="sxs-lookup"><span data-stu-id="0fa15-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [PowerPoint add-in tutorial](../tutorials/powerpoint-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="navigate-between-slides-of-the-document"></a><span data-ttu-id="0fa15-104">ドキュメントのスライド間を移動する</span><span class="sxs-lookup"><span data-stu-id="0fa15-104">Navigate between slides of the document</span></span>

1. <span data-ttu-id="0fa15-105">**Home.html** ファイルで `TODO5` を次のマークアップに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="0fa15-105">In the **Home.html** file, replace `TODO5` with the following markup.</span></span> <span data-ttu-id="0fa15-106">このマークアップにより、アドインの作業ウィンドウ内に表示される 4 つのナビゲーション ボタンを定義します。</span><span class="sxs-lookup"><span data-stu-id="0fa15-106">This markup defines the four navigation buttons that will appear within the add-in's task pane.</span></span>

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

2. <span data-ttu-id="0fa15-107">**Home.js** ファイルで `TODO8` を次のコードに置き換え、4 つのナビゲーション ボタンのイベント ハンドラーを割り当てます。</span><span class="sxs-lookup"><span data-stu-id="0fa15-107">In the **Home.js** file, replace `TODO8` with the following code to assign the event handlers for the four navigation buttons.</span></span>

    ```js
    $('#go-to-first-slide').click(goToFirstSlide);
    $('#go-to-next-slide').click(goToNextSlide);
    $('#go-to-previous-slide').click(goToPreviousSlide);
    $('#go-to-last-slide').click(goToLastSlide);
    ```

3. <span data-ttu-id="0fa15-108">**Home.js** ファイルで `TODO9` を次のコードに置き換え、ナビゲーション関数を定義します。</span><span class="sxs-lookup"><span data-stu-id="0fa15-108">In the **Home.js** file, replace `TODO9` with the following code to define the navigation functions.</span></span> <span data-ttu-id="0fa15-109">これらの関数では `goToByIdAsync` 関数を使用して、ドキュメント内のその位置 (最初、最後、前、次) に基づいてスライドを選択します。</span><span class="sxs-lookup"><span data-stu-id="0fa15-109">Each of these functions uses the `goToByIdAsync` function to select a slide based upon its position in the document (first, last, previous, next).</span></span>

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

## <a name="test-the-add-in"></a><span data-ttu-id="0fa15-110">アドインのテスト</span><span class="sxs-lookup"><span data-stu-id="0fa15-110">Test the add-in</span></span>

1. <span data-ttu-id="0fa15-p104">Visual Studio を使用して、アドインをテストします。そのために、`F5` キーを押すか **[開始]** ボタンをクリックして、リボンに **[作業ウィンドウの表示]** アドイン ボタンが表示された PowerPoint を起動します。アドインは IIS 上でローカルにホストされます。</span><span class="sxs-lookup"><span data-stu-id="0fa15-p104">Using Visual Studio, test the add-in by pressing `F5` or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

    ![[開始] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="0fa15-114">PowerPoint でリボンの **[作業ウィンドウの表示]** ボタンをクリックし、アドインの作業ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="0fa15-114">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![[ホーム] リボンで [作業ウィンドウの表示] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-show-taskpane-button.png)


3. <span data-ttu-id="0fa15-116">**[ホーム]** タブの **[新しいスライド]** ボタンを使用して、2 つの新しいスライドをドキュメントに追加します。</span><span class="sxs-lookup"><span data-stu-id="0fa15-116">Use the **New Slide** button in the ribbon of the **Home** tab to add two new slides to the document.</span></span> 

4. <span data-ttu-id="0fa15-117">作業ウィンドウで **[最初のスライドに移動]** ボタンをクリックします。</span><span class="sxs-lookup"><span data-stu-id="0fa15-117">In the task pane, choose the **Go to First Slide** button.</span></span> <span data-ttu-id="0fa15-118">ドキュメントの最初のスライドが選択され、表示されます。</span><span class="sxs-lookup"><span data-stu-id="0fa15-118">The first slide in the document is selected and displayed.</span></span>

    ![[最初のスライドに移動] ボタンが強調表示されている PowerPoint アドインのスクリーンショット](../images/powerpoint-tutorial-go-to-first-slide.png)

5. <span data-ttu-id="0fa15-120">作業ウィンドウで **[次のスライドに移動]** ボタンをクリックします。</span><span class="sxs-lookup"><span data-stu-id="0fa15-120">In the task pane, choose the **Go to Next Slide** button.</span></span> <span data-ttu-id="0fa15-121">ドキュメントの次のスライドが選択され、表示されます。</span><span class="sxs-lookup"><span data-stu-id="0fa15-121">The next slide in the document is selected and displayed.</span></span>

    ![[次のスライドに移動] ボタンが強調表示されている PowerPoint アドインのスクリーンショット](../images/powerpoint-tutorial-go-to-next-slide.png)

6. <span data-ttu-id="0fa15-123">作業ウィンドウで **[前のスライドに移動]** ボタンをクリックします。</span><span class="sxs-lookup"><span data-stu-id="0fa15-123">In the task pane, choose the **Go to Previous Slide** button.</span></span> <span data-ttu-id="0fa15-124">ドキュメントの前のスライドが選択され、表示されます。</span><span class="sxs-lookup"><span data-stu-id="0fa15-124">The previous slide in the document is selected and displayed.</span></span>

    ![[前のスライドに移動] ボタンが強調表示されている PowerPoint アドインのスクリーンショット](../images/powerpoint-tutorial-go-to-previous-slide.png)

7. <span data-ttu-id="0fa15-126">作業ウィンドウで **[最後のスライドに移動]** ボタンをクリックします。</span><span class="sxs-lookup"><span data-stu-id="0fa15-126">In the task pane, choose the **Go to Last Slide** button.</span></span> <span data-ttu-id="0fa15-127">ドキュメントの最後のスライドが選択され、表示されます。</span><span class="sxs-lookup"><span data-stu-id="0fa15-127">The last slide in the document is selected and displayed.</span></span>

    ![[最後のスライドに移動] ボタンが強調表示されている PowerPoint アドインのスクリーンショット](../images/powerpoint-tutorial-go-to-last-slide.png)

8. <span data-ttu-id="0fa15-129">Visual Studio で `Shift + F5` を押すか **[停止]** ボタンを選択してアドインを停止します。</span><span class="sxs-lookup"><span data-stu-id="0fa15-129">In Visual Studio, stop the add-in by pressing `Shift + F5` or choosing the **Stop** button.</span></span> <span data-ttu-id="0fa15-130">アドインが停止すると、PowerPoint は自動的に閉じます。</span><span class="sxs-lookup"><span data-stu-id="0fa15-130">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![[停止] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-stop.png)
