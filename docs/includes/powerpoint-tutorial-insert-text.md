<span data-ttu-id="5f852-101">このチュートリアルの手順では、その日の [Bing](https://www.bing.com) 写真を含むタイトル スライドにテキストを追加します。</span><span class="sxs-lookup"><span data-stu-id="5f852-101">In this step of the tutorial, you'll add text to the title slide that contains the [Bing](https://www.bing.com) photo of the day.</span></span>

> [!NOTE]
> <span data-ttu-id="5f852-102">このページでは、PowerPoint アドインのチュートリアルの個々の手順について説明します。</span><span class="sxs-lookup"><span data-stu-id="5f852-102">This page describes an individual step of the PowerPoint add-in tutorial.</span></span> <span data-ttu-id="5f852-103">このページに検索エンジンの結果から、または直接リンクからアクセスした場合は、「[PowerPoint アドインのチュートリアル](../tutorials/powerpoint-tutorial.yml)」の紹介ページに移動し、チュートリアルを最初から始めてください。</span><span class="sxs-lookup"><span data-stu-id="5f852-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [PowerPoint add-in tutorial](../tutorials/powerpoint-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="add-text-to-a-slide"></a><span data-ttu-id="5f852-104">スライドにテキストを追加する</span><span class="sxs-lookup"><span data-stu-id="5f852-104">Add text to a slide</span></span> 

1. <span data-ttu-id="5f852-105">**Home.html** ファイルで `TODO3` を次のマークアップに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="5f852-105">In the **Home.html** file, replace `TODO3` with the following markup.</span></span> <span data-ttu-id="5f852-106">このマークアップにより、アドインの作業ウィンドウ内に表示される **[テキストの挿入]** ボタンを定義します。</span><span class="sxs-lookup"><span data-stu-id="5f852-106">This markup defines the **Insert Text** button that will appear within the add-in's task pane.</span></span>

    ```html
        <br /><br />
        <button class="ms-Button ms-Button--primary" id="insert-text">
            <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
            <span class="ms-Button-label">Insert Text</span>
            <span class="ms-Button-description">Inserts text into the slide.</span>
        </button>
    ```

2. <span data-ttu-id="5f852-107">**Home.js** ファイルで `TODO4` を次のコードに置き換え、**[テキストの挿入]** ボタンのイベント ハンドラーを割り当てます。</span><span class="sxs-lookup"><span data-stu-id="5f852-107">In the **Home.js** file, replace `TODO4` with the following code to assign the event handler for the **Insert Text** button.</span></span>

    ```js
    $('#insert-text').click(insertText);
    ```

3. <span data-ttu-id="5f852-108">**Home.js** ファイルで `TODO5` を次のコードに置き換え、**insertText** 関数を定義します。</span><span class="sxs-lookup"><span data-stu-id="5f852-108">In the **Home.js** file, replace `TODO5` with the following code to define the **insertText** function.</span></span> <span data-ttu-id="5f852-109">この関数は、現在のスライドにテキストを挿入します。</span><span class="sxs-lookup"><span data-stu-id="5f852-109">This function inserts text into the current slide.</span></span>

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

## <a name="test-the-add-in"></a><span data-ttu-id="5f852-110">アドインのテスト</span><span class="sxs-lookup"><span data-stu-id="5f852-110">Test the add-in</span></span>

1. <span data-ttu-id="5f852-p104">Visual Studio を使用して、アドインをテストします。そのために、`F5` キーを押すか **[開始]** ボタンをクリックして、リボンに **[作業ウィンドウの表示]** アドイン ボタンが表示された PowerPoint を起動します。アドインは IIS 上でローカルにホストされます。</span><span class="sxs-lookup"><span data-stu-id="5f852-p104">Using Visual Studio, test the add-in by pressing `F5` or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

    ![[開始] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="5f852-114">PowerPoint でリボンの **[作業ウィンドウの表示]** ボタンをクリックし、アドインの作業ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="5f852-114">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![[ホーム] リボンで [作業ウィンドウの表示] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-show-taskpane-button.png)

3. <span data-ttu-id="5f852-116">作業ウィンドウで **[イメージの挿入]** ボタンをクリックしてその日の Bing 写真を現在のスライドに追加し、そのタイトルにテキスト ボックスが含まれるデザインをそのスライドに選択します。</span><span class="sxs-lookup"><span data-stu-id="5f852-116">In the task pane, choose the **Insert Image** button to add the Bing photo of the day to the current slide and choose a design for the slide that contains a text box for the title.</span></span>

    ![[イメージの挿入] ボタンが強調表示されている PowerPoint アドインのスクリーンショット](../images/powerpoint-tutorial-insert-image-slide-design.png)

4. <span data-ttu-id="5f852-118">タイトル スライドのテキスト ボックスにカーソルを置き、作業ウィンドウで **[テキストの挿入]** ボタンをクリックしてテキストをスライドに追加します。</span><span class="sxs-lookup"><span data-stu-id="5f852-118">Put your cursor in the text box on the title slide and then in the task pane, choose the **Insert Text** button to add text to the slide.</span></span>

    ![[テキストの挿入] ボタンが強調表示されている PowerPoint アドインのスクリーンショット](../images/powerpoint-tutorial-insert-text.png)


5. <span data-ttu-id="5f852-120">Visual Studio で `Shift + F5` を押すか **[停止]** ボタンを選択してアドインを停止します。</span><span class="sxs-lookup"><span data-stu-id="5f852-120">In Visual Studio, stop the add-in by pressing `Shift + F5` or choosing the **Stop** button.</span></span> <span data-ttu-id="5f852-121">アドインが停止すると、PowerPoint は自動的に閉じます。</span><span class="sxs-lookup"><span data-stu-id="5f852-121">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![[停止] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-stop.png)