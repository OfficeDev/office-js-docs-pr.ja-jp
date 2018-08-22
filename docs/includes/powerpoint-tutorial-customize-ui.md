<span data-ttu-id="47fa4-101">このチュートリアルの手順では、作業ウィンドウのユーザー インターフェイス (UI) をカスタマイズします。</span><span class="sxs-lookup"><span data-stu-id="47fa4-101">In this step of the tutorial, you'll customize the task pane user interface (UI).</span></span>

> [!NOTE]
> <span data-ttu-id="47fa4-102">このページでは、PowerPoint アドインのチュートリアルの個々の手順について説明します。</span><span class="sxs-lookup"><span data-stu-id="47fa4-102">This page describes an individual step of the PowerPoint add-in tutorial.</span></span> <span data-ttu-id="47fa4-103">このページに検索エンジンの結果から、または直接リンクからアクセスした場合は、「[PowerPoint アドインのチュートリアル](../tutorials/powerpoint-tutorial.yml)」の紹介ページに移動し、チュートリアルを最初から始めてください。</span><span class="sxs-lookup"><span data-stu-id="47fa4-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [PowerPoint add-in tutorial](../tutorials/powerpoint-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="customize-the-task-pane-ui"></a><span data-ttu-id="47fa4-104">作業ウィンドウの UI をカスタマイズする</span><span class="sxs-lookup"><span data-stu-id="47fa4-104">Customize the task pane UI</span></span> 

1. <span data-ttu-id="47fa4-105">**Home.html** ファイルで `TODO2` を次のマークアップと置き換え、ヘッダー セクションとタイトルを作業ウィンドウに追加します。</span><span class="sxs-lookup"><span data-stu-id="47fa4-105">In the **Home.html** file, replace `TODO2` with the following markup to add a header section and title to the task pane.</span></span> <span data-ttu-id="47fa4-106">注意:</span><span class="sxs-lookup"><span data-stu-id="47fa4-106">Note:</span></span>

    - <span data-ttu-id="47fa4-107">`ms-` で始まるスタイルは、[Office UI Fabric](../design/office-ui-fabric.md) で定義されています。これは、Office と Office 365 のユーザー エクスペリエンスを構築するための JavaScript フロント エンドのフレームワークです。</span><span class="sxs-lookup"><span data-stu-id="47fa4-107">The styles that begin with `ms-` are defined by [Office UI Fabric](../design/office-ui-fabric.md), a JavaScript front-end framework for building user experiences for Office and Office 365.</span></span> <span data-ttu-id="47fa4-108">**Home.html** ファイルには、Fabric スタイル シートへの参照が含まれています。</span><span class="sxs-lookup"><span data-stu-id="47fa4-108">The **Home.html** file includes a reference to the Fabric stylesheet.</span></span>

    ```html
    <div id="content-header">
        <div class="ms-Grid ms-bgColor-neutralPrimary">
            <div class="ms-Grid-row">
                <div class="padding ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12"> <div class="ms-font-xl ms-fontColor-white ms-fontWeight-semibold">My PowerPoint Add-in</div></div>
            </div>
        </div>
    </div>
    ```

2. <span data-ttu-id="47fa4-109">**Home.html** ファイルにおいて、`class="footer"` で **div** を検索し、**div** 全体を削除して作業ウィンドウからフッター セクションを削除します。</span><span class="sxs-lookup"><span data-stu-id="47fa4-109">In the **Home.html** file, find the **div** with `class="footer"` and delete that entire **div** to remove the footer section from the task pane.</span></span>

## <a name="test-the-add-in"></a><span data-ttu-id="47fa4-110">アドインのテスト</span><span class="sxs-lookup"><span data-stu-id="47fa4-110">Test the add-in</span></span>

1. <span data-ttu-id="47fa4-p104">Visual Studio を使用して、PowerPoint アドインをテストします。そのために、`F5` キーを押すか **[開始]** ボタンをクリックして、リボンに **[作業ウィンドウの表示]** アドイン ボタンが表示された PowerPoint を起動します。アドインは IIS 上でローカルにホストされます。</span><span class="sxs-lookup"><span data-stu-id="47fa4-p104">Using Visual Studio, test the PowerPoint add-in by pressing `F5` or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

    ![[開始] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="47fa4-114">PowerPoint でリボンの **[作業ウィンドウの表示]** ボタンをクリックし、アドインの作業ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="47fa4-114">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![[ホーム] リボンで [作業ウィンドウの表示] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-show-taskpane-button.png)

3. <span data-ttu-id="47fa4-116">このとき、作業ウィンドウにはヘッダー セクションとタイトルが含まれ、フッター セクションが含まれないことがわかります。</span><span class="sxs-lookup"><span data-stu-id="47fa4-116">Notice that the task pane now contains a header section and title, and no longer contains a footer section.</span></span>

    ![[イメージの挿入] ボタンが強調表示されている PowerPoint アドインのスクリーンショット](../images/powerpoint-tutorial-new-task-pane-ui.png)

4. <span data-ttu-id="47fa4-118">Visual Studio で `Shift + F5` を押すか **[停止]** ボタンを選択してアドインを停止します。</span><span class="sxs-lookup"><span data-stu-id="47fa4-118">In Visual Studio, stop the add-in by pressing `Shift + F5` or choosing the **Stop** button.</span></span> <span data-ttu-id="47fa4-119">アドインが停止すると、PowerPoint は自動的に閉じます。</span><span class="sxs-lookup"><span data-stu-id="47fa4-119">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![[停止] ボタンが強調表示されている Visual Studio のスクリーンショット](../images/powerpoint-tutorial-stop.png)

