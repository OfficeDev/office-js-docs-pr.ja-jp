# <a name="build-your-first-onenote-add-in"></a><span data-ttu-id="24154-101">最初の OneNote アドインをビルドする</span><span class="sxs-lookup"><span data-stu-id="24154-101">Build your first OneNote add-in</span></span>

<span data-ttu-id="24154-102">この記事では、jQuery と Office JavaScript API を使用して OneNote アドインを作成する手順について説明します。</span><span class="sxs-lookup"><span data-stu-id="24154-102">In this article, you'll walk through the process of building a OneNote add-in by using jQuery and the Office JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="24154-103">前提条件</span><span class="sxs-lookup"><span data-stu-id="24154-103">Prerequisites</span></span>

- [<span data-ttu-id="24154-104">Node.js</span><span class="sxs-lookup"><span data-stu-id="24154-104">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="24154-105">[Yeoman](https://github.com/yeoman/yo) の最新バージョンと [Office アドイン用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)をグローバルにインストールします。</span><span class="sxs-lookup"><span data-stu-id="24154-105">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-add-in-project"></a><span data-ttu-id="24154-106">アドイン プロジェクトの作成</span><span class="sxs-lookup"><span data-stu-id="24154-106">Create the add-in project</span></span>

1. <span data-ttu-id="24154-107">ローカル ドライブにフォルダーを作成し、`my-onenote-addin` という名前を付けます。</span><span class="sxs-lookup"><span data-stu-id="24154-107">Create a folder on your local drive and name it `my-onenote-addin`.</span></span> <span data-ttu-id="24154-108">ここにアドインのファイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="24154-108">This is where you'll create the files for your add-in.</span></span>

2. <span data-ttu-id="24154-109">新しいフォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="24154-109">Navigate to your new folder.</span></span>

    ```bash
    cd my-onenote-addin
    ```

3. <span data-ttu-id="24154-110">Yeoman ジェネレーターを使用して、OneNote アドイン プロジェクトを作成します。</span><span class="sxs-lookup"><span data-stu-id="24154-110">Use the Yeoman generator to create a OneNote add-in project.</span></span> <span data-ttu-id="24154-111">次のコマンドを実行し、以下のプロンプトに応答します。</span><span class="sxs-lookup"><span data-stu-id="24154-111">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="24154-112">**Choose a project type:​ (プロジェクト タイプを選択してください)** `Jquery`</span><span class="sxs-lookup"><span data-stu-id="24154-112">**Choose a project type:** `Jquery`</span></span>
    - <span data-ttu-id="24154-113">**Choose a script type: (スクリプト タイプを選択してください)** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="24154-113">**Choose a script type:** `Javascript`</span></span>
    - <span data-ttu-id="24154-114">**What would you want to name your add-in?: (アドインの名前を何にしますか)** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="24154-114">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="24154-115">**Which Office client application would you like to support?: (どの Office クライアント アプリケーションをサポートしますか)** `Onenote`</span><span class="sxs-lookup"><span data-stu-id="24154-115">**Which Office client application would you like to support?:** `Onenote`</span></span>

    ![Yeoman ジェネレーターのプロンプトと応答のスクリーンショット](../images/yo-office-onenote-jquery.png)
    
    <span data-ttu-id="24154-117">ウィザードが完了すると、ジェネレーターはプロジェクトを作成し、サポートする Node コンポーネントをインストールします。</span><span class="sxs-lookup"><span data-stu-id="24154-117">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>


## <a name="update-the-code"></a><span data-ttu-id="24154-118">コードを更新する</span><span class="sxs-lookup"><span data-stu-id="24154-118">Update the code</span></span>

1. <span data-ttu-id="24154-119">コード エディターで、プロジェクトのルートにある **index.html** を開きます。</span><span class="sxs-lookup"><span data-stu-id="24154-119">In your code editor, open **index.html** in the root of the project.</span></span> <span data-ttu-id="24154-120">このファイルには、アドインの作業ウィンドウにレンダリングされる HTML が含まれています。</span><span class="sxs-lookup"><span data-stu-id="24154-120">This file contains the HTML that will be rendered in the add-in's task pane.</span></span>

2. <span data-ttu-id="24154-121">要素内の `<main>` 要素を次のマークアップに置き換えて、ファイルを保存します。`<body>`</span><span class="sxs-lookup"><span data-stu-id="24154-121">Replace the `<main>` element inside the `<body>` element with the following markup and save the file.</span></span> <span data-ttu-id="24154-122">これは、[Office UI Fabric コンポーネント](http://dev.office.com/fabric/components)を使用してテキスト領域とボタンを追加します。</span><span class="sxs-lookup"><span data-stu-id="24154-122">This adds a text area and a button using [Office UI Fabric components](http://dev.office.com/fabric/components).</span></span>

    ```html
    <main class="ms-welcome__main">
        <br />
        <p class="ms-font-l">Enter content below</p>
        <div class="ms-TextField ms-TextField--placeholder">
            <textarea id="textBox" rows="5"></textarea>
        </div>
        <button id="addOutline" class="ms-welcome__action ms-Button ms-Button--hero ms-u-slideUpIn20">
            <span class="ms-Button-label">Add Outline</span>
            <span class="ms-Button-icon"><i class="ms-Icon"></i></span>
            <span class="ms-Button-description">Adds the content above to the current page.</span>
        </button>
    </main>
    ```

3. <span data-ttu-id="24154-123">**src\index.js** ファイルを開いて、アドインのスクリプトを特定します。</span><span class="sxs-lookup"><span data-stu-id="24154-123">Open the file **app.js** to specify the script for the add-in.</span></span> <span data-ttu-id="24154-124">すべての内容を次のコードに置き換え、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="24154-124">Replace the entire contents with the following code and save the file.</span></span>

    ```js
    'use strict';

    (function () {

        Office.initialize = function (reason) {
            $(document).ready(function () {
                // Set up event handler for the UI.
                $('#addOutline').click(addOutlineToPage);
            });
        };

        // Add the contents of the text area to the page.
        function addOutlineToPage() {        
            OneNote.run(function (context) {
                var html = '<p>' + $('#textBox').val() + '</p>';

                // Get the current page.
                var page = context.application.getActivePage();

                // Queue a command to load the page with the title property.             
                page.load('title'); 

                // Add an outline with the specified HTML to the page.
                var outline = page.addOutline(40, 90, html);

                // Run the queued commands, and return a promise to indicate task completion.
                return context.sync()
                    .then(function() {
                        console.log('Added outline to page ' + page.title);
                    })
                    .catch(function(error) {
                        app.showNotification("Error: " + error); 
                        console.log("Error: " + error); 
                        if (error instanceof OfficeExtension.Error) { 
                            console.log("Debug info: " + JSON.stringify(error.debugInfo)); 
                        } 
                    }); 
            });
        }
    })();
    ```

## <a name="update-the-manifest"></a><span data-ttu-id="24154-125">マニフェストを更新する</span><span class="sxs-lookup"><span data-stu-id="24154-125">Update the manifest</span></span>

1. <span data-ttu-id="24154-126">**one-note-add-in-manifest.xml** ファイルを開いて、アドインの設定と機能を定義します。</span><span class="sxs-lookup"><span data-stu-id="24154-126">Open the file **one-note-add-in-manifest.xml** to define the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="24154-127">要素にはプレースホルダー値が含まれています。`ProviderName`</span><span class="sxs-lookup"><span data-stu-id="24154-127">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="24154-128">それを自分の名前に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="24154-128">Replace it with your name.</span></span>

3. <span data-ttu-id="24154-129">要素の `DefaultValue` 属性にはプレースホルダー値が含まれています。`Description`</span><span class="sxs-lookup"><span data-stu-id="24154-129">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="24154-130">これは、**A task pane add-in for OneNote** に置き換えてください。</span><span class="sxs-lookup"><span data-stu-id="24154-130">Replace it with **A task pane add-in for OneNote**.</span></span>

4. <span data-ttu-id="24154-131">ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="24154-131">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="OneNote Add-in" />
    <Description DefaultValue="A task pane add-in for OneNote"/>
    ...
    ```

## <a name="start-the-dev-server"></a><span data-ttu-id="24154-132">開発用サーバーの起動</span><span class="sxs-lookup"><span data-stu-id="24154-132">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)]

## <a name="try-it-out"></a><span data-ttu-id="24154-133">試してみる</span><span class="sxs-lookup"><span data-stu-id="24154-133">Try it out</span></span>

1. <span data-ttu-id="24154-134">[OneNote Online](https://www.onenote.com/notebooks) でノートブックを開きます。</span><span class="sxs-lookup"><span data-stu-id="24154-134">In [OneNote Online](https://www.onenote.com/notebooks), open a notebook.</span></span>

2. <span data-ttu-id="24154-135">**[挿入] > [Office アドイン]** の順に選択し、[Office アドイン] ダイアログを開きます。</span><span class="sxs-lookup"><span data-stu-id="24154-135">Choose **Insert > Office Add-ins** to open the Office Add-ins dialog.</span></span>

    - <span data-ttu-id="24154-136">コンシューマー アカウントでサインインしている場合は、**[マイ アドイン]** タブを選択し、**[マイ アドインのアップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="24154-136">If you're signed in with your consumer account, select the **MY ADD-INS** tab, and then choose **Upload My Add-in**.</span></span>

    - <span data-ttu-id="24154-137">職場または学校アカウントでサインインしている場合は、**[自分の所属組織]** タブを選択し、**[マイ アドインのアップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="24154-137">If you're signed in with your work or school account, select the **MY ORGANIZATION** tab, and then select **Upload My Add-in**.</span></span> 

    <span data-ttu-id="24154-138">次の図は、コンシューマー ノートブックの **[マイ アドイン]** タブを示しています。</span><span class="sxs-lookup"><span data-stu-id="24154-138">The following image shows the **MY ADD-INS** tab for consumer notebooks.</span></span>

    <img alt="The Office Add-ins dialog showing the MY ADD-INS tab" src="../images/onenote-office-add-ins-dialog.png" width="500">

3. <span data-ttu-id="24154-139">[アドインのアップロード] ダイアログで、プロジェクト フォルダー内の **one-note-add-in-manifest.xml** を参照し、**[アップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="24154-139">In the Upload Add-in dialog, browse to **one-note-add-in-manifest.xml** in your project folder, and then choose **Upload**.</span></span> 

4. <span data-ttu-id="24154-140">**ホーム**タブから、リボンの**タスクペインを表示**ボタンを選択します。</span><span class="sxs-lookup"><span data-stu-id="24154-140">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span> <span data-ttu-id="24154-141">アドインは、OneNote ページの横にある iFrame で開きます。</span><span class="sxs-lookup"><span data-stu-id="24154-141">6- The add-in opens in an iFrame next to the OneNote page.</span></span>

5. <span data-ttu-id="24154-142">テキスト領域にテキストを入力し、**[枠線の追加]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="24154-142">Enter some text in the text area and then choose **Add outline**.</span></span> <span data-ttu-id="24154-143">入力したテキストは、ページに追加されます。</span><span class="sxs-lookup"><span data-stu-id="24154-143">The text you entered is added to the page.</span></span> 

    ![このチュートリアルでビルドした OneNote アドイン](../images/onenote-first-add-in.png)

## <a name="troubleshooting-and-tips"></a><span data-ttu-id="24154-145">トラブルシューティングとヒント</span><span class="sxs-lookup"><span data-stu-id="24154-145">Troubleshooting and tips</span></span>

- <span data-ttu-id="24154-p110">ブラウザーの開発者ツールを使ってアドインをデバッグできます。Gulp Web サーバーを使っており、Internet Explorer や Chrome でデバッグしている場合は、ローカルで変更を保存して、アドインの iFrame を更新するだけです。</span><span class="sxs-lookup"><span data-stu-id="24154-p110">You can debug the add-in using your browser's developer tools. When you're using the Gulp web server and debugging in Internet Explorer or Chrome, you can save your changes locally and then just refresh the add-in's iFrame.</span></span>

- <span data-ttu-id="24154-p111">OneNote オブジェクトを調べる場合、現在使用可能なプロパティに実際の値が表示されます。読み込む必要のあるプロパティには、*undefined* と表示されます。`_proto_` ノードを展開し、オブジェクトで定義されているものの、まだ読み込まれていないプロパティを確認します。</span><span class="sxs-lookup"><span data-stu-id="24154-p111">When you inspect a OneNote object, the properties that are currently available for use display actual values. Properties that need to be loaded display *undefined*. Expand the `_proto_` node to see properties that are defined on the object but are not yet loaded.</span></span>

   ![デバッガーでアンロードされた OneNote オブジェクト](../images/onenote-debug.png)

- <span data-ttu-id="24154-p112">アドインで任意の HTTP リソースを使っている場合は、ブラウザーで混在したコンテンツを有効にする必要があります。運用アドインでは、セキュリティで保護された HTTPS リソースのみを使う必要があります。</span><span class="sxs-lookup"><span data-stu-id="24154-p112">You need to enable mixed content in the browser if your add-in uses any HTTP resources. Production add-ins should use only secure HTTPS resources.</span></span>

- <span data-ttu-id="24154-154">作業ウィンドウ アドインは、任意の場所から開くことができますが、コンテンツアドインは、通常のページ コンテンツ (タイトル、イメージ、iframe などは含まない) の内部にのみ挿入できます。</span><span class="sxs-lookup"><span data-stu-id="24154-154">Task pane add-ins can be opened from anywhere, but content add-ins can only be inserted inside regular page content (i.e. not in titles, images, iFrames, etc.).</span></span> 

## <a name="next-steps"></a><span data-ttu-id="24154-155">次の手順</span><span class="sxs-lookup"><span data-stu-id="24154-155">Next steps</span></span>

<span data-ttu-id="24154-156">これで完了です。OneNote アドインが正常に作成されました。</span><span class="sxs-lookup"><span data-stu-id="24154-156">Congratulations, you've successfully created a OneNote add-in!</span></span> <span data-ttu-id="24154-157">次に、OneNote アドイン構築の中心概念の詳細について説明します。</span><span class="sxs-lookup"><span data-stu-id="24154-157">Next, learn more about the core concepts of building OneNote add-ins.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="24154-158">OneNote の JavaScript API のプログラミングの概要</span><span class="sxs-lookup"><span data-stu-id="24154-158">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)

## <a name="see-also"></a><span data-ttu-id="24154-159">関連項目</span><span class="sxs-lookup"><span data-stu-id="24154-159">See also</span></span>

- [<span data-ttu-id="24154-160">OneNote の JavaScript API のプログラミングの概要</span><span class="sxs-lookup"><span data-stu-id="24154-160">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="24154-161">OneNote JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="24154-161">OneNote JavaScript API reference</span></span>](https://dev.office.com/reference/add-ins/onenote/onenote-add-ins-javascript-reference)
- [<span data-ttu-id="24154-162">Rubric Grader のサンプル</span><span class="sxs-lookup"><span data-stu-id="24154-162">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="24154-163">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="24154-163">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
