---
title: 最初の OneNote アドインをビルドする
description: ''
ms.date: 01/17/2019
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: a0b2820f33e3a7cd31c12aec017ca552575a3f9b
ms.sourcegitcommit: 33dcf099c6b3d249811580d67ee9b790c0fdccfb
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/05/2019
ms.locfileid: "29742339"
---
# <a name="build-your-first-onenote-add-in"></a><span data-ttu-id="71c99-102">最初の OneNote アドインをビルドする</span><span class="sxs-lookup"><span data-stu-id="71c99-102">Build your first OneNote add-in</span></span>

<span data-ttu-id="71c99-103">この記事では、jQuery と Office JavaScript API を使用して OneNote アドインを作成する手順について説明します。</span><span class="sxs-lookup"><span data-stu-id="71c99-103">In this article, you'll walk through the process of building a OneNote add-in by using jQuery and the Office JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="71c99-104">前提条件</span><span class="sxs-lookup"><span data-stu-id="71c99-104">Prerequisites</span></span>

- [<span data-ttu-id="71c99-105">Node.js</span><span class="sxs-lookup"><span data-stu-id="71c99-105">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="71c99-106">[Yeoman](https://github.com/yeoman/yo) の最新バージョンと [Office アドイン用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)をグローバルにインストールします。</span><span class="sxs-lookup"><span data-stu-id="71c99-106">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-add-in-project"></a><span data-ttu-id="71c99-107">アドイン プロジェクトの作成</span><span class="sxs-lookup"><span data-stu-id="71c99-107">Create the add-in project</span></span>

1. <span data-ttu-id="71c99-108">Yeoman ジェネレーターを使用して、OneNote アドイン プロジェクトを作成します。</span><span class="sxs-lookup"><span data-stu-id="71c99-108">Use the Yeoman generator to create a OneNote add-in project.</span></span> <span data-ttu-id="71c99-109">次のコマンドを実行し、以下のプロンプトに応答します。</span><span class="sxs-lookup"><span data-stu-id="71c99-109">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="71c99-110">**Choose a project type: (プロジェクトの種類を選択)** `Office Add-in project using Jquery framework`</span><span class="sxs-lookup"><span data-stu-id="71c99-110">**Choose a project type:** `Office Add-in project using Jquery framework`</span></span>
    - <span data-ttu-id="71c99-111">**Choose a script type: (スクリプトの種類を選択)** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="71c99-111">**Choose a script type:** `Javascript`</span></span>
    - <span data-ttu-id="71c99-112">**What would you want to name your add-in?: (アドインの名前を何にしますか)** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="71c99-112">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="71c99-113">**Which Office client application would you like to support?: (どの Office クライアント アプリケーションをサポートしますか)** `Onenote`</span><span class="sxs-lookup"><span data-stu-id="71c99-113">**Which Office client application would you like to support?:** `Onenote`</span></span>

    ![Yeoman ジェネレーターのプロンプトと応答のスクリーンショット](../images/yo-office-onenote-jquery.png)
    
    <span data-ttu-id="71c99-115">ウィザードを完了すると、ジェネレーターによってプロジェクトが作成されて、サポートしているノード コンポーネントがインストールされます。</span><span class="sxs-lookup"><span data-stu-id="71c99-115">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>
    
2. <span data-ttu-id="71c99-116">プロジェクトのルート フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="71c99-116">Navigate to the root folder of the project.</span></span>

    ```bash
    cd "My Office Add-in"
    ```

## <a name="update-the-code"></a><span data-ttu-id="71c99-117">コードを更新する</span><span class="sxs-lookup"><span data-stu-id="71c99-117">Update the code</span></span>

1. <span data-ttu-id="71c99-118">コード エディターで、プロジェクトのルートにある **index.html** を開きます。</span><span class="sxs-lookup"><span data-stu-id="71c99-118">In your code editor, open **index.html** in the root of the project.</span></span> <span data-ttu-id="71c99-119">このファイルには、アドインの作業ウィンドウにレンダリングされる HTML が含まれています。</span><span class="sxs-lookup"><span data-stu-id="71c99-119">This file contains the HTML that will be rendered in the add-in's task pane.</span></span>

2. <span data-ttu-id="71c99-120">`<body>` 要素を次のマークアップに置き換え、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="71c99-120">Replace the `<body>` element with the following markup and save the file.</span></span> 

    ```html
    <body class="ms-font-m ms-welcome">
        <header class="ms-welcome__header ms-bgColor-themeDark ms-u-fadeIn500">
            <h2 class="ms-fontSize-xxl ms-fontWeight-regular ms-fontColor-white">OneNote Add-in</h1>
        </header>
        <main id="app-body" class="ms-welcome__main">
            <br />
            <p class="ms-font-m">Enter HTML content here:</p>
            <div class="ms-TextField ms-TextField--placeholder">
                <textarea id="textBox" rows="8" cols="30"></textarea>
            </div>
            <button id="addOutline" class="ms-Button ms-Button--primary">
                <span class="ms-Button-label">Add outline</span>
            </button>
        </main>
        <script type="text/javascript" src="node_modules/jquery/dist/jquery.js"></script>
        <script type="text/javascript" src="node_modules/office-ui-fabric-js/dist/js/fabric.js"></script>
    </body>
    ```

3. <span data-ttu-id="71c99-121">**src\index.js** ファイルを開いて、アドインのスクリプトを指定します。</span><span class="sxs-lookup"><span data-stu-id="71c99-121">Open the file **src\index.js** to specify the script for the add-in.</span></span> <span data-ttu-id="71c99-122">すべての内容を次のコードに置き換え、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="71c99-122">Replace the entire contents with the following code and save the file.</span></span>

    ```js
    import * as OfficeHelpers from "@microsoft/office-js-helpers";

    Office.onReady(() => {
        // Office is ready
        $(document).ready(() => {
            // The document is ready
            $('#addOutline').click(addOutlineToPage);
        });
    });
    
    async function addOutlineToPage() {
        try {
            await OneNote.run(async context => {
                var html = "<p>" + $("#textBox").val() + "</p>";

                // Get the current page.
                var page = context.application.getActivePage();

                // Queue a command to load the page with the title property.
                page.load("title");

                // Add text to the page by using the specified HTML.
                var outline = page.addOutline(40, 90, html);

                // Run the queued commands, and return a promise to indicate task completion.
                return context.sync()
                    .then(function() {
                        console.log("Added outline to page " + page.title);
                    })
                    .catch(function(error) {
                        app.showNotification("Error: " + error);
                        console.log("Error: " + error);
                        if (error instanceof OfficeExtension.Error) {
                            console.log("Debug info: " + JSON.stringify(error.debugInfo));
                        }
                    });
                });
        } catch (error) {
            OfficeHelpers.UI.notify(error);
            OfficeHelpers.Utilities.log(error);
        }
    }
    ```

4. <span data-ttu-id="71c99-123">**app.css** ファイルを開いて、アドインのカスタム スタイルを指定します。</span><span class="sxs-lookup"><span data-stu-id="71c99-123">Open the file **app.css** to specify the custom styles for the add-in.</span></span> <span data-ttu-id="71c99-124">すべての内容を次のものに置き換え、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="71c99-124">Replace the entire contents with the following and save the file.</span></span>

    ```css
    html, body {
        width: 100%;
        height: 100%;
        margin: 0;
        padding: 0;
    }

    ul, p, h1, h2, h3, h4, h5, h6 {
        margin: 0;
        padding: 0;
    }

    .ms-welcome {
        position: relative;
        display: -webkit-flex;
        display: flex;
        -webkit-flex-direction: column;
        flex-direction: column;
        -webkit-flex-wrap: nowrap;
        flex-wrap: nowrap;
        min-height: 500px;
        min-width: 320px;
        overflow: auto;
        overflow-x: hidden;
    }

    .ms-welcome__header {
        min-height: 30px;
        padding: 0px;
        padding-bottom: 5px;
        display: -webkit-flex;
        display: flex;
        -webkit-flex-direction: column;
        flex-direction: column;
        -webkit-flex-wrap: nowrap;
        flex-wrap: nowrap;
        -webkit-align-items: center;
        align-items: center;
        -webkit-justify-content: flex-end;
        justify-content: flex-end;
    }

    .ms-welcome__header > h1 {
        margin-top: 5px;
        text-align: center;
    }

    .ms-welcome__main {
        display: -webkit-flex;
        display: flex;
        -webkit-flex-direction: column;
        flex-direction: column;
        -webkit-flex-wrap: nowrap;
        flex-wrap: nowrap;
        -webkit-align-items: center;
        align-items: left;
        -webkit-flex: 1 0 0;
        flex: 1 0 0;
        padding: 30px 20px;
    }

    .ms-welcome__main > h2 {
        width: 100%;
        text-align: left;
    }

    @media (min-width: 0) and (max-width: 350px) {
        .ms-welcome__features {
            width: 100%;
        }
    }
    ```

## <a name="update-the-manifest"></a><span data-ttu-id="71c99-125">マニフェストを更新する</span><span class="sxs-lookup"><span data-stu-id="71c99-125">Update the manifest</span></span>

1. <span data-ttu-id="71c99-126">**manifest.xml** ファイルを開いて、アドインの設定と機能を定義します。</span><span class="sxs-lookup"><span data-stu-id="71c99-126">Open the file **manifest.xml** to define the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="71c99-127">`ProviderName` 要素にはプレースホルダー値が含まれています。</span><span class="sxs-lookup"><span data-stu-id="71c99-127">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="71c99-128">それを自分の名前に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="71c99-128">Replace it with your name.</span></span>

3. <span data-ttu-id="71c99-129">`Description` 要素の `DefaultValue` 属性にはプレースホルダー値が含まれています。</span><span class="sxs-lookup"><span data-stu-id="71c99-129">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="71c99-130">これは、**A task pane add-in for OneNote** に置き換えてください。</span><span class="sxs-lookup"><span data-stu-id="71c99-130">Replace it with **A task pane add-in for OneNote**.</span></span>

4. <span data-ttu-id="71c99-131">ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="71c99-131">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for OneNote"/>
    ...
    ```

## <a name="start-the-dev-server"></a><span data-ttu-id="71c99-132">開発用サーバーの起動</span><span class="sxs-lookup"><span data-stu-id="71c99-132">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)]

## <a name="try-it-out"></a><span data-ttu-id="71c99-133">試してみる</span><span class="sxs-lookup"><span data-stu-id="71c99-133">Try it out</span></span>

1. <span data-ttu-id="71c99-134">[OneNote Online](https://www.onenote.com/notebooks) でノートブックを開きます。</span><span class="sxs-lookup"><span data-stu-id="71c99-134">In [OneNote Online](https://www.onenote.com/notebooks), open a notebook.</span></span>

2. <span data-ttu-id="71c99-135">**[挿入] > [Office アドイン]** の順に選択し、[Office アドイン] ダイアログを開きます。</span><span class="sxs-lookup"><span data-stu-id="71c99-135">Choose **Insert > Office Add-ins** to open the Office Add-ins dialog.</span></span>

    - <span data-ttu-id="71c99-136">コンシューマー アカウントでサインインしている場合は、**[マイ アドイン]** タブを選択し、**[マイ アドインのアップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="71c99-136">If you're signed in with your consumer account, select the **MY ADD-INS** tab, and then choose **Upload My Add-in**.</span></span>

    - <span data-ttu-id="71c99-137">職場または学校アカウントでサインインしている場合は、**[自分の所属組織]** タブを選択し、**[マイ アドインのアップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="71c99-137">If you're signed in with your work or school account, select the **MY ORGANIZATION** tab, and then select **Upload My Add-in**.</span></span> 

    <span data-ttu-id="71c99-138">次の図は、コンシューマー ノートブックの **[マイ アドイン]** タブを示しています。</span><span class="sxs-lookup"><span data-stu-id="71c99-138">The following image shows the **MY ADD-INS** tab for consumer notebooks.</span></span>

    <img alt="The Office Add-ins dialog showing the MY ADD-INS tab" src="../images/onenote-office-add-ins-dialog.png" width="500">

3. <span data-ttu-id="71c99-139">[アドインのアップロード] ダイアログで、プロジェクト フォルダー内の **manifest.xml** を参照し、**[アップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="71c99-139">In the Upload Add-in dialog, browse to **manifest.xml** in your project folder, and then choose **Upload**.</span></span> 

4. <span data-ttu-id="71c99-140">**[ホーム]** タブから、リボンの **[作業ウィンドウの表示]** ボタンをクリックします。</span><span class="sxs-lookup"><span data-stu-id="71c99-140">From the **Home** tab, choose the **Show Taskpane** button in the ribbon.</span></span> <span data-ttu-id="71c99-141">アドインの作業ウィンドウは、OneNote ページの横にある iFrame で開きます。</span><span class="sxs-lookup"><span data-stu-id="71c99-141">The add-in task pane opens in an iFrame next to the OneNote page.</span></span>

5. <span data-ttu-id="71c99-142">テキスト エリアに次の HTML コンテンツを入力し、**[アウトラインの追加]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="71c99-142">Enter the following HTML content in the text area, and then choose **Add outline**.</span></span>  

    ```html
    <ol>
    <li>Item #1</li>
    <li>Item #2</li>
    <li>Item #3</li>
    <li>Item #4</li>
    </ol>
    ```

    <span data-ttu-id="71c99-143">指定したアウトラインがページに追加されます。</span><span class="sxs-lookup"><span data-stu-id="71c99-143">The outline that you specified is added to the page.</span></span>

    ![このチュートリアルでビルドした OneNote アドイン](../images/onenote-first-add-in-3.png)

## <a name="troubleshooting-and-tips"></a><span data-ttu-id="71c99-145">トラブルシューティングとヒント</span><span class="sxs-lookup"><span data-stu-id="71c99-145">Troubleshooting and tips</span></span>

- <span data-ttu-id="71c99-p108">ブラウザーの開発者ツールを使ってアドインをデバッグできます。Gulp Web サーバーを使っており、Internet Explorer や Chrome でデバッグしている場合は、ローカルで変更を保存して、アドインの iFrame を更新するだけです。</span><span class="sxs-lookup"><span data-stu-id="71c99-p108">You can debug the add-in using your browser's developer tools. When you're using the Gulp web server and debugging in Internet Explorer or Chrome, you can save your changes locally and then just refresh the add-in's iFrame.</span></span>

- <span data-ttu-id="71c99-p109">OneNote オブジェクトを調べる場合、現在使用可能なプロパティに実際の値が表示されます。読み込む必要のあるプロパティには、*undefined* と表示されます。`_proto_` ノードを展開し、オブジェクトで定義されているものの、まだ読み込まれていないプロパティを確認します。</span><span class="sxs-lookup"><span data-stu-id="71c99-p109">When you inspect a OneNote object, the properties that are currently available for use display actual values. Properties that need to be loaded display *undefined*. Expand the `_proto_` node to see properties that are defined on the object but are not yet loaded.</span></span>

   ![デバッガーでアンロードされた OneNote オブジェクト](../images/onenote-debug.png)

- <span data-ttu-id="71c99-p110">アドインで任意の HTTP リソースを使っている場合は、ブラウザーで混在したコンテンツを有効にする必要があります。運用アドインでは、セキュリティで保護された HTTPS リソースのみを使う必要があります。</span><span class="sxs-lookup"><span data-stu-id="71c99-p110">You need to enable mixed content in the browser if your add-in uses any HTTP resources. Production add-ins should use only secure HTTPS resources.</span></span>

- <span data-ttu-id="71c99-154">作業ウィンドウ アドインは、任意の場所から開くことができますが、コンテンツアドインは、通常のページ コンテンツ (タイトル、イメージ、iframe などは含まない) の内部にのみ挿入できます。</span><span class="sxs-lookup"><span data-stu-id="71c99-154">Task pane add-ins can be opened from anywhere, but content add-ins can only be inserted inside regular page content (i.e. not in titles, images, iFrames, etc.).</span></span> 

## <a name="next-steps"></a><span data-ttu-id="71c99-155">次の手順</span><span class="sxs-lookup"><span data-stu-id="71c99-155">Next steps</span></span>

<span data-ttu-id="71c99-156">これで完了です。OneNote アドインが正常に作成されました。</span><span class="sxs-lookup"><span data-stu-id="71c99-156">Congratulations, you've successfully created a OneNote add-in!</span></span> <span data-ttu-id="71c99-157">次に、OneNote アドイン構築の中心概念の詳細について説明します。</span><span class="sxs-lookup"><span data-stu-id="71c99-157">Next, learn more about the core concepts of building OneNote add-ins.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="71c99-158">OneNote の JavaScript API のプログラミングの概要</span><span class="sxs-lookup"><span data-stu-id="71c99-158">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)

## <a name="see-also"></a><span data-ttu-id="71c99-159">関連項目</span><span class="sxs-lookup"><span data-stu-id="71c99-159">See also</span></span>

- [<span data-ttu-id="71c99-160">OneNote の JavaScript API のプログラミングの概要</span><span class="sxs-lookup"><span data-stu-id="71c99-160">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="71c99-161">OneNote JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="71c99-161">OneNote JavaScript API reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference)
- [<span data-ttu-id="71c99-162">Rubric Grader のサンプル</span><span class="sxs-lookup"><span data-stu-id="71c99-162">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="71c99-163">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="71c99-163">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)

