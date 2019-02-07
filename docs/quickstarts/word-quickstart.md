---
title: 最初の Word アドインをビルドする
description: ''
ms.date: 01/17/2019
ms.prod: word
localization_priority: Priority
ms.openlocfilehash: 20f63c7e61fd4b576eda7c1d737d377439055fe2
ms.sourcegitcommit: 33dcf099c6b3d249811580d67ee9b790c0fdccfb
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/05/2019
ms.locfileid: "29742325"
---
# <a name="build-your-first-word-add-in"></a><span data-ttu-id="da840-102">最初の Word アドインをビルドする</span><span class="sxs-lookup"><span data-stu-id="da840-102">Build your first Word add-in</span></span>

<span data-ttu-id="da840-103">_適用対象: Word 2016、Word for iPad、Word for Mac_</span><span class="sxs-lookup"><span data-stu-id="da840-103">_Applies to: Word 2016, Word for iPad, Word for Mac_</span></span>

<span data-ttu-id="da840-104">この記事では、jQuery と Word JavaScript API を使用して Word アドインを構築する手順について説明します。</span><span class="sxs-lookup"><span data-stu-id="da840-104">In this article, you'll walk through the process of building a Word add-in by using jQuery and the Word JavaScript API.</span></span> 

## <a name="create-the-add-in"></a><span data-ttu-id="da840-105">アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="da840-105">Create the add-in</span></span> 

[!include[Choose your editor](../includes/quickstart-choose-editor.md)]

# <a name="visual-studiotabvisual-studio"></a>[<span data-ttu-id="da840-106">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="da840-106">Visual Studio</span></span>](#tab/visual-studio)

### <a name="prerequisites"></a><span data-ttu-id="da840-107">前提条件</span><span class="sxs-lookup"><span data-stu-id="da840-107">Prerequisites</span></span>

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

### <a name="create-the-add-in-project"></a><span data-ttu-id="da840-108">アドイン プロジェクトの作成</span><span class="sxs-lookup"><span data-stu-id="da840-108">Create the add-in project</span></span>

1. <span data-ttu-id="da840-109">[Visual Studio] メニュー バーで、**[ファイル]**  >  **[新規作成]**  >  **[プロジェクト]** の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="da840-109">On the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>
    
2. <span data-ttu-id="da840-110">**[Visual C#]** または **[Visual Basic]** の下にあるプロジェクトの種類の一覧で、**[Office/SharePoint]** を展開して、**[アドイン]** を選択し、プロジェクトの種類として **[Word Web アドイン]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="da840-110">In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **Word Web Add-in** as the project type.</span></span> 

3. <span data-ttu-id="da840-111">プロジェクトに名前を付けて、**[OK]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="da840-111">Name the project, and then choose **OK**.</span></span>

4. <span data-ttu-id="da840-p101">Visual Studio によってソリューションとその 2 つのプロジェクトが作成され、**ソリューション エクスプローラー**に表示されます。**Home.html** ファイルが Visual Studio で開かれます。</span><span class="sxs-lookup"><span data-stu-id="da840-p101">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>
    
### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="da840-114">Visual Studio ソリューションについて理解する</span><span class="sxs-lookup"><span data-stu-id="da840-114">Explore the Visual Studio solution</span></span>

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-the-code"></a><span data-ttu-id="da840-115">コードを更新する</span><span class="sxs-lookup"><span data-stu-id="da840-115">Update the code</span></span>

1. <span data-ttu-id="da840-116">**Home.html** では、アドインの作業ウィンドウにレンダリングされる HTML を指定します。</span><span class="sxs-lookup"><span data-stu-id="da840-116">**Home.html** specifies the HTML that will be rendered in the add-in's task pane.</span></span> <span data-ttu-id="da840-117">**Home.html** で、`<body>` 要素を次のマークアップに置き換えて、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="da840-117">In **Home.html**, replace the `<body>` element with the following markup and save the file.</span></span>
 
    ```html
    <body>
        <div id="content-header">
            <div class="padding">
                <h1>Welcome</h1>
            </div>
        </div>    
        <div id="content-main">
            <div class="padding">
                <p>Choose the buttons below to add boilerplate text to the document by using the Word JavaScript API.</p>
                <br />
                <h3>Try it out</h3>
                <button id="emerson">Add quote from Ralph Waldo Emerson</button>
                <br /><br />
                <button id="checkhov">Add quote from Anton Chekhov</button>
                <br /><br />
                <button id="proverb">Add Chinese proverb</button>
            </div>
        </div>
        <br />
        <div id="supportedVersion"/>
    </body>
    ```

2. <span data-ttu-id="da840-118">Web アプリケーション プロジェクトのルートにあるファイル **Home.js** を開きます。</span><span class="sxs-lookup"><span data-stu-id="da840-118">Open the file **Home.js** in the root of the web application project.</span></span> <span data-ttu-id="da840-119">このファイルは、アドイン用のスクリプトを指定します。</span><span class="sxs-lookup"><span data-stu-id="da840-119">This file specifies the script for the add-in.</span></span> <span data-ttu-id="da840-120">すべての内容を次のコードに置き換え、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="da840-120">Replace the entire contents with the following code and save the file.</span></span>

    ```js
    'use strict';
    
    (function () {

        Office.onReady(function() {
            // Office is ready
            $(document).ready(function () {
                // The document is ready
                // Use this to check whether the API is supported in the Word client.
                if (Office.context.requirements.isSetSupported('WordApi', 1.1)) {
                    // Do something that is only available via the new APIs
                    $('#emerson').click(insertEmersonQuoteAtSelection);
                    $('#checkhov').click(insertChekhovQuoteAtTheBeginning);
                    $('#proverb').click(insertChineseProverbAtTheEnd);
                    $('#supportedVersion').html('This code is using Word 2016 or later.');
                }
                else {
                    // Just letting you know that this code will not work with your version of Word.
                    $('#supportedVersion').html('This code requires Word 2016 or later.');
                }
            });
        });

        function insertEmersonQuoteAtSelection() {
            Word.run(function (context) {

                // Create a proxy object for the document.
                var thisDocument = context.document;

                // Queue a command to get the current selection.
                // Create a proxy range object for the selection.
                var range = thisDocument.getSelection();

                // Queue a command to replace the selected text.
                range.insertText('"Hitch your wagon to a star."\n', Word.InsertLocation.replace);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from Ralph Waldo Emerson.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }

        function insertChekhovQuoteAtTheBeginning() {
            Word.run(function (context) {

                // Create a proxy object for the document body.
                var body = context.document.body;

                // Queue a command to insert text at the start of the document body.
                body.insertText('"Knowledge is of no value unless you put it into practice."\n', Word.InsertLocation.start);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from Anton Chekhov.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }

        function insertChineseProverbAtTheEnd() {
            Word.run(function (context) {

                // Create a proxy object for the document body.
                var body = context.document.body;

                // Queue a command to insert text at the end of the document body.
                body.insertText('"To know the road ahead, ask those coming back."\n', Word.InsertLocation.end);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from a Chinese proverb.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
    ```

3. <span data-ttu-id="da840-121">Web アプリケーション プロジェクトのルートにあるファイル **Home.css** を開きます。</span><span class="sxs-lookup"><span data-stu-id="da840-121">Open the file **Home.css** in the root of the web application project.</span></span> <span data-ttu-id="da840-122">このファイルは、アドイン用のユーザー設定のスタイルを指定します。</span><span class="sxs-lookup"><span data-stu-id="da840-122">This file specifies the custom styles for the add-in.</span></span> <span data-ttu-id="da840-123">すべての内容を次のコードに置き換え、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="da840-123">Replace the entire contents with the following code and save the file.</span></span>

    ```css
    #content-header {
        background: #2a8dd4;
        color: #fff;
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 80px; 
        overflow: hidden;
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px;
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto; 
    }

    .padding {
        padding: 15px;
    }
    ```

### <a name="update-the-manifest"></a><span data-ttu-id="da840-124">マニフェストを更新する</span><span class="sxs-lookup"><span data-stu-id="da840-124">Update the manifest</span></span>

1. <span data-ttu-id="da840-125">アドイン プロジェクト内の XML マニフェスト ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="da840-125">Open the XML manifest file in the add-in project.</span></span> <span data-ttu-id="da840-126">このファイルは、アドインの設定と機能を定義します。</span><span class="sxs-lookup"><span data-stu-id="da840-126">This file defines the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="da840-127">`ProviderName` 要素にはプレースホルダー値が含まれています。</span><span class="sxs-lookup"><span data-stu-id="da840-127">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="da840-128">それを自分の名前に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="da840-128">Replace it with your name.</span></span>

3. <span data-ttu-id="da840-129">`DisplayName` 要素の `DefaultValue` 属性にはプレースホルダー値が含まれています。</span><span class="sxs-lookup"><span data-stu-id="da840-129">The `DefaultValue` attribute of the `DisplayName` element has a placeholder.</span></span> <span data-ttu-id="da840-130">これは、**My Office Add-in** に置き換えてください。</span><span class="sxs-lookup"><span data-stu-id="da840-130">Replace it with **My Office Add-in**.</span></span>

4. <span data-ttu-id="da840-131">`Description` 要素の `DefaultValue` 属性にはプレースホルダー値が含まれています。</span><span class="sxs-lookup"><span data-stu-id="da840-131">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="da840-132">これは、**A task pane add-in for Word** に置き換えてください。</span><span class="sxs-lookup"><span data-stu-id="da840-132">Replace it with **A task pane add-in for Word**.</span></span>

5. <span data-ttu-id="da840-133">ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="da840-133">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Word"/>
    ...
    ```

### <a name="try-it-out"></a><span data-ttu-id="da840-134">試してみる</span><span class="sxs-lookup"><span data-stu-id="da840-134">Try it out</span></span>

1. <span data-ttu-id="da840-p109">Visual Studio を使用して、新しく作成した Word アドインをテストします。そのために、**F5** キーを押すか **[開始]** ボタンをクリックして、リボンに **[作業ウィンドウの表示]** アドイン ボタンが表示された Word を起動します。アドインは IIS 上でローカルにホストされます。</span><span class="sxs-lookup"><span data-stu-id="da840-p109">Using Visual Studio, test the newly created Word add-in by pressing **F5** or choosing the **Start** button to launch Word with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="da840-137">Word で、**[ホーム]** タブを選択し、リボンの **[作業ウィンドウの表示]** ボタンをクリックして、アドインの作業ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="da840-137">In Word, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span> <span data-ttu-id="da840-138">(Office 365 バージョンではなく、サブスクリプション バージョンではない Office 2016 を使用している場合、カスタム ボタンはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="da840-138">(If you are using the non-subscription version of Office 2016, instead of the Office 365 version, then custom buttons are not supported.</span></span> <span data-ttu-id="da840-139">代わりに、作業ウィンドウがそのまま開きます)。</span><span class="sxs-lookup"><span data-stu-id="da840-139">Instead, the task pane will open immediately.)</span></span>

    ![[作業ウィンドウの表示] ボタンが強調表示されている Word アプリケーションのスクリーンショット](../images/word-quickstart-addin-0.png)

3. <span data-ttu-id="da840-141">作業ウィンドウで、いずれかのボタンを選択して文書に定型句を追加します。</span><span class="sxs-lookup"><span data-stu-id="da840-141">In the task pane, choose any of the buttons to add boilerplate text to the document.</span></span>

    ![定型句アドインが読み込まれている Word アプリケーションのスクリーンショット。](../images/word-quickstart-addin-1b.png)

# <a name="any-editortabvisual-studio-code"></a>[<span data-ttu-id="da840-143">任意のエディター</span><span class="sxs-lookup"><span data-stu-id="da840-143">Any editor</span></span>](#tab/visual-studio-code)

### <a name="prerequisites"></a><span data-ttu-id="da840-144">前提条件</span><span class="sxs-lookup"><span data-stu-id="da840-144">Prerequisites</span></span>

- [<span data-ttu-id="da840-145">Node.js</span><span class="sxs-lookup"><span data-stu-id="da840-145">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="da840-146">[Yeoman](https://github.com/yeoman/yo) の最新バージョンと [Office アドイン用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)をグローバルにインストールします。</span><span class="sxs-lookup"><span data-stu-id="da840-146">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

### <a name="create-the-add-in-project"></a><span data-ttu-id="da840-147">アドイン プロジェクトの作成</span><span class="sxs-lookup"><span data-stu-id="da840-147">Create the add-in project</span></span>

1. <span data-ttu-id="da840-148">Yeoman ジェネレーターを使用して、Word アドイン プロジェクトを作成します。</span><span class="sxs-lookup"><span data-stu-id="da840-148">Use the Yeoman generator to create a Word add-in project.</span></span> <span data-ttu-id="da840-149">次のコマンドを実行し、以下のプロンプトに応答します。</span><span class="sxs-lookup"><span data-stu-id="da840-149">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="da840-150">**Choose a project type: (プロジェクトの種類を選択)** `Office Add-in project using Jquery framework`</span><span class="sxs-lookup"><span data-stu-id="da840-150">**Choose a project type:** `Office Add-in project using Jquery framework`</span></span>
    - <span data-ttu-id="da840-151">**Choose a script type: (スクリプトの種類を選択)** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="da840-151">**Choose a script type:** `Javascript`</span></span>
    - <span data-ttu-id="da840-152">**What would you want to name your add-in?: (アドインの名前を何にしますか)** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="da840-152">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="da840-153">**Which Office client application would you like to support?: (どの Office クライアント アプリケーションをサポートしますか)** `Word`</span><span class="sxs-lookup"><span data-stu-id="da840-153">**Which Office client application would you like to support?:** `Word`</span></span>

    ![Yeoman ジェネレーターのプロンプトと応答のスクリーンショット](../images/yo-office-word-jquery.png)
    
    <span data-ttu-id="da840-155">ウィザードを完了すると、ジェネレーターによってプロジェクトが作成されて、サポートしているノード コンポーネントがインストールされます。</span><span class="sxs-lookup"><span data-stu-id="da840-155">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>
    
2. <span data-ttu-id="da840-156">プロジェクトのルート フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="da840-156">Navigate to the root folder of the project.</span></span>

    ```bash
    cd "My Office Add-in"
    ```

### <a name="update-the-code"></a><span data-ttu-id="da840-157">コードを更新する</span><span class="sxs-lookup"><span data-stu-id="da840-157">Update the code</span></span>

1. <span data-ttu-id="da840-158">コード エディターで、プロジェクトのルートにある **index.html** を開きます。</span><span class="sxs-lookup"><span data-stu-id="da840-158">In your code editor, open **index.html** in the root of the project.</span></span> <span data-ttu-id="da840-159">このファイルには、アドインの作業ウィンドウにレンダリングされる HTML が含まれています。</span><span class="sxs-lookup"><span data-stu-id="da840-159">This file contains the HTML that will be rendered in the add-in's task pane.</span></span> 

2. <span data-ttu-id="da840-160">`<body>` 要素を次のマークアップに置き換え、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="da840-160">Replace the `<body>` element with the following markup and save the file.</span></span>

    ```html
    <body>
        <div id="content-header">
            <div class="padding">
                <h1>Welcome</h1>
            </div>
        </div>
        <div id="content-main">
            <div class="padding">
                <p>Choose the buttons below to add boilerplate text to the document by using the Word JavaScript API.</p>
                <br />
                <h3>Try it out</h3>
                <button id="emerson">Add quote from Ralph Waldo Emerson</button>
                <br /><br />
                <button id="checkhov">Add quote from Anton Chekhov</button>
                <br /><br />
                <button id="proverb">Add Chinese proverb</button>
            </div>
        </div>
        <br />
        <div id="supportedVersion" />
        <script type="text/javascript" src="node_modules/jquery/dist/jquery.js"></script>
        <script type="text/javascript" src="node_modules/office-ui-fabric-js/dist/js/fabric.js"></script>
    </body>
    ```

2. <span data-ttu-id="da840-161">**src/index.js** ファイルを開いて、アドインのスクリプトを指定します。</span><span class="sxs-lookup"><span data-stu-id="da840-161">Open the file **src/index.js** to specify the script for the add-in.</span></span> <span data-ttu-id="da840-162">すべての内容を次のコードに置き換え、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="da840-162">Replace the entire contents with the following code and save the file.</span></span> <span data-ttu-id="da840-163">このスクリプトには、初期化のコードと、Word 文書に変更を加える (ボタンが選択されたときに、ドキュメントにテキストを挿入する) コードが含まれています。</span><span class="sxs-lookup"><span data-stu-id="da840-163">This script contains initialization code as well as the code that makes changes to the Word document, by inserting text into the document when a button is chosen.</span></span> 

    ```js
    'use strict';
    
    (function () {

        Office.onReady(function() {
            // Office is ready
            $(document).ready(function () {
                // The document is ready
                // Use this to check whether the API is supported in the Word client.
                if (Office.context.requirements.isSetSupported('WordApi', 1.1)) {
                    // Do something that is only available via the new APIs
                    $('#emerson').click(insertEmersonQuoteAtSelection);
                    $('#checkhov').click(insertChekhovQuoteAtTheBeginning);
                    $('#proverb').click(insertChineseProverbAtTheEnd);
                    $('#supportedVersion').html('This code is using Word 2016 or later.');
                }
                else {
                    // Just letting you know that this code will not work with your version of Word.
                    $('#supportedVersion').html('This code requires Word 2016 or later.');
                }
            });
        });

        function insertEmersonQuoteAtSelection() {
            Word.run(function (context) {

                // Create a proxy object for the document.
                var thisDocument = context.document;

                // Queue a command to get the current selection.
                // Create a proxy range object for the selection.
                var range = thisDocument.getSelection();

                // Queue a command to replace the selected text.
                range.insertText('"Hitch your wagon to a star."\n', Word.InsertLocation.replace);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from Ralph Waldo Emerson.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }

        function insertChekhovQuoteAtTheBeginning() {
            Word.run(function (context) {

                // Create a proxy object for the document body.
                var body = context.document.body;

                // Queue a command to insert text at the start of the document body.
                body.insertText('"Knowledge is of no value unless you put it into practice."\n', Word.InsertLocation.start);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from Anton Chekhov.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }

        function insertChineseProverbAtTheEnd() {
            Word.run(function (context) {

                // Create a proxy object for the document body.
                var body = context.document.body;

                // Queue a command to insert text at the end of the document body.
                body.insertText('"To know the road ahead, ask those coming back."\n', Word.InsertLocation.end);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from a Chinese proverb.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
    ```

3. <span data-ttu-id="da840-164">プロジェクトのルートにある **app.css** ファイルを開いて、アドインのカスタム スタイルを指定します。</span><span class="sxs-lookup"><span data-stu-id="da840-164">Open the file **app.css** in the root of the project to specify the custom styles for the add-in.</span></span> <span data-ttu-id="da840-165">すべての内容を次の内容に置き換えて、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="da840-165">Replace the entire contents with the following and save the file.</span></span>

    ```css
    #content-header {
        background: #2a8dd4;
        color: #fff;
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 80px; 
        overflow: hidden;
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px;
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto; 
    }

    .padding {
        padding: 15px;
    }
    ```

### <a name="update-the-manifest"></a><span data-ttu-id="da840-166">マニフェストを更新する</span><span class="sxs-lookup"><span data-stu-id="da840-166">Update the manifest</span></span>

1. <span data-ttu-id="da840-167">**manifest.xml** ファイルを開いて、アドインの設定と機能を定義します。</span><span class="sxs-lookup"><span data-stu-id="da840-167">Open the file **manifest.xml** to define the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="da840-168">`ProviderName` 要素にはプレースホルダー値が含まれています。</span><span class="sxs-lookup"><span data-stu-id="da840-168">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="da840-169">それを自分の名前に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="da840-169">Replace it with your name.</span></span>

3. <span data-ttu-id="da840-170">`Description` 要素の `DefaultValue` 属性にはプレースホルダー値が含まれています。</span><span class="sxs-lookup"><span data-stu-id="da840-170">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="da840-171">これは、**A task pane add-in for Word** に置き換えてください。</span><span class="sxs-lookup"><span data-stu-id="da840-171">Replace it with **A task pane add-in for Word**.</span></span>

4. <span data-ttu-id="da840-172">ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="da840-172">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Word"/>
    ...
    ```

### <a name="start-the-dev-server"></a><span data-ttu-id="da840-173">開発用サーバーの起動</span><span class="sxs-lookup"><span data-stu-id="da840-173">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

### <a name="try-it-out"></a><span data-ttu-id="da840-174">試してみる</span><span class="sxs-lookup"><span data-stu-id="da840-174">Try it out</span></span>

1. <span data-ttu-id="da840-175">Word 内でアドインをサイドロードするには、アドインの実行に使用するプラットフォームの指示に従います。</span><span class="sxs-lookup"><span data-stu-id="da840-175">To sideload the add-in within Word, follow the instructions for the platform you'll use to run your add-in.</span></span>

    - <span data-ttu-id="da840-176">Windows: [Windows で Office アドインをサイドロードする](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="da840-176">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="da840-177">Word Online: [Office Online で Office アドインをサイドロードする](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span><span class="sxs-lookup"><span data-stu-id="da840-177">Word Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span></span>
    - <span data-ttu-id="da840-178">iPad および Mac: [iPad と Mac で Office アドインをサイドロードする](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="da840-178">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

2. <span data-ttu-id="da840-179">Word で、**[ホーム]** タブを選択し、リボンの **[作業ウィンドウの表示]** ボタンをクリックして、アドインの作業ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="da840-179">In Word, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![[作業ウィンドウの表示] ボタンが強調表示されている Word アプリケーションのスクリーンショット](../images/word-quickstart-addin-2.png)

3. <span data-ttu-id="da840-181">作業ウィンドウで、いずれかのボタンを選択して文書に定型句を追加します。</span><span class="sxs-lookup"><span data-stu-id="da840-181">In the task pane, choose any of the buttons to add boilerplate text to the document.</span></span>

    ![定型句アドインが読み込まれている Word アプリケーションのスクリーンショット。](../images/word-quickstart-addin-1.png)

---

## <a name="next-steps"></a><span data-ttu-id="da840-183">次の手順</span><span class="sxs-lookup"><span data-stu-id="da840-183">Next steps</span></span>

<span data-ttu-id="da840-184">これで完了です。jQuery を使用して Word アドインが正常に作成されました。</span><span class="sxs-lookup"><span data-stu-id="da840-184">Congratulations, you've successfully created a Word add-in using jQuery!</span></span> <span data-ttu-id="da840-185">次に、Word アドインの機能の詳細について説明します。Word アドインのチュートリアルに従って、より複雑なアドインをビルドします。</span><span class="sxs-lookup"><span data-stu-id="da840-185">Next, learn more about the capabilities of a Word add-in and build a more complex add-in by following along with the Word add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="da840-186">Word アドインのチュートリアル</span><span class="sxs-lookup"><span data-stu-id="da840-186">Word add-in tutorial</span></span>](../tutorials/word-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="da840-187">関連項目</span><span class="sxs-lookup"><span data-stu-id="da840-187">See also</span></span>

* [<span data-ttu-id="da840-188">Word アドインの概要</span><span class="sxs-lookup"><span data-stu-id="da840-188">Word add-ins overview</span></span>](../word/word-add-ins-programming-overview.md)
* <span data-ttu-id="da840-189">
  [Word アドインのコード サンプル](https://developer.microsoft.com/en-us/office/gallery/?filterBy=Samples,Word)</span><span class="sxs-lookup"><span data-stu-id="da840-189">[Word add-in code samples](https://developer.microsoft.com/en-us/office/gallery/?filterBy=Samples,Word)</span></span>
* [<span data-ttu-id="da840-190">Word JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="da840-190">Word JavaScript API reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/word-add-ins-reference-overview)

