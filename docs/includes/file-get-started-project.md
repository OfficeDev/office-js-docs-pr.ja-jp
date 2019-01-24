---
ms.openlocfilehash: 838db9c0e4a65a8b3ee95deeff5dc04fb0907355
ms.sourcegitcommit: 984c425e2ad58577af8f494079923cab165ad36c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/18/2019
ms.locfileid: "28726983"
---
# <a name="build-your-first-project-add-in"></a><span data-ttu-id="66f7f-101">最初の Project アドインをビルドする</span><span class="sxs-lookup"><span data-stu-id="66f7f-101">Build your first Project add-in</span></span>

<span data-ttu-id="66f7f-102">この記事では、jQuery と Office JavaScript API を使用して Project アドインを作成する手順について説明します。</span><span class="sxs-lookup"><span data-stu-id="66f7f-102">In this article, you'll walk through the process of building a Project add-in by using jQuery and the Office JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="66f7f-103">前提条件</span><span class="sxs-lookup"><span data-stu-id="66f7f-103">Prerequisites</span></span>

- [<span data-ttu-id="66f7f-104">Node.js</span><span class="sxs-lookup"><span data-stu-id="66f7f-104">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="66f7f-105">[Yeoman](https://github.com/yeoman/yo) の最新バージョンと [Office アドイン用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)をグローバルにインストールします。</span><span class="sxs-lookup"><span data-stu-id="66f7f-105">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-add-in"></a><span data-ttu-id="66f7f-106">アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="66f7f-106">Create the add-in</span></span>

1. <span data-ttu-id="66f7f-107">Yeoman ジェネレーターを使用して、Project アドイン プロジェクトを作成します。</span><span class="sxs-lookup"><span data-stu-id="66f7f-107">Use the Yeoman generator to create a Project add-in project.</span></span> <span data-ttu-id="66f7f-108">次のコマンドを実行し、以下のプロンプトに応答します。</span><span class="sxs-lookup"><span data-stu-id="66f7f-108">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="66f7f-109">**Choose a project type: (プロジェクトの種類を選択)** `Office Add-in project using Jquery framework`</span><span class="sxs-lookup"><span data-stu-id="66f7f-109">**Choose a project type:** `Office Add-in project using Jquery framework`</span></span>
    - <span data-ttu-id="66f7f-110">**Choose a script type: (スクリプトの種類を選択)** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="66f7f-110">**Choose a script type:** `Javascript`</span></span>
    - <span data-ttu-id="66f7f-111">**What would you want to name your add-in?: (アドインの名前を何にしますか)** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="66f7f-111">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="66f7f-112">**Which Office client application would you like to support?: (どの Office クライアント アプリケーションをサポートしますか)** `Project`</span><span class="sxs-lookup"><span data-stu-id="66f7f-112">**Which Office client application would you like to support?:** `Project`</span></span>

    ![Yeoman ジェネレーターのプロンプトと応答のスクリーンショット](../images/yo-office-project-jquery.png)
    
    <span data-ttu-id="66f7f-114">ウィザードを完了すると、ジェネレーターによってプロジェクトが作成されて、サポートしているノード コンポーネントがインストールされます。</span><span class="sxs-lookup"><span data-stu-id="66f7f-114">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>
    
2. <span data-ttu-id="66f7f-115">プロジェクトのルート フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="66f7f-115">Navigate to the root folder of the project.</span></span>

    ```bash
    cd "My Office Add-in"
    ```

## <a name="update-the-code"></a><span data-ttu-id="66f7f-116">コードを更新する</span><span class="sxs-lookup"><span data-stu-id="66f7f-116">Update the code</span></span>

1. <span data-ttu-id="66f7f-117">コード エディターで、プロジェクトのルートにある **index.html** を開きます。</span><span class="sxs-lookup"><span data-stu-id="66f7f-117">In your code editor, open **index.html** in the root of the project.</span></span> <span data-ttu-id="66f7f-118">このファイルには、アドインの作業ウィンドウにレンダリングされる HTML が含まれています。</span><span class="sxs-lookup"><span data-stu-id="66f7f-118">This file contains the HTML that will be rendered in the add-in's task pane.</span></span>

2. <span data-ttu-id="66f7f-119">`<body>` 要素を次のマークアップに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="66f7f-119">Replace the `<body>` element with the following markup.</span></span>

    ```html
    <body class="ms-font-m ms-welcome">
        <div id="content-header">
            <div class="padding">
                <h1>Welcome</h1>
            </div>
        </div>
        <div id="content-main">
            <div class="padding">
                <p>Select a task and then choose the buttons below and observe the output in the <b>Results</b> textbox.</p>
                <h3>Try it out</h3>
                <button class="ms-Button" id="get-task-guid">Get Task GUID</button>
                <br/><br/>
                <button class="ms-Button" id="get-task">Get Task data</button>
                <br/>
                <h4>Results:</h4>
                <textarea id="result" rows="6" cols="25"></textarea>
            </div>
        </div>
        <script type="text/javascript" src="node_modules/jquery/dist/jquery.js"></script>
        <script type="text/javascript" src="node_modules/office-ui-fabric-js/dist/js/fabric.js"></script>
    </body>
    ```

3. <span data-ttu-id="66f7f-120">**src/index.js** ファイルを開いて、アドインのスクリプトを指定します。</span><span class="sxs-lookup"><span data-stu-id="66f7f-120">Open the file **src/index.js** to specify the script for the add-in.</span></span> <span data-ttu-id="66f7f-121">すべての内容を次のコードに置き換え、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="66f7f-121">Replace the entire contents with the following code and save the file.</span></span>

    ```js
    'use strict';

    (function () {

        var taskGuid;

        Office.onReady(function() {
            // Office is ready
            $(document).ready(function () {
                // The document is ready
                $('#get-task-guid').click(getTaskGUID);
                $('#get-task').click(getTask);
            });
        });

        function getTaskGUID() {
            Office.context.document.getSelectedTaskAsync(function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    result.value = "Task GUID: " + asyncResult.value;
                    taskGuid = asyncResult.value;
                }
                else {
                    console.log(asyncResult.error.message);
                }
            });
        }

        function getTask() {
            if (taskGuid != undefined) {
                Office.context.document.getTaskAsync(
                    taskGuid,
                    function (asyncResult) {
                        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                            var taskInfo = asyncResult.value;
                            var taskOutput = "Task name: " + taskInfo.taskName +
                                            "\nGUID: " + taskGuid +
                                            "\nWSS Id: " + taskInfo.wssTaskId +
                                            "\nResource names: " + taskInfo.resourceNames;
                            result.value = taskOutput;
                        } else {
                            console.log(asyncResult.error.message);
                        }
                    }
                );
            } else {
                result.value = 'Task GUID not valid:\n' + taskGuid;
            } 
        }
    })();
    ```

4. <span data-ttu-id="66f7f-122">プロジェクトのルートにある **app.css** ファイルを開いて、アドインのカスタム スタイルを指定します。</span><span class="sxs-lookup"><span data-stu-id="66f7f-122">Open the file **app.css** in the root of the project to specify the custom styles for the add-in.</span></span> <span data-ttu-id="66f7f-123">すべての内容を次の内容に置き換えて、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="66f7f-123">Replace the entire contents with the following and save the file.</span></span>

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

## <a name="update-the-manifest"></a><span data-ttu-id="66f7f-124">マニフェストを更新する</span><span class="sxs-lookup"><span data-stu-id="66f7f-124">Update the manifest</span></span>

1. <span data-ttu-id="66f7f-125">**manifest.xml** ファイルを開いて、アドインの設定と機能を定義します。</span><span class="sxs-lookup"><span data-stu-id="66f7f-125">Open the file **manifest.xml** to define the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="66f7f-126">`ProviderName` 要素にはプレースホルダー値が含まれています。</span><span class="sxs-lookup"><span data-stu-id="66f7f-126">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="66f7f-127">それを自分の名前に置き換えます。</span><span class="sxs-lookup"><span data-stu-id="66f7f-127">Replace it with your name.</span></span>

3. <span data-ttu-id="66f7f-128">`Description` 要素の `DefaultValue` 属性にはプレースホルダー値が含まれています。</span><span class="sxs-lookup"><span data-stu-id="66f7f-128">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="66f7f-129">これは、**A task pane add-in for Project** に置き換えてください。</span><span class="sxs-lookup"><span data-stu-id="66f7f-129">Replace it with **A task pane add-in for Project**.</span></span>

4. <span data-ttu-id="66f7f-130">ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="66f7f-130">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Project"/>
    ...
    ```

## <a name="start-the-dev-server"></a><span data-ttu-id="66f7f-131">開発用サーバーの起動</span><span class="sxs-lookup"><span data-stu-id="66f7f-131">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

## <a name="try-it-out"></a><span data-ttu-id="66f7f-132">試してみる</span><span class="sxs-lookup"><span data-stu-id="66f7f-132">Try it out</span></span>

1. <span data-ttu-id="66f7f-133">少なくとも 1 つのタスクを含むシンプルなプロジェクトを Project で作成します。</span><span class="sxs-lookup"><span data-stu-id="66f7f-133">In Project, create a simple project that has at least one task.</span></span>

2. <span data-ttu-id="66f7f-134">アドインを実行して、Project 内のアドインをサイドロードするのに使用するプラットフォームの手順に従います。</span><span class="sxs-lookup"><span data-stu-id="66f7f-134">Follow the instructions for the platform you'll use to run your add-in to sideload the add-in within Project.</span></span>

    - <span data-ttu-id="66f7f-135">Windows: [Windows で Office アドインをサイドロードする](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="66f7f-135">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="66f7f-136">Project Online:[Office Online で Office アドインをサイドロードする](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span><span class="sxs-lookup"><span data-stu-id="66f7f-136">Project Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span></span>
    - <span data-ttu-id="66f7f-137">iPad および Mac: [iPad と Mac で Office アドインをサイドロードする](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="66f7f-137">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

3. <span data-ttu-id="66f7f-138">Project でタスクを選択します。</span><span class="sxs-lookup"><span data-stu-id="66f7f-138">In Project, select a task.</span></span>

    ![1 つのタスクが選択された Project のプロジェクト計画のスクリーンショット](../images/project_quickstart_addin_1.png)

4. <span data-ttu-id="66f7f-140">作業ウィンドウで **[タスク GUID を取得]** ボタンを選択して、タスク GUID を **[結果]** テキストボックスに記入します。</span><span class="sxs-lookup"><span data-stu-id="66f7f-140">In the task pane, choose the **Get Task GUID** button to write the task GUID to the **Results** textbox.</span></span>

    ![1 つのタスクが選択された Project のプロジェクト計画および作業ウィンドウのテキストボックスに記入されたタスク GUID のスクリーンショット](../images/project_quickstart_addin_2.png)

5. <span data-ttu-id="66f7f-142">作業ウィンドウで **[タスク データを取得]** ボタンを選択して、選択したタスクのいくつかのプロパティを **[結果]** テキストボックスに記入します。</span><span class="sxs-lookup"><span data-stu-id="66f7f-142">In the task pane, choose the **Get Task data** button to write several properties of the selected task to the **Results** textbox.</span></span>

    ![1 つのタスクが選択された Project のプロジェクト計画および作業ウィンドウのテキストボックスに記入された複数のタスクのプロパティのスクリーンショット](../images/project_quickstart_addin_3.png)

## <a name="next-steps"></a><span data-ttu-id="66f7f-144">次の手順</span><span class="sxs-lookup"><span data-stu-id="66f7f-144">Next steps</span></span>

<span data-ttu-id="66f7f-145">これで完了です。Project アドインが正常に作成されました。</span><span class="sxs-lookup"><span data-stu-id="66f7f-145">Congratulations, you've successfully created a Project add-in!</span></span> <span data-ttu-id="66f7f-146">この後は、Project アドインの機能と一般的なシナリオについて調べます。</span><span class="sxs-lookup"><span data-stu-id="66f7f-146">Next, learn more about the capabilities of a Project add-in and explore common scenarios.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="66f7f-147">Project 用アドイン</span><span class="sxs-lookup"><span data-stu-id="66f7f-147">Project add-ins</span></span>](../project/project-add-ins.md)
