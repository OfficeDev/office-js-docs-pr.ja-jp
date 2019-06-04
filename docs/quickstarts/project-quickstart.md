---
title: 最初の Project の作業ウィンドウ アドインを作成する
description: ''
ms.date: 05/08/2019
ms.prod: project
localization_priority: Priority
ms.openlocfilehash: 7a7c907eeeb85b2a686c49ebba0558f4ec20568d
ms.sourcegitcommit: b299b8a5dfffb6102cb14b431bdde4861abfb47f
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/30/2019
ms.locfileid: "34589210"
---
# <a name="build-your-first-project-task-pane-add-in"></a><span data-ttu-id="7ba8b-102">最初の Project の作業ウィンドウ アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="7ba8b-102">Build your first Project task pane add-in</span></span>

<span data-ttu-id="7ba8b-103">この記事では、Project の作業ウィンドウ アドインを作成するプロセスを紹介します。</span><span class="sxs-lookup"><span data-stu-id="7ba8b-103">In this article, you'll walk through the process of building a Project task pane add-in.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="7ba8b-104">前提条件</span><span class="sxs-lookup"><span data-stu-id="7ba8b-104">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- <span data-ttu-id="7ba8b-105">Windows の Project 2016 またはそれ以降</span><span class="sxs-lookup"><span data-stu-id="7ba8b-105">Project 2016 or later on Windows</span></span>

## <a name="create-the-add-in"></a><span data-ttu-id="7ba8b-106">アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="7ba8b-106">Create the add-in</span></span>

1. <span data-ttu-id="7ba8b-107">Yeoman ジェネレーターを使用して、Project アドイン プロジェクトを作成します。</span><span class="sxs-lookup"><span data-stu-id="7ba8b-107">Use the Yeoman generator to create a Project add-in project.</span></span> <span data-ttu-id="7ba8b-108">次のコマンドを実行し、以下のプロンプトに応答します。</span><span class="sxs-lookup"><span data-stu-id="7ba8b-108">Run the following command and then answer the prompts as follows:</span></span>

    ```command&nbsp;line
    yo office
    ```

    - <span data-ttu-id="7ba8b-109">**Choose a project type: (プロジェクトの種類を選択)** `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="7ba8b-109">**Choose a project type:** `Office Add-in Task Pane project`</span></span>
    - <span data-ttu-id="7ba8b-110">**Choose a script type: (スクリプトの種類を選択)** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="7ba8b-110">**Choose a script type:** `Javascript`</span></span>
    - <span data-ttu-id="7ba8b-111">**What would you want to name your add-in?: (アドインの名前を何にしますか)**</span><span class="sxs-lookup"><span data-stu-id="7ba8b-111">**What do you want to name your add-in?**</span></span> `My Office Add-in`
    - <span data-ttu-id="7ba8b-112">**Which Office client application would you like to support?: (どの Office クライアント アプリケーションをサポートしますか)**</span><span class="sxs-lookup"><span data-stu-id="7ba8b-112">**Which Office client application would you like to support?**</span></span> `Project`

    ![Yeoman ジェネレーターのプロンプトと応答のスクリーンショット](../images/yo-office-project.png)
    
    <span data-ttu-id="7ba8b-114">ウィザードを完了すると、ジェネレーターによってプロジェクトが作成されて、サポートしているノード コンポーネントがインストールされます。</span><span class="sxs-lookup"><span data-stu-id="7ba8b-114">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>
    
2. <span data-ttu-id="7ba8b-115">プロジェクトのルート フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="7ba8b-115">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

## <a name="explore-the-project"></a><span data-ttu-id="7ba8b-116">プロジェクトを確認する</span><span class="sxs-lookup"><span data-stu-id="7ba8b-116">Explore the project</span></span>

<span data-ttu-id="7ba8b-117">Yeomanジェネレーターで作成したアドインプロジェクトには、原型となる作業ペインアドインのサンプルコードが含まれています。</span><span class="sxs-lookup"><span data-stu-id="7ba8b-117">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span> 

- <span data-ttu-id="7ba8b-118">プロジェクトのルートディレクトリにある **./ manifest.xml**ファイルは、アドインの設定と機能性を定義します。</span><span class="sxs-lookup"><span data-stu-id="7ba8b-118">The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="7ba8b-119">**./src/taskpane/taskpane.html**ファイルには、作業ペイン用のHTMLマークアップが含まれています。</span><span class="sxs-lookup"><span data-stu-id="7ba8b-119">The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="7ba8b-120">**./src/taskpane/taskpane.css**ファイルには、作業ウィンドウ内のコンテンツに適用される CSS が含まれています。</span><span class="sxs-lookup"><span data-stu-id="7ba8b-120">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="7ba8b-121">**./src/taskpane/taskpane.js**ファイルには、作業ウィンドウと Office のホスト アプリケーションの間のやり取りを容易にする Office JavaScript API コードが含まれています。</span><span class="sxs-lookup"><span data-stu-id="7ba8b-121">The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office host application.</span></span>

## <a name="update-the-code"></a><span data-ttu-id="7ba8b-122">コードを更新する</span><span class="sxs-lookup"><span data-stu-id="7ba8b-122">Update the code</span></span>

<span data-ttu-id="7ba8b-123">コード エディターでファイル **./src/taskpane/taskpane.js** を開き、次のコードを **実行** 関数内に追加します。</span><span class="sxs-lookup"><span data-stu-id="7ba8b-123">In your code editor, open the file **./src/taskpane/taskpane.js** and add the following code within the **run** function.</span></span> <span data-ttu-id="7ba8b-124">このコードでは、Office JavaScript API を使用して、選択したタスクの `Name`フィールドと `Notes` フィールドを設定します。</span><span class="sxs-lookup"><span data-stu-id="7ba8b-124">This code uses the Office JavaScript API to set the `Name` field and `Notes` field of the selected task.</span></span>

```js
var taskGuid;

// Get the GUID of the selected task
Office.context.document.getSelectedTaskAsync(
    function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            taskGuid = result.value;

            // Set the specified fields for the selected task.
            var targetFields = [Office.ProjectTaskFields.Name, Office.ProjectTaskFields.Notes];
            var fieldValues = ['New task name', 'Notes for the task.'];

            // Set the field value. If the call is successful, set the next field.
            for (var i = 0; i < targetFields.length; i++) {
                Office.context.document.setTaskFieldAsync(
                    taskGuid,
                    targetFields[i],
                    fieldValues[i],
                    function (result) {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            i++;
                        }
                        else {
                            var err = result.error;
                            console.log(err.name + ' ' + err.code + ' ' + err.message);
                        }
                    }
                );
            }
        } else {
            var err = result.error;
            console.log(err.name + ' ' + err.code + ' ' + err.message);
        }
    }
);
```

## <a name="try-it-out"></a><span data-ttu-id="7ba8b-125">試してみる</span><span class="sxs-lookup"><span data-stu-id="7ba8b-125">Try it out</span></span>

> [!NOTE]
> <span data-ttu-id="7ba8b-126">開発の最中でも、Office アドインは HTTP ではなく HTTPS を使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="7ba8b-126">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="7ba8b-127">次のコマンドを実行した後に証明書をインストールするように求められた場合は、Yeoman ジェネレーターによって提供される証明書をインストールするプロンプトを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="7ba8b-127">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

1. <span data-ttu-id="7ba8b-128">プロジェクトのルート ディレクトリから次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="7ba8b-128">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="7ba8b-129">このコマンドを実行すると、ローカル Web サーバーが起動します。</span><span class="sxs-lookup"><span data-stu-id="7ba8b-129">When you run this command, the local web server will start.</span></span>

    ```command&nbsp;line
    npm start
    ```

2. <span data-ttu-id="7ba8b-130">Project で、簡素なプロジェクト計画を作成します。</span><span class="sxs-lookup"><span data-stu-id="7ba8b-130">In Project, create a simple project plan.</span></span>

3. <span data-ttu-id="7ba8b-131">[Windows に Office アドインをサイドロードする](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) の手順に従い、Project でアドインを読み込みます。</span><span class="sxs-lookup"><span data-stu-id="7ba8b-131">Load your add-in in Project by following the instructions in [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span></span>

4. <span data-ttu-id="7ba8b-132">プロジェクト内の単一のタスクを選択します。</span><span class="sxs-lookup"><span data-stu-id="7ba8b-132">Select a single task within the project.</span></span>

5. <span data-ttu-id="7ba8b-133">作業ウィンドウの下部で **Run** リンクを選択して、 選択されたタスクの名前を変更し、そのタスクにメモを追加します。</span><span class="sxs-lookup"><span data-stu-id="7ba8b-133">At the bottom of the task pane, choose the **Run** link to rename the selected task and add notes to the selected task.</span></span>

    ![読み込まれた作業ウィンドウ アドインを用いた Project アプリケーションのスクリーンショット](../images/project-quickstart-addin-1.png)

## <a name="next-steps"></a><span data-ttu-id="7ba8b-135">次の手順</span><span class="sxs-lookup"><span data-stu-id="7ba8b-135">Next steps</span></span>

<span data-ttu-id="7ba8b-136">おめでとうございます。 Project の作業ウィンドウ アドインが正常に作成されました。</span><span class="sxs-lookup"><span data-stu-id="7ba8b-136">Congratulations, you've successfully created a Project task pane add-in!</span></span> <span data-ttu-id="7ba8b-137">この後は、Project アドインの機能と一般的なシナリオについて調べます。</span><span class="sxs-lookup"><span data-stu-id="7ba8b-137">Next, learn more about the capabilities of a Project add-in and explore common scenarios.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="7ba8b-138">Project 用アドイン</span><span class="sxs-lookup"><span data-stu-id="7ba8b-138">Project add-ins</span></span>](../project/project-add-ins.md)

