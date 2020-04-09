---
title: 最初の Project の作業ウィンドウ アドインを作成する
description: Office JS API を使用して単純な Project 作業ウィンドウ アドインを作成する方法について説明します。
ms.date: 04/03/2020
ms.prod: project
localization_priority: Priority
ms.openlocfilehash: f712dbfc10027ff6af0eaf618c667cd542bbf284
ms.sourcegitcommit: c3bfea0818af1f01e71a1feff707fb2456a69488
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/08/2020
ms.locfileid: "43185439"
---
# <a name="build-your-first-project-task-pane-add-in"></a><span data-ttu-id="5863a-103">最初の Project の作業ウィンドウ アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="5863a-103">Build your first Project task pane add-in</span></span>

<span data-ttu-id="5863a-104">この記事では、Project の作業ウィンドウ アドインを作成するプロセスを紹介します。</span><span class="sxs-lookup"><span data-stu-id="5863a-104">In this article, you'll walk through the process of building a Project task pane add-in.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="5863a-105">前提条件</span><span class="sxs-lookup"><span data-stu-id="5863a-105">Prerequisites</span></span>

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- <span data-ttu-id="5863a-106">Windows の Project 2016 またはそれ以降</span><span class="sxs-lookup"><span data-stu-id="5863a-106">Project 2016 or later on Windows</span></span>

## <a name="create-the-add-in"></a><span data-ttu-id="5863a-107">アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="5863a-107">Create the add-in</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="5863a-108">**Choose a project type: (プロジェクトの種類を選択)** `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="5863a-108">**Choose a project type:** `Office Add-in Task Pane project`</span></span>
- <span data-ttu-id="5863a-109">**Choose a script type: (スクリプトの種類を選択)** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="5863a-109">**Choose a script type:** `Javascript`</span></span>
- <span data-ttu-id="5863a-110">**What would you want to name your add-in?: (アドインの名前を何にしますか)**</span><span class="sxs-lookup"><span data-stu-id="5863a-110">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="5863a-111">**Which Office client application would you like to support?: (どの Office クライアント アプリケーションをサポートしますか)**</span><span class="sxs-lookup"><span data-stu-id="5863a-111">**Which Office client application would you like to support?**</span></span> `Project`

![Yeoman ジェネレーターのプロンプトと応答のスクリーンショット](../images/yo-office-project.png)

<span data-ttu-id="5863a-113">ウィザードを完了すると、ジェネレーターによってプロジェクトが作成されて、サポートしているノード コンポーネントがインストールされます。</span><span class="sxs-lookup"><span data-stu-id="5863a-113">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a><span data-ttu-id="5863a-114">プロジェクトを確認する</span><span class="sxs-lookup"><span data-stu-id="5863a-114">Explore the project</span></span>

<span data-ttu-id="5863a-115">Yeomanジェネレーターで作成したアドインプロジェクトには、原型となる作業ペインアドインのサンプルコードが含まれています。</span><span class="sxs-lookup"><span data-stu-id="5863a-115">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span> 

- <span data-ttu-id="5863a-116">プロジェクトのルートディレクトリにある **./ manifest.xml**ファイルは、アドインの設定と機能性を定義します。</span><span class="sxs-lookup"><span data-stu-id="5863a-116">The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="5863a-117">**./src/taskpane/taskpane.html**ファイルには、作業ペイン用のHTMLマークアップが含まれています。</span><span class="sxs-lookup"><span data-stu-id="5863a-117">The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="5863a-118">**./src/taskpane/taskpane.css**ファイルには、作業ウィンドウ内のコンテンツに適用される CSS が含まれています。</span><span class="sxs-lookup"><span data-stu-id="5863a-118">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="5863a-119">**./src/taskpane/taskpane.js**ファイルには、作業ウィンドウと Office のホスト アプリケーションの間のやり取りを容易にする Office JavaScript API コードが含まれています。</span><span class="sxs-lookup"><span data-stu-id="5863a-119">The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office host application.</span></span>

## <a name="update-the-code"></a><span data-ttu-id="5863a-120">コードを更新する</span><span class="sxs-lookup"><span data-stu-id="5863a-120">Update the code</span></span>

<span data-ttu-id="5863a-121">コード エディターでファイル **./src/taskpane/taskpane.js** を開き、次のコードを `run` 関数内に追加します。</span><span class="sxs-lookup"><span data-stu-id="5863a-121">In your code editor, open the file **./src/taskpane/taskpane.js** and add the following code within the `run` function.</span></span> <span data-ttu-id="5863a-122">このコードでは、Office JavaScript API を使用して、選択したタスクの `Name`フィールドと `Notes` フィールドを設定します。</span><span class="sxs-lookup"><span data-stu-id="5863a-122">This code uses the Office JavaScript API to set the `Name` field and `Notes` field of the selected task.</span></span>

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

## <a name="try-it-out"></a><span data-ttu-id="5863a-123">試してみる</span><span class="sxs-lookup"><span data-stu-id="5863a-123">Try it out</span></span>

1. <span data-ttu-id="5863a-124">プロジェクトのルート フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="5863a-124">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. <span data-ttu-id="5863a-125">ローカル Web サーバーを開始します。</span><span class="sxs-lookup"><span data-stu-id="5863a-125">Start the local web server.</span></span>

    > [!NOTE]
    > <span data-ttu-id="5863a-126">開発の最中でも、OfficeアドインはHTTPではなくHTTPSを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="5863a-126">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="5863a-127">次のコマンドを実行した後に証明書をインストールするように求められた場合は、Yeoman ジェネレーターによって提供される証明書をインストールするプロンプトを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="5863a-127">If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    <span data-ttu-id="5863a-128">プロジェクトのルート ディレクトリから次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="5863a-128">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="5863a-129">このコマンドを実行すると、ローカル Web サーバーが起動します。</span><span class="sxs-lookup"><span data-stu-id="5863a-129">When you run this command, the local web server will start.</span></span>

    ```command&nbsp;line
    npm start
    ```

3. <span data-ttu-id="5863a-130">Project で、簡素なプロジェクト計画を作成します。</span><span class="sxs-lookup"><span data-stu-id="5863a-130">In Project, create a simple project plan.</span></span>

4. <span data-ttu-id="5863a-131">[Windows に Office アドインをサイドロードする](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) の手順に従い、Project でアドインを読み込みます。</span><span class="sxs-lookup"><span data-stu-id="5863a-131">Load your add-in in Project by following the instructions in [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span></span>

5. <span data-ttu-id="5863a-132">プロジェクト内の単一のタスクを選択します。</span><span class="sxs-lookup"><span data-stu-id="5863a-132">Select a single task within the project.</span></span>

6. <span data-ttu-id="5863a-133">作業ウィンドウの下部で **Run** リンクを選択して、 選択されたタスクの名前を変更し、そのタスクにメモを追加します。</span><span class="sxs-lookup"><span data-stu-id="5863a-133">At the bottom of the task pane, choose the **Run** link to rename the selected task and add notes to the selected task.</span></span>

    ![読み込まれた作業ウィンドウ アドインを用いた Project アプリケーションのスクリーンショット](../images/project-quickstart-addin-1.png)

## <a name="next-steps"></a><span data-ttu-id="5863a-135">次の手順</span><span class="sxs-lookup"><span data-stu-id="5863a-135">Next steps</span></span>

<span data-ttu-id="5863a-136">おめでとうございます。 Project の作業ウィンドウ アドインが正常に作成されました。</span><span class="sxs-lookup"><span data-stu-id="5863a-136">Congratulations, you've successfully created a Project task pane add-in!</span></span> <span data-ttu-id="5863a-137">この後は、Project アドインの機能と一般的なシナリオについて調べます。</span><span class="sxs-lookup"><span data-stu-id="5863a-137">Next, learn more about the capabilities of a Project add-in and explore common scenarios.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="5863a-138">Project 用アドイン</span><span class="sxs-lookup"><span data-stu-id="5863a-138">Project add-ins</span></span>](../project/project-add-ins.md)

## <a name="see-also"></a><span data-ttu-id="5863a-139">関連項目</span><span class="sxs-lookup"><span data-stu-id="5863a-139">See also</span></span>

- [<span data-ttu-id="5863a-140">Office アドインを構築する</span><span class="sxs-lookup"><span data-stu-id="5863a-140">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
- [<span data-ttu-id="5863a-141">Office アドインの中心概念</span><span class="sxs-lookup"><span data-stu-id="5863a-141">Core concepts for Office Add-ins</span></span>](../overview/core-concepts-office-add-ins.md)
- [<span data-ttu-id="5863a-142">Office アドインを開発する</span><span class="sxs-lookup"><span data-stu-id="5863a-142">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
