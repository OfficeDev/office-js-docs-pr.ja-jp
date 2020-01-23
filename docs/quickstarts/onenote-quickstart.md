---
title: 最初の OneNote の作業ウィンドウ アドインを作成する
description: ''
ms.date: 01/16/2020
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: 2c7cd2a7adbfc8b5c368c4933a3910d4b3d32b82
ms.sourcegitcommit: 8bce9c94540ed484d0749f07123dc7c72a6ca126
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/22/2020
ms.locfileid: "41265686"
---
# <a name="build-your-first-onenote-task-pane-add-in"></a><span data-ttu-id="f4fcc-102">最初の OneNote の作業ウィンドウ アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="f4fcc-102">Build your first OneNote task pane add-in</span></span>

<span data-ttu-id="f4fcc-103">この記事では、OneNote の作業ウィンドウ アドインを作成するプロセスを紹介します。</span><span class="sxs-lookup"><span data-stu-id="f4fcc-103">In this article, you'll walk through the process of building a OneNote task pane add-in.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="f4fcc-104">必須条件</span><span class="sxs-lookup"><span data-stu-id="f4fcc-104">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a><span data-ttu-id="f4fcc-105">アドイン プロジェクトの作成</span><span class="sxs-lookup"><span data-stu-id="f4fcc-105">Create the add-in project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="f4fcc-106">**Choose a project type: (プロジェクトの種類を選択)** `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="f4fcc-106">**Choose a project type:** `Office Add-in Task Pane project`</span></span>
- <span data-ttu-id="f4fcc-107">**Choose a script type: (スクリプトの種類を選択)** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="f4fcc-107">**Choose a script type:** `Javascript`</span></span>
- <span data-ttu-id="f4fcc-108">**What would you want to name your add-in?: (アドインの名前を何にしますか)**</span><span class="sxs-lookup"><span data-stu-id="f4fcc-108">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="f4fcc-109">**Which Office client application would you like to support?: (どの Office クライアント アプリケーションをサポートしますか)**</span><span class="sxs-lookup"><span data-stu-id="f4fcc-109">**Which Office client application would you like to support?**</span></span> `OneNote`

![Yeoman ジェネレーターのプロンプトと応答のスクリーンショット](../images/yo-office-onenote.png)

<span data-ttu-id="f4fcc-111">ウィザードを完了すると、ジェネレーターによってプロジェクトが作成されて、サポートしているノード コンポーネントがインストールされます。</span><span class="sxs-lookup"><span data-stu-id="f4fcc-111">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a><span data-ttu-id="f4fcc-112">プロジェクトを確認する</span><span class="sxs-lookup"><span data-stu-id="f4fcc-112">Explore the project</span></span>

<span data-ttu-id="f4fcc-113">Yeomanジェネレーターで作成したアドインプロジェクトには、原型となる作業ペインアドインのサンプルコードが含まれています。</span><span class="sxs-lookup"><span data-stu-id="f4fcc-113">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span> 

- <span data-ttu-id="f4fcc-114">プロジェクトのルートディレクトリにある **./ manifest.xml**ファイルは、アドインの設定と機能性を定義します。</span><span class="sxs-lookup"><span data-stu-id="f4fcc-114">The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="f4fcc-115">**./src/taskpane/taskpane.html**ファイルには、作業ペイン用のHTMLマークアップが含まれています。</span><span class="sxs-lookup"><span data-stu-id="f4fcc-115">The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="f4fcc-116">**./src/taskpane/taskpane.css**ファイルには、作業ウィンドウ内のコンテンツに適用される CSS が含まれています。</span><span class="sxs-lookup"><span data-stu-id="f4fcc-116">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="f4fcc-117">**./src/taskpane/taskpane.js**ファイルには、作業ウィンドウと Office のホスト アプリケーションの間のやり取りを容易にする Office JavaScript API コードが含まれています。</span><span class="sxs-lookup"><span data-stu-id="f4fcc-117">The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office host application.</span></span>

## <a name="update-the-code"></a><span data-ttu-id="f4fcc-118">コードを更新する</span><span class="sxs-lookup"><span data-stu-id="f4fcc-118">Update the code</span></span>

<span data-ttu-id="f4fcc-119">コード エディターでファイル **./src/taskpane/taskpane.js** を開き、次のコードを **実行** 関数内に追加します。</span><span class="sxs-lookup"><span data-stu-id="f4fcc-119">In your code editor, open the file **./src/taskpane/taskpane.js** and add the following code within the **run** function.</span></span> <span data-ttu-id="f4fcc-120">このコードは、OneNote JavaScript API を使用してページ タイトルを設定し、ページの本文にアウトラインを追加します。</span><span class="sxs-lookup"><span data-stu-id="f4fcc-120">This code uses the OneNote JavaScript API to set the page title and add an outline to the body of the page.</span></span>

```js
try {
    await OneNote.run(async context => {

        // Get the current page.
        var page = context.application.getActivePage();

        // Queue a command to set the page title.
        page.title = "Hello World";

        // Queue a command to add an outline to the page.
        var html = "<p><ol><li>Item #1</li><li>Item #2</li></ol></p>";
        page.addOutline(40, 90, html);

        // Run the queued commands, and return a promise to indicate task completion.
        return context.sync();
    });
} catch (error) {
    console.log("Error: " + error);
}
```

## <a name="try-it-out"></a><span data-ttu-id="f4fcc-121">試してみる</span><span class="sxs-lookup"><span data-stu-id="f4fcc-121">Try it out</span></span>

1. <span data-ttu-id="f4fcc-122">プロジェクトのルート フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="f4fcc-122">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. <span data-ttu-id="f4fcc-123">ローカル Web サーバーを起動し、アドインのサイドロードを行います。</span><span class="sxs-lookup"><span data-stu-id="f4fcc-123">Start the local web server and sideload your add-in.</span></span>

    > [!NOTE]
    > <span data-ttu-id="f4fcc-124">Office アドインは、開発中であっても HTTP ではなく HTTPS を使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f4fcc-124">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="f4fcc-125">次のいずれかのコマンドを実行した後に証明書をインストールするように求められた場合は、Yeoman ジェネレーターによって提供される証明書をインストールするプロンプトを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="f4fcc-125">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    > [!TIP]
    > <span data-ttu-id="f4fcc-126">Mac でアドインをテストしている場合は、先に進む前に次のコマンドを実行してください。</span><span class="sxs-lookup"><span data-stu-id="f4fcc-126">If you're testing your add-in on Mac, run the following command before proceeding.</span></span> <span data-ttu-id="f4fcc-127">このコマンドを実行すると、ローカル Web サーバーが起動します。</span><span class="sxs-lookup"><span data-stu-id="f4fcc-127">When you run this command, the local web server starts.</span></span>
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    <span data-ttu-id="f4fcc-128">プロジェクトのルート ディレクトリから次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="f4fcc-128">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="f4fcc-129">このコマンドを実行すると、ローカル Web サーバーが起動します (まだ実行されていない場合)。</span><span class="sxs-lookup"><span data-stu-id="f4fcc-129">When you run this command, the local web server will start (if it's not already running).</span></span>

    ```command&nbsp;line
    npm run start:web
    ```

3. <span data-ttu-id="f4fcc-130">[OneNote on the web](https://www.onenote.com/notebooks) でノートブックを開き、新しいページを作成します。</span><span class="sxs-lookup"><span data-stu-id="f4fcc-130">In [OneNote on the web](https://www.onenote.com/notebooks), open a notebook and create a new page.</span></span>

4. <span data-ttu-id="f4fcc-131">**[挿入] > [Office アドイン]** の順に選択し、[Office アドイン] ダイアログを開きます。</span><span class="sxs-lookup"><span data-stu-id="f4fcc-131">Choose **Insert > Office Add-ins** to open the Office Add-ins dialog.</span></span>

    - <span data-ttu-id="f4fcc-132">コンシューマー アカウントでサインインしている場合は、**[マイ アドイン]** タブを選択し、**[マイ アドインのアップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="f4fcc-132">If you're signed in with your consumer account, select the **MY ADD-INS** tab, and then choose **Upload My Add-in**.</span></span>

    - <span data-ttu-id="f4fcc-133">職場または学校アカウントでサインインしている場合は、**[自分の所属組織]** タブを選択し、**[マイ アドインのアップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="f4fcc-133">If you're signed in with your work or school account, select the **MY ORGANIZATION** tab, and then select **Upload My Add-in**.</span></span> 

    <span data-ttu-id="f4fcc-134">次の図は、コンシューマー ノートブックの **[マイ アドイン]** タブを示しています。</span><span class="sxs-lookup"><span data-stu-id="f4fcc-134">The following image shows the **MY ADD-INS** tab for consumer notebooks.</span></span>

    <img alt="The Office Add-ins dialog showing the MY ADD-INS tab" src="../images/onenote-office-add-ins-dialog.png" width="500">

5. <span data-ttu-id="f4fcc-135">[アドインのアップロード] ダイアログで、プロジェクト フォルダー内の **manifest.xml** を参照し、**[アップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="f4fcc-135">In the Upload Add-in dialog, browse to **manifest.xml** in your project folder, and then choose **Upload**.</span></span> 

6. <span data-ttu-id="f4fcc-136">**[ホーム]** タブから、リボンの **[作業ウィンドウの表示]** ボタンをクリックします。</span><span class="sxs-lookup"><span data-stu-id="f4fcc-136">From the **Home** tab, choose the **Show Taskpane** button in the ribbon.</span></span> <span data-ttu-id="f4fcc-137">アドインの作業ウィンドウは、OneNote ページの横にある iFrame で開きます。</span><span class="sxs-lookup"><span data-stu-id="f4fcc-137">The add-in task pane opens in an iFrame next to the OneNote page.</span></span>

7. <span data-ttu-id="f4fcc-138">作業ウィンドウの下部にある [**実行**] リンクをクリックしてページ タイトルを設定し、ページの本文にアウトラインを追加します。</span><span class="sxs-lookup"><span data-stu-id="f4fcc-138">At the bottom of the task pane, choose the **Run** link to set the page title and add an outline to the body of the page.</span></span>

    ![このチュートリアルでビルドした OneNote アドイン](../images/onenote-first-add-in-4.png)

## <a name="next-steps"></a><span data-ttu-id="f4fcc-140">次の手順</span><span class="sxs-lookup"><span data-stu-id="f4fcc-140">Next steps</span></span>

<span data-ttu-id="f4fcc-141">おめでとうございます。OneNote の作業ウィンドウ アドインが正常に作成されました。</span><span class="sxs-lookup"><span data-stu-id="f4fcc-141">Congratulations, you've successfully created a OneNote task pane add-in!</span></span> <span data-ttu-id="f4fcc-142">次に、OneNote アドイン構築の中心概念の詳細について説明します。</span><span class="sxs-lookup"><span data-stu-id="f4fcc-142">Next, learn more about the core concepts of building OneNote add-ins.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="f4fcc-143">OneNote の JavaScript API のプログラミングの概要</span><span class="sxs-lookup"><span data-stu-id="f4fcc-143">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)

## <a name="see-also"></a><span data-ttu-id="f4fcc-144">関連項目</span><span class="sxs-lookup"><span data-stu-id="f4fcc-144">See also</span></span>

* [<span data-ttu-id="f4fcc-145">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="f4fcc-145">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="f4fcc-146">Office アドインを構築する</span><span class="sxs-lookup"><span data-stu-id="f4fcc-146">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
* [<span data-ttu-id="f4fcc-147">Office アドインを開発する</span><span class="sxs-lookup"><span data-stu-id="f4fcc-147">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="f4fcc-148">OneNote の JavaScript API のプログラミングの概要</span><span class="sxs-lookup"><span data-stu-id="f4fcc-148">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="f4fcc-149">OneNote JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="f4fcc-149">OneNote JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference)
- [<span data-ttu-id="f4fcc-150">Rubric Grader のサンプル</span><span class="sxs-lookup"><span data-stu-id="f4fcc-150">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)

