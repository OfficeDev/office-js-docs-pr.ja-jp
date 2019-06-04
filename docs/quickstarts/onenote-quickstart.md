---
title: 最初の OneNote の作業ウィンドウ アドインを作成する
description: ''
ms.date: 05/02/2019
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: 48cd9395b269a83630608c52d972508828c5c007
ms.sourcegitcommit: b299b8a5dfffb6102cb14b431bdde4861abfb47f
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/30/2019
ms.locfileid: "34589217"
---
# <a name="build-your-first-onenote-task-pane-add-in"></a><span data-ttu-id="41a64-102">最初の OneNote の作業ウィンドウ アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="41a64-102">Build your first Word task pane add-in</span></span>

<span data-ttu-id="41a64-103">この記事では、OneNote の作業ウィンドウ アドインを作成するプロセスを紹介します。</span><span class="sxs-lookup"><span data-stu-id="41a64-103">In this article, you'll walk through the process of building a Project task pane add-in.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="41a64-104">必須条件</span><span class="sxs-lookup"><span data-stu-id="41a64-104">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a><span data-ttu-id="41a64-105">アドイン プロジェクトの作成</span><span class="sxs-lookup"><span data-stu-id="41a64-105">Create the add-in project</span></span>

1. <span data-ttu-id="41a64-106">Yeoman ジェネレーターを使用して、OneNote アドイン プロジェクトを作成します。</span><span class="sxs-lookup"><span data-stu-id="41a64-106">Use the Yeoman generator to create a OneNote add-in project.</span></span> <span data-ttu-id="41a64-107">次のコマンドを実行し、以下のプロンプトに応答します。</span><span class="sxs-lookup"><span data-stu-id="41a64-107">Run the following command and then answer the prompts as follows:</span></span>

    ```command&nbsp;line
    yo office
    ```

    - <span data-ttu-id="41a64-108">**Choose a project type: (プロジェクトの種類を選択)** `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="41a64-108">**Choose a project type:** `Office Add-in Task Pane project`</span></span>
    - <span data-ttu-id="41a64-109">**Choose a script type: (スクリプトの種類を選択)** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="41a64-109">**Choose a script type:** `Javascript`</span></span>
    - <span data-ttu-id="41a64-110">**What would you want to name your add-in?: (アドインの名前を何にしますか)**</span><span class="sxs-lookup"><span data-stu-id="41a64-110">**What do you want to name your add-in?**</span></span> `My Office Add-in`
    - <span data-ttu-id="41a64-111">**Which Office client application would you like to support?: (どの Office クライアント アプリケーションをサポートしますか)**</span><span class="sxs-lookup"><span data-stu-id="41a64-111">**Which Office client application would you like to support?**</span></span> `OneNote`

    ![Yeoman ジェネレーターのプロンプトと応答のスクリーンショット](../images/yo-office-onenote.png)
    
    <span data-ttu-id="41a64-113">ウィザードを完了すると、ジェネレーターによってプロジェクトが作成されて、サポートしているノード コンポーネントがインストールされます。</span><span class="sxs-lookup"><span data-stu-id="41a64-113">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>
    
2. <span data-ttu-id="41a64-114">プロジェクトのルート フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="41a64-114">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

## <a name="explore-the-project"></a><span data-ttu-id="41a64-115">プロジェクトを確認する</span><span class="sxs-lookup"><span data-stu-id="41a64-115">Explore the project</span></span>

<span data-ttu-id="41a64-116">Yeomanジェネレーターで作成したアドインプロジェクトには、原型となる作業ペインアドインのサンプルコードが含まれています。</span><span class="sxs-lookup"><span data-stu-id="41a64-116">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span> 

- <span data-ttu-id="41a64-117">プロジェクトのルートディレクトリにある **./ manifest.xml**ファイルは、アドインの設定と機能性を定義します。</span><span class="sxs-lookup"><span data-stu-id="41a64-117">The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="41a64-118">**./src/taskpane/taskpane.html**ファイルには、作業ペイン用のHTMLマークアップが含まれています。</span><span class="sxs-lookup"><span data-stu-id="41a64-118">The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="41a64-119">**./src/taskpane/taskpane.css**ファイルには、作業ウィンドウ内のコンテンツに適用される CSS が含まれています。</span><span class="sxs-lookup"><span data-stu-id="41a64-119">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="41a64-120">**./src/taskpane/taskpane.js**ファイルには、作業ウィンドウと Office のホスト アプリケーションの間のやり取りを容易にする Office JavaScript API コードが含まれています。</span><span class="sxs-lookup"><span data-stu-id="41a64-120">The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office host application.</span></span>

## <a name="update-the-code"></a><span data-ttu-id="41a64-121">コードを更新する</span><span class="sxs-lookup"><span data-stu-id="41a64-121">Update the code</span></span>

<span data-ttu-id="41a64-122">コード エディターでファイル **./src/taskpane/taskpane.js** を開き、次のコードを **実行** 関数内に追加します。</span><span class="sxs-lookup"><span data-stu-id="41a64-122">In your code editor, open the file **./src/taskpane/taskpane.js** and add the following code within the **run** function.</span></span> <span data-ttu-id="41a64-123">このコードは、OneNote JavaScript API を使用してページ タイトルを設定し、ページの本文にアウトラインを追加します。</span><span class="sxs-lookup"><span data-stu-id="41a64-123">This code uses the OneNote JavaScript API to set the page title and add an outline to the body of the page.</span></span>

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

## <a name="try-it-out"></a><span data-ttu-id="41a64-124">お試しください</span><span class="sxs-lookup"><span data-stu-id="41a64-124">Try it out</span></span>

> [!NOTE]
> <span data-ttu-id="41a64-125">開発の最中でも、OfficeアドインはHTTPではなくHTTPSを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="41a64-125">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="41a64-126">次のいずれかのコマンドを実行した後に証明書をインストールするように求められた場合は、Yeoman ジェネレーターによって提供される証明書をインストールするプロンプトを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="41a64-126">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

> [!TIP]
> <span data-ttu-id="41a64-127">Mac でアドインをテストしている場合は、先に進む前に次のコマンドを実行してください。</span><span class="sxs-lookup"><span data-stu-id="41a64-127">If you're testing your add-in on Mac, run the following command before proceeding.</span></span> <span data-ttu-id="41a64-128">このコマンドを実行すると、ローカル Web サーバーが起動します。</span><span class="sxs-lookup"><span data-stu-id="41a64-128">When you run this command, the local web server will start.</span></span>
>
> ```command&nbsp;line
> npm run dev-server
> ```

1. <span data-ttu-id="41a64-129">プロジェクトのルート ディレクトリから次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="41a64-129">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="41a64-130">このコマンドを実行すると、ローカル Web サーバーが起動します (まだ実行されていない場合)。</span><span class="sxs-lookup"><span data-stu-id="41a64-130">When you run this command, the local web server will start.</span></span>

    ```command&nbsp;line
    npm run start:web
    ```

2. <span data-ttu-id="41a64-131">[OneNote Online](https://www.onenote.com/notebooks) でノートブックを開いて新しいページを作成します。</span><span class="sxs-lookup"><span data-stu-id="41a64-131">In [OneNote Online](https://www.onenote.com/notebooks), open a notebook and create a new page.</span></span>

3. <span data-ttu-id="41a64-132">**[挿入] > [Office アドイン]** の順に選択し、[Office アドイン] ダイアログを開きます。</span><span class="sxs-lookup"><span data-stu-id="41a64-132">Choose **Insert > Office Add-ins** to open the Office Add-ins dialog.</span></span>

    - <span data-ttu-id="41a64-133">コンシューマー アカウントでサインインしている場合は、**[マイ アドイン]** タブを選択し、**[マイ アドインのアップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="41a64-133">If you're signed in with your consumer account, select the **MY ADD-INS** tab, and then choose **Upload My Add-in**.</span></span>

    - <span data-ttu-id="41a64-134">職場または学校アカウントでサインインしている場合は、**[自分の所属組織]** タブを選択し、**[マイ アドインのアップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="41a64-134">If you're signed in with your work or school account, select the **MY ORGANIZATION** tab, and then select **Upload My Add-in**.</span></span> 

    <span data-ttu-id="41a64-135">次の図は、コンシューマー ノートブックの **[マイ アドイン]** タブを示しています。</span><span class="sxs-lookup"><span data-stu-id="41a64-135">The following image shows the **MY ADD-INS** tab for consumer notebooks.</span></span>

    <img alt="The Office Add-ins dialog showing the MY ADD-INS tab" src="../images/onenote-office-add-ins-dialog.png" width="500">

3. <span data-ttu-id="41a64-136">[アドインのアップロード] ダイアログで、プロジェクト フォルダー内の **manifest.xml** を参照し、**[アップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="41a64-136">In the Upload Add-in dialog, browse to **manifest.xml** in your project folder, and then choose **Upload**.</span></span> 

4. <span data-ttu-id="41a64-137">**[ホーム]** タブから、リボンの **[作業ウィンドウの表示]** ボタンをクリックします。</span><span class="sxs-lookup"><span data-stu-id="41a64-137">From the **Home** tab, choose the **Show Taskpane** button in the ribbon.</span></span> <span data-ttu-id="41a64-138">アドインの作業ウィンドウは、OneNote ページの横にある iFrame で開きます。</span><span class="sxs-lookup"><span data-stu-id="41a64-138">The add-in task pane opens in an iFrame next to the OneNote page.</span></span>

5. <span data-ttu-id="41a64-139">作業ウィンドウの下部にある [**実行**] リンクをクリックしてページ タイトルを設定し、ページの本文にアウトラインを追加します。</span><span class="sxs-lookup"><span data-stu-id="41a64-139">At the bottom of the task pane, choose the **Run** link to set the page title and add an outline to the body of the page.</span></span>

    ![このチュートリアルでビルドした OneNote アドイン](../images/onenote-first-add-in-4.png)

## <a name="next-steps"></a><span data-ttu-id="41a64-141">次の手順</span><span class="sxs-lookup"><span data-stu-id="41a64-141">Next steps</span></span>

<span data-ttu-id="41a64-142">おめでとうございます。OneNote の作業ウィンドウ アドインが正常に作成されました。</span><span class="sxs-lookup"><span data-stu-id="41a64-142">Congratulations, you've successfully created a Word task pane add-in!</span></span> <span data-ttu-id="41a64-143">次に、OneNote アドイン構築の中心概念の詳細について説明します。</span><span class="sxs-lookup"><span data-stu-id="41a64-143">Next, learn more about the core concepts of building OneNote add-ins.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="41a64-144">OneNote の JavaScript API のプログラミングの概要</span><span class="sxs-lookup"><span data-stu-id="41a64-144">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)

## <a name="see-also"></a><span data-ttu-id="41a64-145">関連項目</span><span class="sxs-lookup"><span data-stu-id="41a64-145">See also</span></span>

- [<span data-ttu-id="41a64-146">OneNote の JavaScript API のプログラミングの概要</span><span class="sxs-lookup"><span data-stu-id="41a64-146">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="41a64-147">OneNote JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="41a64-147">OneNote JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/onenote-add-ins-javascript-reference)
- [<span data-ttu-id="41a64-148">Rubric Grader のサンプル</span><span class="sxs-lookup"><span data-stu-id="41a64-148">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [<span data-ttu-id="41a64-149">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="41a64-149">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)

