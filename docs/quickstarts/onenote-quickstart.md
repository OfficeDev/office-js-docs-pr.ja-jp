---
title: 最初の OneNote の作業ウィンドウ アドインを作成する
description: Office JS API を使用して単純な OneNote 作業ウィンドウ アドインを作成する方法について説明します。
ms.date: 10/14/2020
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: de0729a483057a61be3793e299995aa05d287441
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132292"
---
# <a name="build-your-first-onenote-task-pane-add-in"></a><span data-ttu-id="4fda3-103">最初の OneNote の作業ウィンドウ アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="4fda3-103">Build your first OneNote task pane add-in</span></span>

<span data-ttu-id="4fda3-104">この記事では、OneNote の作業ウィンドウ アドインを作成するプロセスを紹介します。</span><span class="sxs-lookup"><span data-stu-id="4fda3-104">In this article, you'll walk through the process of building a OneNote task pane add-in.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="4fda3-105">必須条件</span><span class="sxs-lookup"><span data-stu-id="4fda3-105">Prerequisites</span></span>

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a><span data-ttu-id="4fda3-106">アドイン プロジェクトの作成</span><span class="sxs-lookup"><span data-stu-id="4fda3-106">Create the add-in project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="4fda3-107">**Choose a project type: (プロジェクトの種類を選択)** `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="4fda3-107">**Choose a project type:** `Office Add-in Task Pane project`</span></span>
- <span data-ttu-id="4fda3-108">**Choose a script type: (スクリプトの種類を選択)** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="4fda3-108">**Choose a script type:** `Javascript`</span></span>
- <span data-ttu-id="4fda3-109">**What would you want to name your add-in?: (アドインの名前を何にしますか)**</span><span class="sxs-lookup"><span data-stu-id="4fda3-109">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="4fda3-110">**Which Office client application would you like to support?: (どの Office クライアント アプリケーションをサポートしますか)**</span><span class="sxs-lookup"><span data-stu-id="4fda3-110">**Which Office client application would you like to support?**</span></span> `OneNote`

![コマンドライン インターフェイスでの Yeoman ジェネレーターのプロンプトと回答を示すスクリーンショット](../images/yo-office-onenote.png)

<span data-ttu-id="4fda3-112">ウィザードを完了すると、ジェネレーターによってプロジェクトが作成されて、サポートしているノード コンポーネントがインストールされます。</span><span class="sxs-lookup"><span data-stu-id="4fda3-112">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a><span data-ttu-id="4fda3-113">プロジェクトを確認する</span><span class="sxs-lookup"><span data-stu-id="4fda3-113">Explore the project</span></span>

<span data-ttu-id="4fda3-114">Yeomanジェネレーターで作成したアドインプロジェクトには、原型となる作業ペインアドインのサンプルコードが含まれています。</span><span class="sxs-lookup"><span data-stu-id="4fda3-114">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span>

- <span data-ttu-id="4fda3-115">プロジェクトのルートディレクトリにある **./ manifest.xml** ファイルは、アドインの設定と機能性を定義します。</span><span class="sxs-lookup"><span data-stu-id="4fda3-115">The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="4fda3-116">**./src/taskpane/taskpane.html** ファイルには、作業ペイン用のHTMLマークアップが含まれています。</span><span class="sxs-lookup"><span data-stu-id="4fda3-116">The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="4fda3-117">**./src/taskpane/taskpane.css** ファイルには、作業ペインのコンテンツに適用されるCSSが含まれています。</span><span class="sxs-lookup"><span data-stu-id="4fda3-117">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="4fda3-118">**./src/taskpane/taskpane.js** ファイルには、作業ウィンドウと Office クライアント アプリケーションの間のやり取りを容易にする Office JavaScript API コードが含まれています。</span><span class="sxs-lookup"><span data-stu-id="4fda3-118">The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office client application.</span></span>

## <a name="update-the-code"></a><span data-ttu-id="4fda3-119">コードを更新する</span><span class="sxs-lookup"><span data-stu-id="4fda3-119">Update the code</span></span>

<span data-ttu-id="4fda3-120">コード エディターでファイル **./src/taskpane/taskpane.js** を開き、次のコードを `run` 関数内に追加します。</span><span class="sxs-lookup"><span data-stu-id="4fda3-120">In your code editor, open the file **./src/taskpane/taskpane.js** and add the following code within the `run` function.</span></span> <span data-ttu-id="4fda3-121">このコードは、OneNote JavaScript API を使用してページ タイトルを設定し、ページの本文にアウトラインを追加します。</span><span class="sxs-lookup"><span data-stu-id="4fda3-121">This code uses the OneNote JavaScript API to set the page title and add an outline to the body of the page.</span></span>

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

## <a name="try-it-out"></a><span data-ttu-id="4fda3-122">試してみる</span><span class="sxs-lookup"><span data-stu-id="4fda3-122">Try it out</span></span>

1. <span data-ttu-id="4fda3-123">プロジェクトのルート フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="4fda3-123">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. <span data-ttu-id="4fda3-124">ローカル Web サーバーを起動し、アドインのサイドロードを行います。</span><span class="sxs-lookup"><span data-stu-id="4fda3-124">Start the local web server and sideload your add-in.</span></span>

    > [!NOTE]
    > <span data-ttu-id="4fda3-125">Office アドインは、開発中であっても HTTP ではなく HTTPS を使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="4fda3-125">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="4fda3-126">次のいずれかのコマンドを実行した後に証明書をインストールするように求められた場合は、Yeoman ジェネレーターによって提供される証明書をインストールするプロンプトを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="4fda3-126">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span> <span data-ttu-id="4fda3-127">変更を行うには、管理者としてコマンド プロンプトまたはターミナルを実行する必要がある場合もあります。</span><span class="sxs-lookup"><span data-stu-id="4fda3-127">You may also have to run your command prompt or terminal as an administrator for the changes to be made.</span></span>

    > [!TIP]
    > <span data-ttu-id="4fda3-128">Mac でアドインをテストしている場合は、先に進む前に次のコマンドを実行してください。</span><span class="sxs-lookup"><span data-stu-id="4fda3-128">If you're testing your add-in on Mac, run the following command before proceeding.</span></span> <span data-ttu-id="4fda3-129">このコマンドを実行すると、ローカル Web サーバーが起動します。</span><span class="sxs-lookup"><span data-stu-id="4fda3-129">When you run this command, the local web server starts.</span></span>
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    <span data-ttu-id="4fda3-130">プロジェクトのルート ディレクトリから次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="4fda3-130">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="4fda3-131">このコマンドを実行すると、ローカル Web サーバーが起動します (まだ実行されていない場合)。</span><span class="sxs-lookup"><span data-stu-id="4fda3-131">When you run this command, the local web server will start (if it's not already running).</span></span>

    ```command&nbsp;line
    npm run start:web
    ```

3. <span data-ttu-id="4fda3-132">[OneNote on the web](https://www.onenote.com/notebooks) でノートブックを開き、新しいページを作成します。</span><span class="sxs-lookup"><span data-stu-id="4fda3-132">In [OneNote on the web](https://www.onenote.com/notebooks), open a notebook and create a new page.</span></span>

4. <span data-ttu-id="4fda3-133">**[挿入] > [Office アドイン]** の順に選択し、[Office アドイン] ダイアログを開きます。</span><span class="sxs-lookup"><span data-stu-id="4fda3-133">Choose **Insert > Office Add-ins** to open the Office Add-ins dialog.</span></span>

    - <span data-ttu-id="4fda3-134">コンシューマー アカウントでサインインしている場合は、**[マイ アドイン]** タブを選択し、**[マイ アドインのアップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="4fda3-134">If you're signed in with your consumer account, select the **MY ADD-INS** tab, and then choose **Upload My Add-in**.</span></span>

    - <span data-ttu-id="4fda3-135">職場または教育用アカウントでサインインしている場合は、**[自分の所属組織]** タブを選択し、**[マイ アドインのアップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="4fda3-135">If you're signed in with your work or education account, select the **MY ORGANIZATION** tab, and then select **Upload My Add-in**.</span></span>

    <span data-ttu-id="4fda3-136">次の図は、コンシューマー ノートブックの **[マイ アドイン]** タブを示しています。</span><span class="sxs-lookup"><span data-stu-id="4fda3-136">The following image shows the **MY ADD-INS** tab for consumer notebooks.</span></span>

    <img alt="Screenshot of the Office Add-ins dialog showing the MY ADD-INS tab" src="../images/onenote-office-add-ins-dialog.png" width="500">

5. <span data-ttu-id="4fda3-137">[アドインのアップロード] ダイアログで、プロジェクト フォルダー内の **manifest.xml** を参照し、**[アップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="4fda3-137">In the Upload Add-in dialog, browse to **manifest.xml** in your project folder, and then choose **Upload**.</span></span>

6. <span data-ttu-id="4fda3-138">**[ホーム]** タブから、リボンの **[作業ウィンドウの表示]** ボタンをクリックします。</span><span class="sxs-lookup"><span data-stu-id="4fda3-138">From the **Home** tab, choose the **Show Taskpane** button in the ribbon.</span></span> <span data-ttu-id="4fda3-139">アドインの作業ウィンドウは、OneNote ページの横にある iFrame で開きます。</span><span class="sxs-lookup"><span data-stu-id="4fda3-139">The add-in task pane opens in an iFrame next to the OneNote page.</span></span>

7. <span data-ttu-id="4fda3-140">作業ウィンドウの下部にある [**実行**] リンクをクリックしてページ タイトルを設定し、ページの本文にアウトラインを追加します。</span><span class="sxs-lookup"><span data-stu-id="4fda3-140">At the bottom of the task pane, choose the **Run** link to set the page title and add an outline to the body of the page.</span></span>

    ![このウォークスルーから作成されたアドインを示すスクリーンショット: OneNote で [タスクペインのリボンボタンとタスクペインを表示する]](../images/onenote-first-add-in-4.png)

## <a name="next-steps"></a><span data-ttu-id="4fda3-142">次の手順</span><span class="sxs-lookup"><span data-stu-id="4fda3-142">Next steps</span></span>

<span data-ttu-id="4fda3-143">おめでとうございます。OneNote の作業ウィンドウ アドインが正常に作成されました。</span><span class="sxs-lookup"><span data-stu-id="4fda3-143">Congratulations, you've successfully created a OneNote task pane add-in!</span></span> <span data-ttu-id="4fda3-144">次に、OneNote アドイン構築の中心概念の詳細について説明します。</span><span class="sxs-lookup"><span data-stu-id="4fda3-144">Next, learn more about the core concepts of building OneNote add-ins.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="4fda3-145">OneNote の JavaScript API のプログラミングの概要</span><span class="sxs-lookup"><span data-stu-id="4fda3-145">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)

## <a name="see-also"></a><span data-ttu-id="4fda3-146">関連項目</span><span class="sxs-lookup"><span data-stu-id="4fda3-146">See also</span></span>

- [<span data-ttu-id="4fda3-147">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="4fda3-147">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
- [<span data-ttu-id="4fda3-148">Office アドインを開発する</span><span class="sxs-lookup"><span data-stu-id="4fda3-148">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="4fda3-149">OneNote の JavaScript API のプログラミングの概要</span><span class="sxs-lookup"><span data-stu-id="4fda3-149">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="4fda3-150">OneNote JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="4fda3-150">OneNote JavaScript API reference</span></span>](../reference/overview/onenote-add-ins-javascript-reference.md)
- [<span data-ttu-id="4fda3-151">Rubric Grader のサンプル</span><span class="sxs-lookup"><span data-stu-id="4fda3-151">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
