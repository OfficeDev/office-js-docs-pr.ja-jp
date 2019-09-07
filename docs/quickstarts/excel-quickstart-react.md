---
title: React を使用して Excel 作業ウィンドウ アドインを構築する
description: ''
ms.date: 09/06/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 40ecd0f29ab37df56a8d4456ced0b13f8fdc837b
ms.sourcegitcommit: ce7e7087a4550b9c090dc565fee5eac08a2985a2
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/06/2019
ms.locfileid: "36782276"
---
# <a name="build-an-excel-task-pane-add-in-using-react"></a><span data-ttu-id="a2b6c-102">React を使用して Excel 作業ウィンドウ アドインを構築する</span><span class="sxs-lookup"><span data-stu-id="a2b6c-102">Build an Excel task pane add-in using Angular</span></span>

<span data-ttu-id="a2b6c-103">この記事では、React と Excel JavaScript API を使用して Excel 作業ウィンドウ アドインを構築するプロセスについて説明します。</span><span class="sxs-lookup"><span data-stu-id="a2b6c-103">In this article, you'll walk through the process of building an Excel task pane add-in using Angular and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="a2b6c-104">前提条件</span><span class="sxs-lookup"><span data-stu-id="a2b6c-104">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a><span data-ttu-id="a2b6c-105">アドイン プロジェクトの作成</span><span class="sxs-lookup"><span data-stu-id="a2b6c-105">Create the add-in project</span></span>

<span data-ttu-id="a2b6c-106">Yeoman ジェネレーターを使用して、Excel アドイン プロジェクトを作成します。</span><span class="sxs-lookup"><span data-stu-id="a2b6c-106">Use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="a2b6c-107">次のコマンドを実行し、以下のプロンプトに応答します。</span><span class="sxs-lookup"><span data-stu-id="a2b6c-107">Run the following command and then answer the prompts as follows:</span></span>

```command&nbsp;line
yo office
```

- <span data-ttu-id="a2b6c-108">**Choose a project type: (プロジェクトの種類を選択)** `Office Add-in Task Pane project using React framework`</span><span class="sxs-lookup"><span data-stu-id="a2b6c-108">**Choose a project type:** `Office Add-in Task Pane project using React framework`</span></span>
- <span data-ttu-id="a2b6c-109">**Choose a script type: (スクリプトの種類を選択)** `TypeScript`</span><span class="sxs-lookup"><span data-stu-id="a2b6c-109">**Choose a script type:** `TypeScript`</span></span>
- <span data-ttu-id="a2b6c-110">**What would you want to name your add-in?: (アドインの名前を何にしますか)**</span><span class="sxs-lookup"><span data-stu-id="a2b6c-110">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="a2b6c-111">**Which Office client application would you like to support?: (どの Office クライアント アプリケーションをサポートしますか)**</span><span class="sxs-lookup"><span data-stu-id="a2b6c-111">**Which Office client application would you like to support?**</span></span> `Excel`

![Yeoman ジェネレーター](../images/yo-office-excel-react-2.png)

<span data-ttu-id="a2b6c-113">ウィザードを完了すると、ジェネレーターによってプロジェクトが作成されて、サポートしているノード コンポーネントがインストールされます。</span><span class="sxs-lookup"><span data-stu-id="a2b6c-113">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

## <a name="explore-the-project"></a><span data-ttu-id="a2b6c-114">プロジェクトを確認する</span><span class="sxs-lookup"><span data-stu-id="a2b6c-114">Explore the project</span></span>

<span data-ttu-id="a2b6c-115">Yeoman ジェネレーターで作成したアドイン プロジェクトには、とても基本的な作業ウィンドウ アドインのサンプル コードが含まれています。</span><span class="sxs-lookup"><span data-stu-id="a2b6c-115">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span> <span data-ttu-id="a2b6c-116">アドイン プロジェクトの主要な構成要素を確認したい場合は、コード エディターでプロジェクトを開き、以下に一覧表示されているファイルを確認します。</span><span class="sxs-lookup"><span data-stu-id="a2b6c-116">If you'd like to explore the key components of your add-in project, open the project in your code editor and review the files listed below.</span></span> <span data-ttu-id="a2b6c-117">アドインを試す準備ができたら、次のセクションに進みます。</span><span class="sxs-lookup"><span data-stu-id="a2b6c-117">When you're ready to try out your add-in, proceed to the next section.</span></span>

- <span data-ttu-id="a2b6c-118">プロジェクトのルート ディレクトリにある **manifest.xml** ファイルで、アドインの機能と設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="a2b6c-118">The **manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="a2b6c-119">**./src/taskpane/taskpane.html** ファイルは作業ウィンドウの HTML フレームワークを定義し、**./src/taskpane/components** フォルダー内のファイルは作業ウィンドウ UI のさまざまな部分を定義します。</span><span class="sxs-lookup"><span data-stu-id="a2b6c-119">The **./src/taskpane/taskpane.html** file defines the HTML framework of the task pane, and the files within the **./src/taskpane/components** folder define the various parts of the task pane UI.</span></span>
- <span data-ttu-id="a2b6c-120">**./src/taskpane/taskpane.css**ファイルには、作業ウィンドウ内のコンテンツに適用される CSS が含まれています。</span><span class="sxs-lookup"><span data-stu-id="a2b6c-120">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="a2b6c-121">**./src/taskpane/components/App.tsx** ファイルには、作業ウィンドウと Excel の間のやり取りを容易にする Office JavaScript API コードが含まれています。</span><span class="sxs-lookup"><span data-stu-id="a2b6c-121">The **./src/taskpane/app/app.component.ts** file contains the Office JavaScript API code that facilitates interaction between the task pane and Excel.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="a2b6c-122">試してみる</span><span class="sxs-lookup"><span data-stu-id="a2b6c-122">Try it out</span></span>

1. <span data-ttu-id="a2b6c-123">プロジェクトのルート フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="a2b6c-123">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. [!include[Start server section](../includes/quickstart-yo-start-server-excel.md)] 

3. <span data-ttu-id="a2b6c-124">Excel で、**[ホーム]** タブを選択し、リボンの **[作業ウィンドウの表示]** ボタンをクリックして、アドインの作業ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="a2b6c-124">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Excel アドイン ボタン](../images/excel-quickstart-addin-3b.png)

4. <span data-ttu-id="a2b6c-126">ワークシート内で任意のセルの範囲を選択します。</span><span class="sxs-lookup"><span data-stu-id="a2b6c-126">Select any range of cells in the worksheet.</span></span>

5. <span data-ttu-id="a2b6c-127">作業ウィンドウの下部で、**[実行]** リンクを選択して、選択範囲の色を黄色に設定します。</span><span class="sxs-lookup"><span data-stu-id="a2b6c-127">At the bottom of the task pane, choose the **Run** link to set the color of the selected range to yellow.</span></span>

    ![Excel アドイン](../images/excel-quickstart-addin-3c.png)

## <a name="next-steps"></a><span data-ttu-id="a2b6c-129">次の手順</span><span class="sxs-lookup"><span data-stu-id="a2b6c-129">Next steps</span></span>

<span data-ttu-id="a2b6c-130">おめでとうございます! これで React を使用して Excel 作業ウィンドウ アドインを作成できました。</span><span class="sxs-lookup"><span data-stu-id="a2b6c-130">Congratulations, you've successfully created an Excel task pane add-in using Angular!</span></span> <span data-ttu-id="a2b6c-131">次に、Excel アドインの機能の詳細について説明します。Excel アドインのチュートリアルに従って、より複雑なアドインをビルドします。</span><span class="sxs-lookup"><span data-stu-id="a2b6c-131">Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="a2b6c-132">Excel アドインのチュートリアル</span><span class="sxs-lookup"><span data-stu-id="a2b6c-132">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="a2b6c-133">関連項目</span><span class="sxs-lookup"><span data-stu-id="a2b6c-133">See also</span></span>

* [<span data-ttu-id="a2b6c-134">Excel アドインのチュートリアル</span><span class="sxs-lookup"><span data-stu-id="a2b6c-134">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="a2b6c-135">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="a2b6c-135">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="a2b6c-136">Excel アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="a2b6c-136">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="a2b6c-137">Excel JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="a2b6c-137">Excel JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
