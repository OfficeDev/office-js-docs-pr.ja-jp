---
title: Angular を使用して Excel 作業ウィンドウ アドインをビルドする
description: ''
ms.date: 01/16/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 6fd28f3b572b7bca27d5bf08b9799333285064bc
ms.sourcegitcommit: 8bce9c94540ed484d0749f07123dc7c72a6ca126
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/22/2020
ms.locfileid: "41265579"
---
# <a name="build-an-excel-task-pane-add-in-using-angular"></a><span data-ttu-id="7e5cc-102">Angular を使用して Excel 作業ウィンドウ アドインをビルドする</span><span class="sxs-lookup"><span data-stu-id="7e5cc-102">Build an Excel task pane add-in using Angular</span></span>

<span data-ttu-id="7e5cc-103">この記事では、Angular と Excel JavaScript API を使用して Excel 作業ウィンドウ アドインを構築するプロセスについて説明します。</span><span class="sxs-lookup"><span data-stu-id="7e5cc-103">In this article, you'll walk through the process of building an Excel task pane add-in using Angular and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="7e5cc-104">前提条件</span><span class="sxs-lookup"><span data-stu-id="7e5cc-104">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a><span data-ttu-id="7e5cc-105">アドイン プロジェクトの作成</span><span class="sxs-lookup"><span data-stu-id="7e5cc-105">Create the add-in project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="7e5cc-106">**Choose a project type: (プロジェクトの種類を選択)** `Office Add-in Task Pane project using Angular framework`</span><span class="sxs-lookup"><span data-stu-id="7e5cc-106">**Choose a project type:** `Office Add-in Task Pane project using Angular framework`</span></span>
- <span data-ttu-id="7e5cc-107">**Choose a script type: (スクリプトの種類を選択)** `TypeScript`</span><span class="sxs-lookup"><span data-stu-id="7e5cc-107">**Choose a script type:** `TypeScript`</span></span>
- <span data-ttu-id="7e5cc-108">**What would you want to name your add-in?: (アドインの名前を何にしますか)**</span><span class="sxs-lookup"><span data-stu-id="7e5cc-108">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="7e5cc-109">**Which Office client application would you like to support?: (どの Office クライアント アプリケーションをサポートしますか)**</span><span class="sxs-lookup"><span data-stu-id="7e5cc-109">**Which Office client application would you like to support?**</span></span> `Excel`

![Yeoman ジェネレーター](../images/yo-office-excel-angular-2.png)

<span data-ttu-id="7e5cc-111">ウィザードを完了すると、ジェネレーターによってプロジェクトが作成されて、サポートしているノード コンポーネントがインストールされます。</span><span class="sxs-lookup"><span data-stu-id="7e5cc-111">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a><span data-ttu-id="7e5cc-112">プロジェクトを確認する</span><span class="sxs-lookup"><span data-stu-id="7e5cc-112">Explore the project</span></span>

<span data-ttu-id="7e5cc-113">Yeoman ジェネレーターで作成したアドイン プロジェクトには、とても基本的な作業ウィンドウ アドインのサンプル コードが含まれています。</span><span class="sxs-lookup"><span data-stu-id="7e5cc-113">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span> <span data-ttu-id="7e5cc-114">アドイン プロジェクトの主要な構成要素を確認したい場合は、コード エディターでプロジェクトを開き、以下に一覧表示されているファイルを確認します。</span><span class="sxs-lookup"><span data-stu-id="7e5cc-114">If you'd like to explore the key components of your add-in project, open the project in your code editor and review the files listed below.</span></span> <span data-ttu-id="7e5cc-115">アドインを試す準備ができたら、次のセクションに進みます。</span><span class="sxs-lookup"><span data-stu-id="7e5cc-115">When you're ready to try out your add-in, proceed to the next section.</span></span>

- <span data-ttu-id="7e5cc-116">プロジェクトのルート ディレクトリにある **manifest.xml** ファイルで、アドインの機能と設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="7e5cc-116">The **manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="7e5cc-117">**./src/taskpane/app/app.component.html** ファイルには、作業ウィンドウ用の HTML マークアップが含まれています。</span><span class="sxs-lookup"><span data-stu-id="7e5cc-117">The **./src/taskpane/app/app.component.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="7e5cc-118">**./src/taskpane/taskpane.css** ファイルには、作業ウィンドウ内のコンテンツに適用される CSS が含まれています。</span><span class="sxs-lookup"><span data-stu-id="7e5cc-118">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="7e5cc-119">**./src/taskpane/app/app.component.ts** ファイルには、作業ウィンドウと Excel の間のやり取りを容易にする Office JavaScript API コードが含まれています。</span><span class="sxs-lookup"><span data-stu-id="7e5cc-119">The **./src/taskpane/app/app.component.ts** file contains the Office JavaScript API code that facilitates interaction between the task pane and Excel.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="7e5cc-120">試してみる</span><span class="sxs-lookup"><span data-stu-id="7e5cc-120">Try it out</span></span>

1. <span data-ttu-id="7e5cc-121">プロジェクトのルート フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="7e5cc-121">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. [!include[Start server section](../includes/quickstart-yo-start-server-excel.md)] 

3. <span data-ttu-id="7e5cc-122">Excel で、**[ホーム]** タブを選択し、リボンの **[作業ウィンドウの表示]** ボタンをクリックして、アドインの作業ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="7e5cc-122">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Excel アドイン ボタン](../images/excel-quickstart-addin-3b.png)

4. <span data-ttu-id="7e5cc-124">ワークシート内で任意のセルの範囲を選択します。</span><span class="sxs-lookup"><span data-stu-id="7e5cc-124">Select any range of cells in the worksheet.</span></span>

5. <span data-ttu-id="7e5cc-125">作業ウィンドウの下部で、**[実行]** リンクを選択して、選択範囲の色を黄色に設定します。</span><span class="sxs-lookup"><span data-stu-id="7e5cc-125">At the bottom of the task pane, choose the **Run** link to set the color of the selected range to yellow.</span></span>

    ![Excel アドイン](../images/excel-quickstart-addin-3c.png)

## <a name="next-steps"></a><span data-ttu-id="7e5cc-127">次の手順</span><span class="sxs-lookup"><span data-stu-id="7e5cc-127">Next steps</span></span>

<span data-ttu-id="7e5cc-128">おめでとうございます! これで Angular を使用して Excel 作業ウィンドウ アドインを作成できました。</span><span class="sxs-lookup"><span data-stu-id="7e5cc-128">Congratulations, you've successfully created an Excel task pane add-in using Angular!</span></span> <span data-ttu-id="7e5cc-129">次に、Excel アドインの機能の詳細について説明します。Excel アドインのチュートリアルに従って、より複雑なアドインをビルドします。</span><span class="sxs-lookup"><span data-stu-id="7e5cc-129">Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="7e5cc-130">Excel アドインのチュートリアル</span><span class="sxs-lookup"><span data-stu-id="7e5cc-130">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="7e5cc-131">関連項目</span><span class="sxs-lookup"><span data-stu-id="7e5cc-131">See also</span></span>

* [<span data-ttu-id="7e5cc-132">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="7e5cc-132">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="7e5cc-133">Office アドインを構築する</span><span class="sxs-lookup"><span data-stu-id="7e5cc-133">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
* [<span data-ttu-id="7e5cc-134">Office アドインを開発する</span><span class="sxs-lookup"><span data-stu-id="7e5cc-134">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
* [<span data-ttu-id="7e5cc-135">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="7e5cc-135">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="7e5cc-136">Excel アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="7e5cc-136">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="7e5cc-137">Excel JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="7e5cc-137">Excel JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)