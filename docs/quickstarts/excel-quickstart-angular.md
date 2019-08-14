---
title: Angular を使用して Excel 作業ウィンドウ アドインをビルドする
description: ''
ms.date: 05/02/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 5d27b7d70fb013ca15115af3ad550a604ce10bb9
ms.sourcegitcommit: 1dc1bb0befe06d19b587961da892434bd0512fb5
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/13/2019
ms.locfileid: "36308037"
---
# <a name="build-an-excel-task-pane-add-in-using-angular"></a><span data-ttu-id="7a1a2-102">Angular を使用して Excel 作業ウィンドウ アドインをビルドする</span><span class="sxs-lookup"><span data-stu-id="7a1a2-102">Build an Excel task pane add-in using Angular</span></span>

<span data-ttu-id="7a1a2-103">この記事では、Angular と Excel JavaScript API を使用して Excel 作業ウィンドウ アドインを構築するプロセスについて説明します。</span><span class="sxs-lookup"><span data-stu-id="7a1a2-103">In this article, you'll walk through the process of building an Excel task pane add-in using Angular and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="7a1a2-104">前提条件</span><span class="sxs-lookup"><span data-stu-id="7a1a2-104">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a><span data-ttu-id="7a1a2-105">アドイン プロジェクトの作成</span><span class="sxs-lookup"><span data-stu-id="7a1a2-105">Create the add-in project</span></span>

[!include[note about Yeoman generator bug](../includes/note-yeoman-generator-bug-201908.md)]

<span data-ttu-id="7a1a2-106">Yeoman ジェネレーターを使用して、Excel アドイン プロジェクトを作成します。</span><span class="sxs-lookup"><span data-stu-id="7a1a2-106">Use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="7a1a2-107">次のコマンドを実行し、以下のプロンプトに応答します。</span><span class="sxs-lookup"><span data-stu-id="7a1a2-107">Run the following command and then answer the prompts as follows:</span></span>

```command&nbsp;line
yo office
```

- <span data-ttu-id="7a1a2-108">**Choose a project type: (プロジェクトの種類を選択)** `Office Add-in Task Pane project using Angular framework`</span><span class="sxs-lookup"><span data-stu-id="7a1a2-108">**Choose a project type:** `Office Add-in Task Pane project using Angular framework`</span></span>
- <span data-ttu-id="7a1a2-109">**Choose a script type: (スクリプトの種類を選択)** `TypeScript`</span><span class="sxs-lookup"><span data-stu-id="7a1a2-109">**Choose a script type:** `TypeScript`</span></span>
- <span data-ttu-id="7a1a2-110">**What would you want to name your add-in?: (アドインの名前を何にしますか)**</span><span class="sxs-lookup"><span data-stu-id="7a1a2-110">**What do you want to name your add-in?**</span></span> `my-office-add-in`
- <span data-ttu-id="7a1a2-111">**Which Office client application would you like to support?: (どの Office クライアント アプリケーションをサポートしますか)**</span><span class="sxs-lookup"><span data-stu-id="7a1a2-111">**Which Office client application would you like to support?**</span></span> `Excel`

<span data-ttu-id="7a1a2-112">ウィザードを完了すると、ジェネレーターによってプロジェクトが作成されて、サポートしているノード コンポーネントがインストールされます。</span><span class="sxs-lookup"><span data-stu-id="7a1a2-112">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

## <a name="explore-the-project"></a><span data-ttu-id="7a1a2-113">プロジェクトを確認する</span><span class="sxs-lookup"><span data-stu-id="7a1a2-113">Explore the project</span></span>

<span data-ttu-id="7a1a2-114">Yeoman ジェネレーターで作成したアドイン プロジェクトには、とても基本的な作業ウィンドウ アドインのサンプル コードが含まれています。</span><span class="sxs-lookup"><span data-stu-id="7a1a2-114">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span> <span data-ttu-id="7a1a2-115">アドイン プロジェクトの主要な構成要素を確認したい場合は、コード エディターでプロジェクトを開き、以下に一覧表示されているファイルを確認します。</span><span class="sxs-lookup"><span data-stu-id="7a1a2-115">If you'd like to explore the key components of your add-in project, open the project in your code editor and review the files listed below.</span></span> <span data-ttu-id="7a1a2-116">アドインを試す準備ができたら、次のセクションに進みます。</span><span class="sxs-lookup"><span data-stu-id="7a1a2-116">When you're ready to try out your add-in, proceed to the next section.</span></span>

- <span data-ttu-id="7a1a2-117">プロジェクトのルート ディレクトリにある **manifest.xml** ファイルで、アドインの機能と設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="7a1a2-117">The **manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="7a1a2-118">**./src/taskpane/app/app.component.html** ファイルには、作業ウィンドウ用の HTML マークアップが含まれています。</span><span class="sxs-lookup"><span data-stu-id="7a1a2-118">The **./src/taskpane/app/app.component.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="7a1a2-119">**./src/taskpane/taskpane.css** ファイルには、作業ウィンドウ内のコンテンツに適用される CSS が含まれています。</span><span class="sxs-lookup"><span data-stu-id="7a1a2-119">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="7a1a2-120">**./src/taskpane/app/app.component.ts** ファイルには、作業ウィンドウと Excel の間のやり取りを容易にする Office JavaScript API コードが含まれています。</span><span class="sxs-lookup"><span data-stu-id="7a1a2-120">The **./src/taskpane/app/app.component.ts** file contains the Office JavaScript API code that facilitates interaction between the task pane and Excel.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="7a1a2-121">試してみる</span><span class="sxs-lookup"><span data-stu-id="7a1a2-121">Try it out</span></span>

1. <span data-ttu-id="7a1a2-122">プロジェクトのルート フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="7a1a2-122">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "my-office-add-in"
    ```

2. [!include[Start server section](../includes/quickstart-yo-start-server-excel.md)] 

3. <span data-ttu-id="7a1a2-123">Excel で、**[ホーム]** タブを選択し、リボンの **[作業ウィンドウの表示]** ボタンをクリックして、アドインの作業ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="7a1a2-123">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Excel アドイン ボタン](../images/excel-quickstart-addin-3b.png)

4. <span data-ttu-id="7a1a2-125">ワークシート内で任意のセルの範囲を選択します。</span><span class="sxs-lookup"><span data-stu-id="7a1a2-125">Select any range of cells in the worksheet.</span></span>

5. <span data-ttu-id="7a1a2-126">作業ウィンドウの下部で、**[実行]** リンクを選択して、選択範囲の色を黄色に設定します。</span><span class="sxs-lookup"><span data-stu-id="7a1a2-126">At the bottom of the task pane, choose the **Run** link to set the color of the selected range to yellow.</span></span>

    ![Excel アドイン](../images/excel-quickstart-addin-3c.png)

## <a name="next-steps"></a><span data-ttu-id="7a1a2-128">次の手順</span><span class="sxs-lookup"><span data-stu-id="7a1a2-128">Next steps</span></span>

<span data-ttu-id="7a1a2-129">おめでとうございます! これで Angular を使用して Excel 作業ウィンドウ アドインを作成できました。</span><span class="sxs-lookup"><span data-stu-id="7a1a2-129">Congratulations, you've successfully created an Excel task pane add-in using Angular!</span></span> <span data-ttu-id="7a1a2-130">次に、Excel アドインの機能の詳細について説明します。Excel アドインのチュートリアルに従って、より複雑なアドインをビルドします。</span><span class="sxs-lookup"><span data-stu-id="7a1a2-130">Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="7a1a2-131">Excel アドインのチュートリアル</span><span class="sxs-lookup"><span data-stu-id="7a1a2-131">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="7a1a2-132">関連項目</span><span class="sxs-lookup"><span data-stu-id="7a1a2-132">See also</span></span>

* [<span data-ttu-id="7a1a2-133">Excel アドインのチュートリアル</span><span class="sxs-lookup"><span data-stu-id="7a1a2-133">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="7a1a2-134">Excel JavaScript API を使用した基本的なプログラミングの概念</span><span class="sxs-lookup"><span data-stu-id="7a1a2-134">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="7a1a2-135">Excel アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="7a1a2-135">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="7a1a2-136">Excel JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="7a1a2-136">Excel JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
