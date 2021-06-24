---
title: Angular を使用して Excel 作業ウィンドウ アドインをビルドする
description: Office JS API と Angular を使用して単純な Excel 作業ウィンドウ アドインを作成する方法について説明します。
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: d843a74b3542df8dbc462ae2876179de7b42a2d2
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076932"
---
# <a name="build-an-excel-task-pane-add-in-using-angular"></a><span data-ttu-id="2d0ab-103">Angular を使用して Excel 作業ウィンドウ アドインをビルドする</span><span class="sxs-lookup"><span data-stu-id="2d0ab-103">Build an Excel task pane add-in using Angular</span></span>

<span data-ttu-id="2d0ab-104">この記事では、Angular と Excel JavaScript API を使用して Excel 作業ウィンドウ アドインを構築するプロセスについて説明します。</span><span class="sxs-lookup"><span data-stu-id="2d0ab-104">In this article, you'll walk through the process of building an Excel task pane add-in using Angular and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="2d0ab-105">前提条件</span><span class="sxs-lookup"><span data-stu-id="2d0ab-105">Prerequisites</span></span>

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a><span data-ttu-id="2d0ab-106">アドイン プロジェクトの作成</span><span class="sxs-lookup"><span data-stu-id="2d0ab-106">Create the add-in project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="2d0ab-107">**Choose a project type: (プロジェクトの種類を選択)** `Office Add-in Task Pane project using Angular framework`</span><span class="sxs-lookup"><span data-stu-id="2d0ab-107">**Choose a project type:** `Office Add-in Task Pane project using Angular framework`</span></span>
- <span data-ttu-id="2d0ab-108">**Choose a script type: (スクリプトの種類を選択)** `TypeScript`</span><span class="sxs-lookup"><span data-stu-id="2d0ab-108">**Choose a script type:** `TypeScript`</span></span>
- <span data-ttu-id="2d0ab-109">**What would you want to name your add-in?: (アドインの名前を何にしますか)**</span><span class="sxs-lookup"><span data-stu-id="2d0ab-109">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="2d0ab-110">**Which Office client application would you like to support?: (どの Office クライアント アプリケーションをサポートしますか)**</span><span class="sxs-lookup"><span data-stu-id="2d0ab-110">**Which Office client application would you like to support?**</span></span> `Excel`

![プロジェクトの種類が Angular フレームワークに設定されている Yeoman Office アドイン ジェネレーター コマンドライン インターフェイスのスクリーンショット。](../images/yo-office-excel-angular-2.png)

<span data-ttu-id="2d0ab-112">ウィザードを完了すると、ジェネレーターによってプロジェクトが作成されて、サポートしているノード コンポーネントがインストールされます。</span><span class="sxs-lookup"><span data-stu-id="2d0ab-112">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a><span data-ttu-id="2d0ab-113">プロジェクトを確認する</span><span class="sxs-lookup"><span data-stu-id="2d0ab-113">Explore the project</span></span>

<span data-ttu-id="2d0ab-114">Yeoman ジェネレーターで作成したアドイン プロジェクトには、とても基本的な作業ウィンドウ アドインのサンプル コードが含まれています。</span><span class="sxs-lookup"><span data-stu-id="2d0ab-114">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span> <span data-ttu-id="2d0ab-115">アドイン プロジェクトの主要な構成要素を確認したい場合は、コード エディターでプロジェクトを開き、以下に一覧表示されているファイルを確認します。</span><span class="sxs-lookup"><span data-stu-id="2d0ab-115">If you'd like to explore the key components of your add-in project, open the project in your code editor and review the files listed below.</span></span> <span data-ttu-id="2d0ab-116">アドインを試す準備ができたら、次のセクションに進みます。</span><span class="sxs-lookup"><span data-stu-id="2d0ab-116">When you're ready to try out your add-in, proceed to the next section.</span></span>

- <span data-ttu-id="2d0ab-117">プロジェクトのルート ディレクトリにある **manifest.xml** ファイルで、アドインの機能と設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="2d0ab-117">The **manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="2d0ab-118">**./src/taskpane/app/app.component.html** ファイルには、作業ウィンドウ用の HTML マークアップが含まれています。</span><span class="sxs-lookup"><span data-stu-id="2d0ab-118">The **./src/taskpane/app/app.component.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="2d0ab-119">**./src/taskpane/taskpane.css** ファイルには、作業ウィンドウ内のコンテンツに適用される CSS が含まれています。</span><span class="sxs-lookup"><span data-stu-id="2d0ab-119">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="2d0ab-120">**./src/taskpane/app/app.component.ts** ファイルには、作業ウィンドウと Excel の間のやり取りを容易にする Office JavaScript API コードが含まれています。</span><span class="sxs-lookup"><span data-stu-id="2d0ab-120">The **./src/taskpane/app/app.component.ts** file contains the Office JavaScript API code that facilitates interaction between the task pane and Excel.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="2d0ab-121">試してみる</span><span class="sxs-lookup"><span data-stu-id="2d0ab-121">Try it out</span></span>

1. <span data-ttu-id="2d0ab-122">プロジェクトのルート フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="2d0ab-122">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. [!include[Start server section](../includes/quickstart-yo-start-server-excel.md)] 

3. <span data-ttu-id="2d0ab-123">Excel で、**[ホーム]** タブを選択し、リボンの **[作業ウィンドウの表示]** ボタンをクリックして、アドインの作業ウィンドウを開きます。</span><span class="sxs-lookup"><span data-stu-id="2d0ab-123">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![[作業ウィンドウの表示] ボタンが強調表示されている Excel ホーム メニューのスクリーンショット。](../images/excel-quickstart-addin-3b.png)

4. <span data-ttu-id="2d0ab-125">ワークシート内で任意のセルの範囲を選択します。</span><span class="sxs-lookup"><span data-stu-id="2d0ab-125">Select any range of cells in the worksheet.</span></span>

5. <span data-ttu-id="2d0ab-126">作業ウィンドウの下部で、**[実行]** リンクを選択して、選択範囲の色を黄色に設定します。</span><span class="sxs-lookup"><span data-stu-id="2d0ab-126">At the bottom of the task pane, choose the **Run** link to set the color of the selected range to yellow.</span></span>

    ![アドイン作業ウィンドウが開いており、アドイン作業ウィンドウで [実行] ボタンが強調表示されている Excel のスクリーンショット。](../images/excel-quickstart-addin-3c.png)

## <a name="next-steps"></a><span data-ttu-id="2d0ab-128">次の手順</span><span class="sxs-lookup"><span data-stu-id="2d0ab-128">Next steps</span></span>

<span data-ttu-id="2d0ab-p102">おめでとうございます! これで Angular を使用して Excel 作業ウィンドウ アドインを作成できました。次に、Excel アドインの機能の詳細について説明します。Excel アドインのチュートリアルに従って、より複雑なアドインをビルドします。</span><span class="sxs-lookup"><span data-stu-id="2d0ab-p102">Congratulations, you've successfully created an Excel task pane add-in using Angular! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="2d0ab-131">Excel アドインのチュートリアル</span><span class="sxs-lookup"><span data-stu-id="2d0ab-131">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="2d0ab-132">関連項目</span><span class="sxs-lookup"><span data-stu-id="2d0ab-132">See also</span></span>

* [<span data-ttu-id="2d0ab-133">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="2d0ab-133">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="2d0ab-134">Office アドインを開発する</span><span class="sxs-lookup"><span data-stu-id="2d0ab-134">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
* [<span data-ttu-id="2d0ab-135">Office アドインの Excel JavaScript オブジェクト モデル</span><span class="sxs-lookup"><span data-stu-id="2d0ab-135">Excel JavaScript object model in Office Add-ins</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="2d0ab-136">Excel アドインのコード サンプル</span><span class="sxs-lookup"><span data-stu-id="2d0ab-136">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="2d0ab-137">Excel JavaScript API リファレンス</span><span class="sxs-lookup"><span data-stu-id="2d0ab-137">Excel JavaScript API reference</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
