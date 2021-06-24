---
ms.date: 11/09/2020
description: Excel カスタム関数開発のためのクイック スタート ガイド。
title: カスタム関数クイック スタート
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 0b0e42149e771978026db3eb84594bd172d09459
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076631"
---
# <a name="get-started-developing-excel-custom-functions"></a><span data-ttu-id="7bd13-103">Excel カスタム関数の開発を開始する</span><span class="sxs-lookup"><span data-stu-id="7bd13-103">Get started developing Excel custom functions</span></span>

<span data-ttu-id="7bd13-104">カスタム関数機能により、開発者は、アドインの一部としてカスタム関数を JavaScript または Typescript で定義することによって、新しい関数を Excel に追加できるようになりました。</span><span class="sxs-lookup"><span data-stu-id="7bd13-104">With custom functions, developers can now add new functions to Excel by defining them in JavaScript or Typescript as part of an add-in.</span></span> <span data-ttu-id="7bd13-105">Excel のユーザーは、`SUM()` など、Excel のすべてのネイティブ関数にアクセスするとの同じようにカスタム関数にアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="7bd13-105">Excel users can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="7bd13-106">前提条件</span><span class="sxs-lookup"><span data-stu-id="7bd13-106">Prerequisites</span></span>

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

* <span data-ttu-id="7bd13-107">Windows 版 Excel (Microsoft 365 サブスクリプションに接続されている、バージョン 1904 以降) または Excel on the web</span><span class="sxs-lookup"><span data-stu-id="7bd13-107">Excel on Windows (version 1904 or later, connected to a Microsoft 365 subscription) or Excel on the web</span></span>
* <span data-ttu-id="7bd13-108">Excel カスタム関数は (Microsoft 365 サブスクリプションに接続されている) Mac 版 Office でサポートされており、このチュートリアルはまもなく更新されます。</span><span class="sxs-lookup"><span data-stu-id="7bd13-108">Excel custom functions are supported in Office on Mac (connected to a Microsoft 365 subscription) and an update to this tutorial is forthcoming.</span></span>

>[!NOTE]
><span data-ttu-id="7bd13-109">Excel カスタム関数は Office 2019 (1 回限りの購入) ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="7bd13-109">Excel custom functions are not supported in Office 2019 (one-time purchase).</span></span>

## <a name="build-your-first-custom-functions-project"></a><span data-ttu-id="7bd13-110">カスタム関数プロジェクトを初めて作成する</span><span class="sxs-lookup"><span data-stu-id="7bd13-110">Build your first custom functions project</span></span>

<span data-ttu-id="7bd13-111">はじめに、Yeoman ジェネレーターを使って、カスタム関数プロジェクトを作成します。</span><span class="sxs-lookup"><span data-stu-id="7bd13-111">To start, you'll use the Yeoman generator to create the custom functions project.</span></span> <span data-ttu-id="7bd13-112">これにより、カスタム関数のコーディングを開始するための正しいフォルダー構造、ソース ファイル、依存関係によるプロジェクトがセットアップされます。</span><span class="sxs-lookup"><span data-stu-id="7bd13-112">This will set up your project with the correct folder structure, source files, and dependencies to begin coding your custom functions.</span></span>

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

    - <span data-ttu-id="7bd13-113">**Choose a project type: (プロジェクトの種類を選択)** `Excel Custom Functions Add-in project`</span><span class="sxs-lookup"><span data-stu-id="7bd13-113">**Choose a project type:** `Excel Custom Functions Add-in project`</span></span>
    - <span data-ttu-id="7bd13-114">**Choose a script type: (スクリプトの種類を選択)** `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="7bd13-114">**Choose a script type:** `JavaScript`</span></span>
    - <span data-ttu-id="7bd13-115">**What would you want to name your add-in?: (アドインの名前を何にしますか)**</span><span class="sxs-lookup"><span data-stu-id="7bd13-115">**What do you want to name your add-in?**</span></span> `starcount`

    ![カスタム関数プロジェクトの Yeoman Office アドイン ジェネレーター コマンドライン インターフェイス プロンプトのスクリーンショット。](../images/starcountPrompt.png)

    <span data-ttu-id="7bd13-117">Yeoman ジェネレーターはプロジェクト ファイルを作成し、サポートしているノード コンポーネントをインストールします。</span><span class="sxs-lookup"><span data-stu-id="7bd13-117">The Yeoman generator will create the project files and install supporting Node components.</span></span>

2. <span data-ttu-id="7bd13-p103">Yeoman ジェネレーターによりプロジェクトの作業に関する手順がコマンド ライン内にいくつか示されますが、これらは無視し、引き続き指示に従います。プロジェクトのルート フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="7bd13-p103">The Yeoman generator will give you some instructions in your command line about what to do with the project, but ignore them and continue to follow our instructions. Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd starcount
    ```

3. <span data-ttu-id="7bd13-120">プロジェクトをビルドします。</span><span class="sxs-lookup"><span data-stu-id="7bd13-120">Build the project.</span></span> 

    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > <span data-ttu-id="7bd13-121">Office アドインは、開発中であっても HTTP ではなく HTTPS を使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="7bd13-121">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="7bd13-122">`npm run build`の実行後に証明書をインストールするように指示が出された場合は、Yeomanジェネレーターが提供する証明書をインストールする手順に従ってください。</span><span class="sxs-lookup"><span data-stu-id="7bd13-122">If you are prompted to install a certificate after you run `npm run build`, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

4. <span data-ttu-id="7bd13-123">Node.js で実行しているローカル Web サーバーを開始します。</span><span class="sxs-lookup"><span data-stu-id="7bd13-123">Start the local web server, which runs in Node.js.</span></span> <span data-ttu-id="7bd13-124">カスタム関数アドインは Web 版 Excel または Windows 版 Excel で試すことができます。</span><span class="sxs-lookup"><span data-stu-id="7bd13-124">You can try out the custom function add-in in Excel on the web or Windows.</span></span> <span data-ttu-id="7bd13-125">アドインの作業ウィンドウを開くように求められる場合がありますが、これは省略可能です。</span><span class="sxs-lookup"><span data-stu-id="7bd13-125">You may be prompted to open the add-in's task pane, although this is optional.</span></span> <span data-ttu-id="7bd13-126">カスタム関数はアドインの作業ウィンドウを開かなくても実行できます。</span><span class="sxs-lookup"><span data-stu-id="7bd13-126">You can still run your custom functions without opening your add-in's task pane.</span></span>

# <a name="excel-on-windows"></a>[<span data-ttu-id="7bd13-127">Windows 版 Excel</span><span class="sxs-lookup"><span data-stu-id="7bd13-127">Excel on Windows</span></span>](#tab/excel-windows)

<span data-ttu-id="7bd13-128">アドインを Windows 版 Excel で試すには、次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="7bd13-128">To test your add-in in Excel on Windows, run the following command.</span></span> <span data-ttu-id="7bd13-129">このコマンドを実行すると、ローカル Web サーバーが起動し、アドインが読み込まれた状態で Excel が開きます。</span><span class="sxs-lookup"><span data-stu-id="7bd13-129">When you run this command, the local web server will start and Excel will open with your add-in loaded.</span></span>

```command&nbsp;line
npm run start:desktop
```

# <a name="excel-on-the-web"></a>[<span data-ttu-id="7bd13-130">Web 版 Excel</span><span class="sxs-lookup"><span data-stu-id="7bd13-130">Excel on the web</span></span>](#tab/excel-online)

<span data-ttu-id="7bd13-131">アドインを Web 版 Excel で試すには、次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="7bd13-131">To test your add-in in Excel on the web, run the following command.</span></span> <span data-ttu-id="7bd13-132">このコマンドを実行すると、ローカル Web サーバーが起動します。</span><span class="sxs-lookup"><span data-stu-id="7bd13-132">When you run this command, the local web server will start.</span></span>

```command&nbsp;line
npm run start:web
```

<span data-ttu-id="7bd13-133">カスタム関数アドインを使用するには、ブラウザー上の Excel で新しいブックを開きます。</span><span class="sxs-lookup"><span data-stu-id="7bd13-133">To use your custom functions add-in, open a new workbook in Excel on a browser.</span></span> <span data-ttu-id="7bd13-134">このブックで次の手順を実行してアドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="7bd13-134">In this workbook, complete the following steps to sideload your add-in.</span></span>

1. <span data-ttu-id="7bd13-135">Excel で、[**挿入**] タブを選択して、[**アドイン**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="7bd13-135">In Excel, choose the **Insert** tab and then choose **Add-ins**.</span></span>

   ![[個人用アドイン] ボタンが強調表示された Excel on the web の [挿入] リボンのスクリーンショット。](../images/excel-cf-online-register-add-in-1.png)
   
2. <span data-ttu-id="7bd13-137">**[マイ アドインの管理]** を選択し、**[マイ アドインのアップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="7bd13-137">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="7bd13-138">**[参照...]** を選択し、Yeoman ジェネレーターによって作成されたプロジェクトのルート ディレクトリに移動します。</span><span class="sxs-lookup"><span data-stu-id="7bd13-138">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="7bd13-139">**manifest.xml** ファイルを選択し、**[開く]** を選択し、**[アップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="7bd13-139">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

---

## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="7bd13-140">既製のカスタム関数を試す</span><span class="sxs-lookup"><span data-stu-id="7bd13-140">Try out a prebuilt custom function</span></span>

<span data-ttu-id="7bd13-141">Yeoman ジェネレーター使用して作成したカスタム関数プロジェクトには既製のカスタム関数がいくつか含まれており、これらは **./src/functions/functions.js** ファイル内で定義されています。</span><span class="sxs-lookup"><span data-stu-id="7bd13-141">The custom functions project that you created by using the Yeoman generator contains some prebuilt custom functions, defined within the **./src/functions/functions.js** file.</span></span> <span data-ttu-id="7bd13-142">カスタム関数はすべて `CONTOSO` 名前空間に属するということは、プロジェクトのルート ディレクトリの **./manifest.xml** ファイルで指定されています。</span><span class="sxs-lookup"><span data-stu-id="7bd13-142">The **./manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="7bd13-143">Excel ブックで次の手順を実行し、`ADD` カスタム関数を試してみてください。</span><span class="sxs-lookup"><span data-stu-id="7bd13-143">In your Excel workbook, try out the `ADD` custom function by completing the following steps:</span></span>

1. <span data-ttu-id="7bd13-144">セルを 1 つ選択し、「`=CONTOSO`」と入力します。</span><span class="sxs-lookup"><span data-stu-id="7bd13-144">Select a cell and type `=CONTOSO`.</span></span> <span data-ttu-id="7bd13-145">`CONTOSO` 名前空間にあるすべての関数がオートコンプリート メニューに一覧表示されます。</span><span class="sxs-lookup"><span data-stu-id="7bd13-145">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="7bd13-146">セル内に「`=CONTOSO.ADD(10,200)`」という値を入力して Enter キーを押し、入力パラメーターとして数値「`10`」 と「`200`」を指定して、`CONTOSO.ADD` 関数を実行します。</span><span class="sxs-lookup"><span data-stu-id="7bd13-146">Run the `CONTOSO.ADD` function, using numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="7bd13-147">`ADD` カスタム関数によって、入力パラメーターとして指定した 2 つの数字の合計が計算されます。</span><span class="sxs-lookup"><span data-stu-id="7bd13-147">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters.</span></span> <span data-ttu-id="7bd13-148">「`=CONTOSO.ADD(10,200)`」と入力して Enter キーを押すと、**210** という結果が生成されるはずです。</span><span class="sxs-lookup"><span data-stu-id="7bd13-148">Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="next-steps"></a><span data-ttu-id="7bd13-149">次の手順</span><span class="sxs-lookup"><span data-stu-id="7bd13-149">Next steps</span></span>

<span data-ttu-id="7bd13-150">これで、カスタム関数が Excel アドイン内に正常に作成されました。</span><span class="sxs-lookup"><span data-stu-id="7bd13-150">Congratulations, you've successfully created a custom function in an Excel add-in!</span></span> <span data-ttu-id="7bd13-151">次は、ストリーミング データ機能を使用してより複雑なアドインを作成してください。</span><span class="sxs-lookup"><span data-stu-id="7bd13-151">Next, build a more complex add-in with streaming data capability.</span></span> <span data-ttu-id="7bd13-152">カスタム関数を使用した Excel アドインのチュートリアルの次の手順を確認するには、次のリンクをクリックしてください。</span><span class="sxs-lookup"><span data-stu-id="7bd13-152">The following link takes you through the next steps in the Excel add-in with custom functions tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="7bd13-153">Excel カスタム関数アドインのチュートリアル</span><span class="sxs-lookup"><span data-stu-id="7bd13-153">Excel custom functions add-in tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md#create-a-custom-function-that-requests-data-from-the-web
)

## <a name="see-also"></a><span data-ttu-id="7bd13-154">関連項目</span><span class="sxs-lookup"><span data-stu-id="7bd13-154">See also</span></span>

* [<span data-ttu-id="7bd13-155">カスタム関数の概要</span><span class="sxs-lookup"><span data-stu-id="7bd13-155">Custom functions overview</span></span>](../excel/custom-functions-overview.md)
* [<span data-ttu-id="7bd13-156">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="7bd13-156">Custom functions metadata</span></span>](../excel/custom-functions-json.md)
* [<span data-ttu-id="7bd13-157">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="7bd13-157">Runtime for Excel custom functions</span></span>](../excel/custom-functions-runtime.md)