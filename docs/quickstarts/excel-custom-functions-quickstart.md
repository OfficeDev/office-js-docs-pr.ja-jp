---
ms.date: 09/18/2019
description: Excel クイックスタートガイドでのカスタム関数の開発。
title: カスタム関数のクイックスタート
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: f34a8817a7c8ef2679fc8ce0a6ad17cec600531b
ms.sourcegitcommit: a0257feabcfe665061c14b8bdb70cf82f7aca414
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/18/2019
ms.locfileid: "37035330"
---
# <a name="get-started-developing-excel-custom-functions"></a><span data-ttu-id="bee4a-103">Excel カスタム関数の開発を始める</span><span class="sxs-lookup"><span data-stu-id="bee4a-103">Get started developing Excel custom functions</span></span>

<span data-ttu-id="bee4a-104">カスタム関数を使用すると、開発者は、JavaScript または Typescript でアドインの一部として定義することによって、Excel に新しい関数を追加できるようになります。</span><span class="sxs-lookup"><span data-stu-id="bee4a-104">With custom functions, developers can now add new functions to Excel by defining them in JavaScript or Typescript as part of an add-in.</span></span> <span data-ttu-id="bee4a-105">Excel ユーザーは、Excel の任意のネイティブ関数の場合と同じように、カスタム`SUM()`関数にアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="bee4a-105">Excel users can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="bee4a-106">前提条件</span><span class="sxs-lookup"><span data-stu-id="bee4a-106">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

* <span data-ttu-id="bee4a-107">Windows 上の Excel (バージョン1904以降、Office 365 サブスクリプションに接続されている) または web 上の Excel</span><span class="sxs-lookup"><span data-stu-id="bee4a-107">Excel on Windows (version 1904 or later, connected to Office 365 subscription) or Excel on the web</span></span>
* <span data-ttu-id="bee4a-108">Excel カスタム関数は Office on Mac でサポートされています (Office 365 サブスクリプションに接続されています)。また、このチュートリアルへの更新はまもなく公開されます。</span><span class="sxs-lookup"><span data-stu-id="bee4a-108">Excel custom functions are supported in Office on Mac (connected to Office 365 subscription) and an update to this tutorial is forthcoming.</span></span>

>[!NOTE]
><span data-ttu-id="bee4a-109">Excel カスタム関数は Office 2019 (1 回限りの購入) ではサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="bee4a-109">Excel custom functions are not supported in Office 2019 (one-time purchase).</span></span>

## <a name="build-your-first-custom-functions-project"></a><span data-ttu-id="bee4a-110">最初のカスタム関数プロジェクトを作成する</span><span class="sxs-lookup"><span data-stu-id="bee4a-110">Build your first custom functions project</span></span>

<span data-ttu-id="bee4a-111">はじめに、Yeoman ジェネレーターを使って、カスタム関数プロジェクトを作成します。</span><span class="sxs-lookup"><span data-stu-id="bee4a-111">To start, you'll use the Yeoman generator to create the custom functions project.</span></span> <span data-ttu-id="bee4a-112">これにより、カスタム関数のコーディングを開始するための正しいフォルダー構造、ソース ファイル、依存関係によるプロジェクトがセットアップされます。</span><span class="sxs-lookup"><span data-stu-id="bee4a-112">This will set up your project with the correct folder structure, source files, and dependencies to begin coding your custom functions.</span></span>

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

    - <span data-ttu-id="bee4a-113">**Choose a project type: (プロジェクトの種類を選択)** `Excel Custom Functions Add-in project`</span><span class="sxs-lookup"><span data-stu-id="bee4a-113">**Choose a project type:** `Excel Custom Functions Add-in project`</span></span>
    - <span data-ttu-id="bee4a-114">**Choose a script type: (スクリプトの種類を選択)** `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="bee4a-114">**Choose a script type:** `JavaScript`</span></span>
    - <span data-ttu-id="bee4a-115">**What would you want to name your add-in?: (アドインの名前を何にしますか)**</span><span class="sxs-lookup"><span data-stu-id="bee4a-115">**What do you want to name your add-in?**</span></span> `starcount`

    ![カスタム関数の Office アドイン用の Yeoman ジェネレーターのプロンプト](../images/starcountPrompt.png)

    <span data-ttu-id="bee4a-117">Yeoman ジェネレーターはプロジェクト ファイルを作成し、サポートしているノード コンポーネントをインストールします。</span><span class="sxs-lookup"><span data-stu-id="bee4a-117">The Yeoman generator will create the project files and install supporting Node components.</span></span>

2. <span data-ttu-id="bee4a-118">[ごみ箱] ジェネレーターでは、プロジェクトの処理に関するいくつかの命令がコマンドラインに表示されますが、それらは無視して、手順に従って続行します。</span><span class="sxs-lookup"><span data-stu-id="bee4a-118">The Yeoman generator will give you some instructions in your command line about what to do with the project, but ignore them and continue to follow our instructions.</span></span> <span data-ttu-id="bee4a-119">プロジェクトのルート フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="bee4a-119">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd starcount
    ```

3. <span data-ttu-id="bee4a-120">プロジェクトをビルドします。</span><span class="sxs-lookup"><span data-stu-id="bee4a-120">Build the project.</span></span> 

    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > <span data-ttu-id="bee4a-121">開発の最中でも、OfficeアドインはHTTPではなくHTTPSを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="bee4a-121">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="bee4a-122">`npm run build`の実行後に証明書をインストールするように指示が出された場合は、Yeomanジェネレーターが提供する証明書をインストールする手順に従ってください。</span><span class="sxs-lookup"><span data-stu-id="bee4a-122">If you are prompted to install a certificate after you run `npm run build`, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

4. <span data-ttu-id="bee4a-123">Node.js で実行しているローカル Web サーバーを開始します。</span><span class="sxs-lookup"><span data-stu-id="bee4a-123">Start the local web server, which runs in Node.js.</span></span> <span data-ttu-id="bee4a-124">Web または Windows 上の Excel でカスタム関数アドインを試すことができます。</span><span class="sxs-lookup"><span data-stu-id="bee4a-124">You can try out the custom function add-in in Excel on the web or Windows.</span></span> <span data-ttu-id="bee4a-125">アドインの作業ウィンドウを開くように求められる場合がありますが、これはオプションです。</span><span class="sxs-lookup"><span data-stu-id="bee4a-125">You may be prompted to open the add-in's task pane, although this is optional.</span></span> <span data-ttu-id="bee4a-126">アドインの作業ウィンドウを開かなくても、カスタム関数を実行できます。</span><span class="sxs-lookup"><span data-stu-id="bee4a-126">You can still run your custom functions without opening your add-in's task pane.</span></span>

# <a name="excel-on-windowstabexcel-windows"></a>[<span data-ttu-id="bee4a-127">Windows 上の Excel</span><span class="sxs-lookup"><span data-stu-id="bee4a-127">Excel on Windows</span></span>](#tab/excel-windows)

<span data-ttu-id="bee4a-128">Windows の Excel でアドインをテストするには、次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="bee4a-128">To test your add-in in Excel on Windows, run the following command.</span></span> <span data-ttu-id="bee4a-129">このコマンドを実行すると、ローカル web サーバーが起動し、アドインが読み込まれた状態で Excel が開きます。</span><span class="sxs-lookup"><span data-stu-id="bee4a-129">When you run this command, the local web server will start and Excel will open with your add-in loaded.</span></span>

```command&nbsp;line
npm run start:desktop
```

# <a name="excel-on-the-webtabexcel-online"></a>[<span data-ttu-id="bee4a-130">Excel on the web</span><span class="sxs-lookup"><span data-stu-id="bee4a-130">Excel on the web</span></span>](#tab/excel-online)

<span data-ttu-id="bee4a-131">Web 上の Excel でアドインをテストするには、次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="bee4a-131">To test your add-in in Excel on the web, run the following command.</span></span> <span data-ttu-id="bee4a-132">このコマンドを実行すると、ローカル Web サーバーが起動します。</span><span class="sxs-lookup"><span data-stu-id="bee4a-132">When you run this command, the local web server will start.</span></span>

```command&nbsp;line
npm run start:web
```

<span data-ttu-id="bee4a-133">カスタム関数アドインを使用するには、ブラウザー上の Excel で新しいブックを開きます。</span><span class="sxs-lookup"><span data-stu-id="bee4a-133">To use your custom functions add-in, open a new workbook in Excel on a browser.</span></span> <span data-ttu-id="bee4a-134">このブックでは、次の手順を実行して、アドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="bee4a-134">In this workbook, complete the following steps to sideload your add-in.</span></span>

1. <span data-ttu-id="bee4a-135">Excel で、[**挿入**] タブを選択し、[**アドイン**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="bee4a-135">In Excel, choose the **Insert** tab and then choose **Add-ins**.</span></span>

   ![[個人用アドイン] アイコンが強調表示されている web 上の Excel にリボンを挿入する](../images/excel-cf-online-register-add-in-1.png)
   
2. <span data-ttu-id="bee4a-137">**[マイ アドインの管理]** を選択し、**[マイ アドインのアップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="bee4a-137">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="bee4a-138">**[参照...]** を選択し、Yeoman ジェネレーターによって作成されたプロジェクトのルート ディレクトリに移動します。</span><span class="sxs-lookup"><span data-stu-id="bee4a-138">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="bee4a-139">**manifest.xml** ファイルを選択し、**[開く]** を選択し、**[アップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="bee4a-139">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

---

## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="bee4a-140">あらかじめ用意されているカスタム関数を試す</span><span class="sxs-lookup"><span data-stu-id="bee4a-140">Try out a prebuilt custom function</span></span>

<span data-ttu-id="bee4a-141">[ごみ箱] ジェネレーターを使用して作成したカスタム関数プロジェクトには、 **/src/functions/functions.js**ファイル内で定義されているいくつかのあらかじめ用意されたカスタム関数があります。</span><span class="sxs-lookup"><span data-stu-id="bee4a-141">The custom functions project that you created by using the Yeoman generator contains some prebuilt custom functions, defined within the **./src/functions/functions.js** file.</span></span> <span data-ttu-id="bee4a-142">プロジェクトのルートディレクトリの **./manifest¥ xml**ファイルは、すべてのカスタム関数が`CONTOSO`名前空間に属することを指定します。</span><span class="sxs-lookup"><span data-stu-id="bee4a-142">The **./manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="bee4a-143">Excel ブックで、次の手順を`ADD`実行してカスタム関数を試してみます。</span><span class="sxs-lookup"><span data-stu-id="bee4a-143">In your Excel workbook, try out the `ADD` custom function by completing the following steps:</span></span>

1. <span data-ttu-id="bee4a-144">セルを選択し、 `=CONTOSO`テキストを入力します。</span><span class="sxs-lookup"><span data-stu-id="bee4a-144">Select a cell and type `=CONTOSO`.</span></span> <span data-ttu-id="bee4a-145">`CONTOSO` 名前空間にあるすべての関数がオートコンプリート メニューに一覧表示されます。</span><span class="sxs-lookup"><span data-stu-id="bee4a-145">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="bee4a-146">セルに`CONTOSO.ADD`値`=CONTOSO.ADD(10,200)`を入力し`10` 、 `200` enter キーを押して、数値と入力パラメーターを使用して、関数を実行します。</span><span class="sxs-lookup"><span data-stu-id="bee4a-146">Run the `CONTOSO.ADD` function, using numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="bee4a-147">`ADD` カスタム関数によって、入力パラメーターとして指定した 2 つの数字の合計が計算されます。</span><span class="sxs-lookup"><span data-stu-id="bee4a-147">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters.</span></span> <span data-ttu-id="bee4a-148">「`=CONTOSO.ADD(10,200)`」と入力して Enter キーを押すと、**210** という結果が生成されるはずです。</span><span class="sxs-lookup"><span data-stu-id="bee4a-148">Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="next-steps"></a><span data-ttu-id="bee4a-149">次の手順</span><span class="sxs-lookup"><span data-stu-id="bee4a-149">Next steps</span></span>

<span data-ttu-id="bee4a-150">おめでとうございます。 Excel アドインでカスタム関数が正常に作成されました。</span><span class="sxs-lookup"><span data-stu-id="bee4a-150">Congratulations, you've successfully created a custom function in an Excel add-in!</span></span> <span data-ttu-id="bee4a-151">次に、ストリーミングデータ機能を使用して、より複雑なアドインをビルドします。</span><span class="sxs-lookup"><span data-stu-id="bee4a-151">Next, build a more complex add-in with streaming data capability.</span></span> <span data-ttu-id="bee4a-152">次のリンクでは、「カスタム関数を使用した Excel アドインのチュートリアル」の次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="bee4a-152">The following link takes you through the next steps in the Excel add-in with custom functions tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="bee4a-153">Excel カスタム関数アドインのチュートリアル</span><span class="sxs-lookup"><span data-stu-id="bee4a-153">Excel custom functions add-in tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md#create-a-custom-function-that-requests-data-from-the-web
)

## <a name="see-also"></a><span data-ttu-id="bee4a-154">関連項目</span><span class="sxs-lookup"><span data-stu-id="bee4a-154">See also</span></span>

* [<span data-ttu-id="bee4a-155">カスタム関数の概要</span><span class="sxs-lookup"><span data-stu-id="bee4a-155">Custom functions overview</span></span>](../excel/custom-functions-overview.md)
* [<span data-ttu-id="bee4a-156">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="bee4a-156">Custom functions metadata</span></span>](../excel/custom-functions-json.md)
* [<span data-ttu-id="bee4a-157">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="bee4a-157">Runtime for Excel custom functions</span></span>](../excel/custom-functions-runtime.md)