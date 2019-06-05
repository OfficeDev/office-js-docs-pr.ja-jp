---
title: Excel カスタム関数のチュートリアル
description: このチュートリアルでは、計算の実行、Web データの要求、Web データのストリームが可能なカスタム関数を含む Excel アドインを作成します。
ms.date: 05/30/2019
ms.prod: excel
ms.topic: tutorial
localization_priority: Normal
ms.openlocfilehash: f167125fcc24e47f0805d6c46e5338455d94b277
ms.sourcegitcommit: 567aa05d6ee6b3639f65c50188df2331b7685857
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/04/2019
ms.locfileid: "34706373"
---
# <a name="tutorial-create-custom-functions-in-excel"></a><span data-ttu-id="10cfc-103">チュートリアル: Excel でのカスタム関数の作成</span><span class="sxs-lookup"><span data-stu-id="10cfc-103">Tutorial: Create custom functions in Excel</span></span>

<span data-ttu-id="10cfc-104">カスタム関数では、関数をアドインの一部として JavaScript で定義することによって、Excel に新しい関数を追加できます。</span><span class="sxs-lookup"><span data-stu-id="10cfc-104">Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="10cfc-105">ユーザーは、Excel 内から、`SUM()` などの Excel のあらゆるネイティブ関数の場合と同じようにカスタム関数にアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="10cfc-105">Users within Excel can access custom functions as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="10cfc-106">計算のような単純なタスク、または Web からワークシートへのデータのリアルタイム ストリーミングのようなより複雑なタスクを実行するカスタム関数を作成できます。</span><span class="sxs-lookup"><span data-stu-id="10cfc-106">You can create custom functions that perform simple tasks like calculations or more complex tasks such as streaming real-time data from the web into a worksheet.</span></span>

<span data-ttu-id="10cfc-107">このチュートリアルの内容:</span><span class="sxs-lookup"><span data-stu-id="10cfc-107">In this tutorial, you will:</span></span>
> [!div class="checklist"]
> * <span data-ttu-id="10cfc-108">[Office アドイン用の Yeoman ジェネレーター](https://www.npmjs.com/package/generator-office)を使用して、カスタム関数アドインを作成します。</span><span class="sxs-lookup"><span data-stu-id="10cfc-108">Create a custom function add-in using the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office).</span></span> 
> * <span data-ttu-id="10cfc-109">あらかじめ用意されているカスタム関数を使用し、単純な計算を実行します。</span><span class="sxs-lookup"><span data-stu-id="10cfc-109">Use a prebuilt custom function to perform a simple calculation.</span></span>
> * <span data-ttu-id="10cfc-110">Web からデータを取得するカスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="10cfc-110">Create a custom function that gets data from the web.</span></span>
> * <span data-ttu-id="10cfc-111">Web からデータをリアルタイムでストリーミングするカスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="10cfc-111">Create a custom function that streams real-time data from the web.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="10cfc-112">前提条件</span><span class="sxs-lookup"><span data-stu-id="10cfc-112">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

* <span data-ttu-id="10cfc-113">Excel on Windows (バージョン1810以降) または Excel Online</span><span class="sxs-lookup"><span data-stu-id="10cfc-113">Excel on Windows (version 1810 or later) or Excel Online</span></span>

## <a name="create-a-custom-functions-project"></a><span data-ttu-id="10cfc-114">カスタム関数プロジェクトを作成する</span><span class="sxs-lookup"><span data-stu-id="10cfc-114">Create a custom functions project</span></span>

 <span data-ttu-id="10cfc-115">まず、カスタム関数アドインをビルドするコード プロジェクトを作成します。</span><span class="sxs-lookup"><span data-stu-id="10cfc-115">To start, you'll create the code project to build your custom function add-in.</span></span> <span data-ttu-id="10cfc-116">[Office アドイン用の [ごみ箱] ジェネレーター](https://www.npmjs.com/package/generator-office)では、プロジェクトに事前に用意されているカスタム関数を使用してセットアップし、試すことができます。カスタム関数のクイックスタートを既に実行してプロジェクトを生成した場合は、そのプロジェクトを引き続き使用して、[この手順](#create-a-custom-function-that-requests-data-from-the-web)に進んでください。</span><span class="sxs-lookup"><span data-stu-id="10cfc-116">The [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) will set up your project with some prebuilt custom functions that you can try out. If you have already run the custom functions quick start and generated a project, continue to use that project and skip to [this step](#create-a-custom-function-that-requests-data-from-the-web) instead.</span></span>

1. <span data-ttu-id="10cfc-117">次のコマンドを実行し、以下のようにプロンプトに応答します。</span><span class="sxs-lookup"><span data-stu-id="10cfc-117">Run the following command and then answer the prompts as follows.</span></span>
    
    ```command&nbsp;line
    yo office
    ```
    
    * <span data-ttu-id="10cfc-118">**Choose a project type: (プロジェクトの種類を選択)** `Excel Custom Functions Add-in project`</span><span class="sxs-lookup"><span data-stu-id="10cfc-118">**Choose a project type:** `Excel Custom Functions Add-in project`</span></span>
    * <span data-ttu-id="10cfc-119">**Choose a script type: (スクリプトの種類を選択)** `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="10cfc-119">**Choose a script type:** `JavaScript`</span></span>
    * <span data-ttu-id="10cfc-120">**What would you want to name your add-in?: (アドインの名前を何にしますか)**</span><span class="sxs-lookup"><span data-stu-id="10cfc-120">**What do you want to name your add-in?**</span></span> `stock-ticker`

    ![カスタム関数の Office アドイン用の Yeoman ジェネレーターのプロンプト](../images/UpdatedYoOfficePrompt.png)
    
    <span data-ttu-id="10cfc-122">Yeoman ジェネレーターはプロジェクト ファイルを作成し、サポートしているノード コンポーネントをインストールします。</span><span class="sxs-lookup"><span data-stu-id="10cfc-122">The Yeoman generator will create the project files and install supporting Node components.</span></span>

2. <span data-ttu-id="10cfc-123">プロジェクトのルート フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="10cfc-123">Navigate to the root folder of the project.</span></span>
    
    ```command&nbsp;line
    cd stock-ticker
    ```

3. <span data-ttu-id="10cfc-124">プロジェクトをビルドします。</span><span class="sxs-lookup"><span data-stu-id="10cfc-124">Build the project.</span></span>
    
    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > <span data-ttu-id="10cfc-125">開発の最中でも、OfficeアドインはHTTPではなくHTTPSを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="10cfc-125">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="10cfc-126">`npm run build`の実行後に証明書をインストールするように指示が出された場合は、Yeomanジェネレーターが提供する証明書をインストールする手順に従ってください。</span><span class="sxs-lookup"><span data-stu-id="10cfc-126">If you are prompted to install a certificate after you run `npm run build`, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

4. <span data-ttu-id="10cfc-127">Node.js で実行しているローカル Web サーバーを開始します。</span><span class="sxs-lookup"><span data-stu-id="10cfc-127">Start the local web server, which runs in Node.js.</span></span> <span data-ttu-id="10cfc-128">カスタム関数アドインは、Windows または Excel Online で Excel で試すことができます。</span><span class="sxs-lookup"><span data-stu-id="10cfc-128">You can try out the custom function add-in in Excel on Windows or Excel Online.</span></span>

# <a name="excel-on-windowstabexcel-windows"></a>[<span data-ttu-id="10cfc-129">Windows 上の Excel</span><span class="sxs-lookup"><span data-stu-id="10cfc-129">Excel on Windows</span></span>](#tab/excel-windows)

<span data-ttu-id="10cfc-130">Windows の Excel でアドインをテストするには、次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="10cfc-130">To test your add-in in Excel on Windows, run the following command.</span></span> <span data-ttu-id="10cfc-131">このコマンドを実行すると、ローカル web サーバーが起動し、アドインが読み込まれた状態で Excel が開きます。</span><span class="sxs-lookup"><span data-stu-id="10cfc-131">When you run this command, the local web server will start and Excel will open with your add-in loaded.</span></span>

```command&nbsp;line
npm run start:desktop
```

# <a name="excel-onlinetabexcel-online"></a>[<span data-ttu-id="10cfc-132">Excel Online</span><span class="sxs-lookup"><span data-stu-id="10cfc-132">Excel Online</span></span>](#tab/excel-online)

<span data-ttu-id="10cfc-133">Excel Online でアドインをテストするには、次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="10cfc-133">To test your add-in in Excel Online, run the following command.</span></span> <span data-ttu-id="10cfc-134">このコマンドを実行すると、ローカル Web サーバーが起動します。</span><span class="sxs-lookup"><span data-stu-id="10cfc-134">When you run this command, the local web server will start.</span></span>

```command&nbsp;line
npm run start:web
```

<span data-ttu-id="10cfc-135">カスタム関数アドインを使用するには、Excel Online で新しいブックを開きます。</span><span class="sxs-lookup"><span data-stu-id="10cfc-135">To use your custom functions add-in, open a new workbook in Excel Online.</span></span> <span data-ttu-id="10cfc-136">このブックでは、次の手順を実行して、アドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="10cfc-136">In this workbook, complete the following steps to sideload your add-in.</span></span>

1. <span data-ttu-id="10cfc-137">Excel Online で、**[挿入]** タブを選択して、**[アドイン]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="10cfc-137">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.</span></span>

   ![[個人用アドイン] アイコンが強調表示された状態で Excel Online にリボンを挿入する](../images/excel-cf-online-register-add-in-1.png)
   
2. <span data-ttu-id="10cfc-139">**[マイ アドインの管理]** を選択し、**[マイ アドインのアップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="10cfc-139">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="10cfc-140">**[参照...]** を選択し、Yeoman ジェネレーターによって作成されたプロジェクトのルート ディレクトリに移動します。</span><span class="sxs-lookup"><span data-stu-id="10cfc-140">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="10cfc-141">**manifest.xml** ファイルを選択し、**[開く]** を選択し、**[アップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="10cfc-141">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

--- 
    
## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="10cfc-142">あらかじめ用意されているカスタム関数を試す</span><span class="sxs-lookup"><span data-stu-id="10cfc-142">Try out a prebuilt custom function</span></span>

<span data-ttu-id="10cfc-143">作成したカスタム関数プロジェクトには、 **/src/functions/functions.js**ファイル内で定義されたあらかじめ用意されたカスタム関数がいくつか含まれています。</span><span class="sxs-lookup"><span data-stu-id="10cfc-143">The custom functions project that you created contains some prebuilt custom functions, defined within the **./src/functions/functions.js** file.</span></span> <span data-ttu-id="10cfc-144">**./manifest.xml** ファイルによって、カスタム関数はすべて `CONTOSO` 名前空間に属することが指定されます。</span><span class="sxs-lookup"><span data-stu-id="10cfc-144">The **./manifest.xml** file specifies that all custom functions belong to the `CONTOSO` namespace.</span></span> <span data-ttu-id="10cfc-145">Excel でカスタム関数にアクセスするには、CONTOSO 名前空間を使用します。</span><span class="sxs-lookup"><span data-stu-id="10cfc-145">You'll use the CONTOSO namespace to access the custom functions in Excel.</span></span>

<span data-ttu-id="10cfc-146">その後、次の手順を実行し、`ADD` カスタム関数を試します。</span><span class="sxs-lookup"><span data-stu-id="10cfc-146">Next you'll try out the `ADD` custom function by completing the following steps:</span></span>

1. <span data-ttu-id="10cfc-147">Excel で、任意のセルに移動し、`=CONTOSO` と入力します。</span><span class="sxs-lookup"><span data-stu-id="10cfc-147">In Excel, go to any cell and enter `=CONTOSO`.</span></span> <span data-ttu-id="10cfc-148">`CONTOSO` 名前空間にあるすべての関数がオートコンプリート メニューに一覧表示されます。</span><span class="sxs-lookup"><span data-stu-id="10cfc-148">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="10cfc-149">セル内で値 `=CONTOSO.ADD(10,200)` を入力して Enter キーを押し、入力パラメーターとして数値 `10` と `200` を指定して、`CONTOSO.ADD` 関数を実行します。</span><span class="sxs-lookup"><span data-stu-id="10cfc-149">Run the `CONTOSO.ADD` function, with numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="10cfc-150">`ADD` カスタム関数によって、指定した 2 つの数字の合計が計算され、**210** という結果が返されます。</span><span class="sxs-lookup"><span data-stu-id="10cfc-150">The `ADD` custom function computes the sum of the two numbers that you provided and returns the result of **210**.</span></span>

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a><span data-ttu-id="10cfc-151">Web からデータを要求するカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="10cfc-151">Create a custom function that requests data from the web</span></span>

<span data-ttu-id="10cfc-152">Web からデータを統合することは、カスタム関数を使用して Excel を拡張する優れた方法です。</span><span class="sxs-lookup"><span data-stu-id="10cfc-152">Integrating data from the Web is a great way to extend Excel through custom functions.</span></span> <span data-ttu-id="10cfc-153">次に、Web API から株価情報を取得し、ワークシートのセルに結果を返す、`stockPrice` というカスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="10cfc-153">Next you’ll create a custom function named `stockPrice` that gets a stock quote from a Web API and returns the result to the cell of a worksheet.</span></span> <span data-ttu-id="10cfc-154">IEX Trading API を使用します。これは無料であり、認証を必要としません。</span><span class="sxs-lookup"><span data-stu-id="10cfc-154">You’ll use the IEX Trading API, which is free and does not require authentication.</span></span>

1. <span data-ttu-id="10cfc-155">**銘柄**コードプロジェクトで、 **/src/functions/functions.js**を見つけて、コードエディターで開きます。</span><span class="sxs-lookup"><span data-stu-id="10cfc-155">In the **stock-ticker** project, find the file **./src/functions/functions.js** and open it in your code editor.</span></span>

2. <span data-ttu-id="10cfc-156">**Js**で、 `increment`関数を見つけて、その関数の後に次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="10cfc-156">In **functions.js**, locate the `increment` function and add the following code after that function.</span></span>

    ```js
    /**
    * Fetches current stock price
    * @customfunction 
    * @param {string} ticker Stock symbol
    * @returns {number} The current stock price.
    */
    function stockPrice(ticker) {
        var url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price";
        return fetch(url)
            .then(function(response) {
                return response.text();
            })
            .then(function(text) {
                return parseFloat(text);
            });

        // Note: in case of an error, the returned rejected Promise
        //    will be bubbled up to Excel to indicate an error.
    }
    CustomFunctions.associate("STOCKPRICE", stockPrice);
    ```

    <span data-ttu-id="10cfc-157">`CustomFunctions.associate` コードは、JavaScript で関数の `id` と `stockPrice` の関数アドレスを関連付けて、Excel により関数を呼び出せるようにします。</span><span class="sxs-lookup"><span data-stu-id="10cfc-157">The `CustomFunctions.associate` code associates the `id` of the function with the function address of `stockPrice` in JavaScript so that Excel can call your function.</span></span>

3. <span data-ttu-id="10cfc-158">次のコマンドを実行してプロジェクトを再構築します。</span><span class="sxs-lookup"><span data-stu-id="10cfc-158">Run the following command to rebuild the project.</span></span>

    ```command&nbsp;line
    npm run build
    ```

4. <span data-ttu-id="10cfc-159">次の手順を実行して (Excel on Windows または Excel Online の場合)、Excel でアドインを再登録します。</span><span class="sxs-lookup"><span data-stu-id="10cfc-159">Complete the following steps (for either Excel on Windows or Excel Online) to re-register the add-in in Excel.</span></span> <span data-ttu-id="10cfc-160">新しい関数を使用できるようにするには、これらの手順を完了する必要があります。</span><span class="sxs-lookup"><span data-stu-id="10cfc-160">You must complete these steps before the new function will be available.</span></span> 

# <a name="excel-on-windowstabexcel-windows"></a>[<span data-ttu-id="10cfc-161">Windows 上の Excel</span><span class="sxs-lookup"><span data-stu-id="10cfc-161">Excel on Windows</span></span>](#tab/excel-windows)

1. <span data-ttu-id="10cfc-162">Excel を閉じて再び開きます。</span><span class="sxs-lookup"><span data-stu-id="10cfc-162">Close Excel and then reopen Excel.</span></span>

2. <span data-ttu-id="10cfc-163">Excel で [**挿入**] タブを選択し、[**マイ**アドイン] の右側にある下向き矢印を選択します。 ![[個人用アドイン] 矢印が強調表示されている Windows 上の Excel でのリボンの挿入](../images/select-insert.png)</span><span class="sxs-lookup"><span data-stu-id="10cfc-163">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel on Windows with the My Add-ins arrow highlighted](../images/select-insert.png)</span></span>

3. <span data-ttu-id="10cfc-164">使用可能なアドインの一覧から **[開発者向けアドイン]** セクションを見つけ、**銘柄コード** アドインを選択して登録します。</span><span class="sxs-lookup"><span data-stu-id="10cfc-164">In the list of available add-ins, find the **Developer Add-ins** section and select the **stock-ticker** add-in to register it.</span></span>
    <span data-ttu-id="10cfc-165">![[個人用アドイン] ボックスの一覧で強調表示された Excel カスタム関数アドインを使用して、Excel の Excel にリボンを挿入する](../images/list-stock-ticker-red.png)</span><span class="sxs-lookup"><span data-stu-id="10cfc-165">![Insert ribbon in Excel on Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/list-stock-ticker-red.png)</span></span>

# <a name="excel-onlinetabexcel-online"></a>[<span data-ttu-id="10cfc-166">Excel Online</span><span class="sxs-lookup"><span data-stu-id="10cfc-166">Excel Online</span></span>](#tab/excel-online)

1. <span data-ttu-id="10cfc-167">Excel Online で **[挿入]** タブを選択し、**[アドイン]** を選択します。![[個人用アドイン] アイコンが強調表示されている Excel Online の [挿入] リボン](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="10cfc-167">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

2. <span data-ttu-id="10cfc-168">**[マイ アドインの管理]** を選択し、**[マイ アドインのアップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="10cfc-168">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

3. <span data-ttu-id="10cfc-169">**[参照...]** を選択し、Yeoman ジェネレーターによって作成されたプロジェクトのルート ディレクトリに移動します。</span><span class="sxs-lookup"><span data-stu-id="10cfc-169">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span> 

4. <span data-ttu-id="10cfc-170">**manifest.xml** ファイルを選択し、**[開く]** を選択し、**[アップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="10cfc-170">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

---

<ol start="5">
<li> <span data-ttu-id="10cfc-171">新しい関数をお試しください。</span><span class="sxs-lookup"><span data-stu-id="10cfc-171">Try out the new function.</span></span> <span data-ttu-id="10cfc-172">セル <strong>B1</strong> に <strong>=CONTOSO.STOCKPRICE("MSFT")</strong> と入力し、Enter キーを押します。</span><span class="sxs-lookup"><span data-stu-id="10cfc-172">In cell <strong>B1</strong>, type the text <strong>=CONTOSO.STOCKPRICE("MSFT")</strong> and press enter.</span></span> <span data-ttu-id="10cfc-173">セル <strong>B1</strong> の結果が Microsoft の最新株価になっているはずです。</span><span class="sxs-lookup"><span data-stu-id="10cfc-173">You should see that the result in cell <strong>B1</strong> is the current stock price for one share of Microsoft stock.</span></span></li>
</ol>

## <a name="create-a-streaming-asynchronous-custom-function"></a><span data-ttu-id="10cfc-174">非同期でデータをストリーミングするカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="10cfc-174">Create a streaming asynchronous custom function</span></span>

<span data-ttu-id="10cfc-175">`stockPrice` 関数では、特定の時点での株価が返されますが、株価は常に変動するものです。</span><span class="sxs-lookup"><span data-stu-id="10cfc-175">The `stockPrice` function returns the price of a stock at a specific moment in time, but stock prices are always changing.</span></span> <span data-ttu-id="10cfc-176">次に、1000 ミリ秒ごと株価を取得する、`stockPriceStream` という名前のカスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="10cfc-176">Next you’ll create a custom function named `stockPriceStream` that gets the price of a stock every 1000 milliseconds.</span></span>

1. <span data-ttu-id="10cfc-177">**銘柄**コードプロジェクトで、次のコードを **/src/functions/functions.js**に追加し、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="10cfc-177">In the **stock-ticker** project, add the following code to **./src/functions/functions.js** and save the file.</span></span>

    ```js
    /**
    * Streams real time stock price
    * @customfunction 
    * @param {string} ticker Stock symbol
    * @param {CustomFunctions.StreamingInvocation<number>} invocation
    */
    function stockPriceStream(ticker, invocation) {
        var updateFrequency = 1000 /* milliseconds*/;
        var isPending = false;

        var timer = setInterval(function() {
            // If there is already a pending request, skip this iteration:
            if (isPending) {
                return;
            }

            var url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price";
            isPending = true;

            fetch(url)
                .then(function(response) {
                    return response.text();
                })
                .then(function(text) {
                    invocation.setResult(parseFloat(text));
                })
                .catch(function(error) {
                    invocation.setResult(error);
                })
                .then(function() {
                    isPending = false;
                });
        }, updateFrequency);

        invocation.onCanceled = () => {
            clearInterval(timer);
        };
    }
    CustomFunctions.associate("STOCKPRICESTREAM", stockPriceStream);
    ```
    
    <span data-ttu-id="10cfc-178">`CustomFunctions.associate` コードは、JavaScript で関数の `id` と `stockPriceStream` の関数アドレスを関連付けて、Excel により関数を呼び出せるようにします。</span><span class="sxs-lookup"><span data-stu-id="10cfc-178">The `CustomFunctions.associate` code associates the `id` of the function with the function address of `stockPriceStream` in JavaScript so that Excel can call your function.</span></span>
    
2. <span data-ttu-id="10cfc-179">次のコマンドを実行してプロジェクトを再構築します。</span><span class="sxs-lookup"><span data-stu-id="10cfc-179">Run the following command to rebuild the project.</span></span>

    ```command&nbsp;line
    npm run build
    ```

3. <span data-ttu-id="10cfc-180">次の手順を実行して (Excel on Windows または Excel Online の場合)、Excel でアドインを再登録します。</span><span class="sxs-lookup"><span data-stu-id="10cfc-180">Complete the following steps (for either Excel on Windows or Excel Online) to re-register the add-in in Excel.</span></span> <span data-ttu-id="10cfc-181">新しい関数を使用できるようにするには、これらの手順を完了する必要があります。</span><span class="sxs-lookup"><span data-stu-id="10cfc-181">You must complete these steps before the new function will be available.</span></span> 

# <a name="excel-on-windowstabexcel-windows"></a>[<span data-ttu-id="10cfc-182">Windows 上の Excel</span><span class="sxs-lookup"><span data-stu-id="10cfc-182">Excel on Windows</span></span>](#tab/excel-windows)

1. <span data-ttu-id="10cfc-183">Excel を閉じて再び開きます。</span><span class="sxs-lookup"><span data-stu-id="10cfc-183">Close Excel and then reopen Excel.</span></span>

2. <span data-ttu-id="10cfc-184">Excel で [**挿入**] タブを選択し、[**マイ**アドイン] の右側にある下向き矢印を選択します。 ![[個人用アドイン] 矢印が強調表示されている Windows 上の Excel でのリボンの挿入](../images/select-insert.png)</span><span class="sxs-lookup"><span data-stu-id="10cfc-184">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel on Windows with the My Add-ins arrow highlighted](../images/select-insert.png)</span></span>

3. <span data-ttu-id="10cfc-185">使用可能なアドインの一覧から **[開発者向けアドイン]** セクションを見つけ、**銘柄コード** アドインを選択して登録します。</span><span class="sxs-lookup"><span data-stu-id="10cfc-185">In the list of available add-ins, find the **Developer Add-ins** section and select the **stock-ticker** add-in to register it.</span></span>
    <span data-ttu-id="10cfc-186">![[個人用アドイン] ボックスの一覧で強調表示された Excel カスタム関数アドインを使用して、Excel の Excel にリボンを挿入する](../images/list-stock-ticker-red.png)</span><span class="sxs-lookup"><span data-stu-id="10cfc-186">![Insert ribbon in Excel on Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/list-stock-ticker-red.png)</span></span>

# <a name="excel-onlinetabexcel-online"></a>[<span data-ttu-id="10cfc-187">Excel Online</span><span class="sxs-lookup"><span data-stu-id="10cfc-187">Excel Online</span></span>](#tab/excel-online)

1. <span data-ttu-id="10cfc-188">Excel Online で **[挿入]** タブを選択し、**[アドイン]** を選択します。![[個人用アドイン] アイコンが強調表示されている Excel Online の [挿入] リボン](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="10cfc-188">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

2. <span data-ttu-id="10cfc-189">**[マイ アドインの管理]** を選択し、**[マイ アドインのアップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="10cfc-189">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="10cfc-190">**[参照...]** を選択し、Yeoman ジェネレーターによって作成されたプロジェクトのルート ディレクトリに移動します。</span><span class="sxs-lookup"><span data-stu-id="10cfc-190">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="10cfc-191">**manifest.xml** ファイルを選択し、**[開く]** を選択し、**[アップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="10cfc-191">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

--- 

<ol start="4">
<li><span data-ttu-id="10cfc-192">新しい関数をお試しください。</span><span class="sxs-lookup"><span data-stu-id="10cfc-192">Try out the new function.</span></span> <span data-ttu-id="10cfc-193">セル <strong>C1</strong> に <strong>=CONTOSO.STOCKPRICESTREAM("MSFT")</strong> と入力し、Enter キーを押します。</span><span class="sxs-lookup"><span data-stu-id="10cfc-193">In cell <strong>C1</strong>, type the text <strong>=CONTOSO.STOCKPRICESTREAM("MSFT")</strong> and press enter.</span></span> <span data-ttu-id="10cfc-194">株式市場が開いている場合、セル <strong>C1</strong> の結果が継続的に更新され、Microsoft の株価がリアルタイムで反映されます。</span><span class="sxs-lookup"><span data-stu-id="10cfc-194">Provided that the stock market is open, you should see that the result in cell <strong>C1</strong> is constantly updated to reflect the real-time price for one share of Microsoft stock.</span></span></li>
</ol>

## <a name="next-steps"></a><span data-ttu-id="10cfc-195">次のステップ</span><span class="sxs-lookup"><span data-stu-id="10cfc-195">Next steps</span></span>

<span data-ttu-id="10cfc-196">おめでとうございます。</span><span class="sxs-lookup"><span data-stu-id="10cfc-196">Congratulations!</span></span> <span data-ttu-id="10cfc-197">新しいカスタム関数プロジェクトを作成し、あらかじめ用意されている関数を試し、Web にデータを要求するカスタム関数を作成し、Web からデータをリアルタイムでストリーミングするカスタム関数を作成しました。</span><span class="sxs-lookup"><span data-stu-id="10cfc-197">You've created a new custom functions project, tried out a prebuilt function, created a custom function that requests data from the web, and created a custom function that streams real-time data from the web.</span></span> <span data-ttu-id="10cfc-198">この関数のデバッグは[、カスタム関数のデバッグ手順](../excel/custom-functions-debugging.md)を使用して実行することもできます。</span><span class="sxs-lookup"><span data-stu-id="10cfc-198">You can also try out debugging this function using [the custom function debugging instructions](../excel/custom-functions-debugging.md).</span></span> <span data-ttu-id="10cfc-199">Excel のカスタム関数に関する詳細については、次の記事にお進みください。</span><span class="sxs-lookup"><span data-stu-id="10cfc-199">To learn more about custom functions in Excel, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="10cfc-200">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="10cfc-200">Create custom functions in Excel</span></span>](../excel/custom-functions-overview.md)

### <a name="legal-information"></a><span data-ttu-id="10cfc-201">法的情報</span><span class="sxs-lookup"><span data-stu-id="10cfc-201">Legal information</span></span>

<span data-ttu-id="10cfc-202">データは [IEX](https://iextrading.com/developer/) より無料提供されました。</span><span class="sxs-lookup"><span data-stu-id="10cfc-202">Data provided free by [IEX](https://iextrading.com/developer/).</span></span> <span data-ttu-id="10cfc-203">[IEX の利用規約](https://iextrading.com/api-exhibit-a/)をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="10cfc-203">View [IEX's Terms of Use](https://iextrading.com/api-exhibit-a/).</span></span> <span data-ttu-id="10cfc-204">Microsoft はこのチュートリアルで IEX API を教育目的でのみ使用しています。</span><span class="sxs-lookup"><span data-stu-id="10cfc-204">Microsoft's use of the IEX API in this tutorial is for educational purposes only.</span></span>
