---
title: Excel カスタム関数のチュートリアル (プレビュー)
description: このチュートリアルでは、計算の実行、Web データの要求、Web データのストリームが可能なカスタム関数を含む Excel アドインを作成します。
ms.date: 01/08/2019
ms.prod: excel
ms.topic: tutorial
localization_priority: Normal
ms.openlocfilehash: 4ac735e6fc19f13859d07df6cb3d2443e6dfe2fd
ms.sourcegitcommit: a59f4e322238efa187f388a75b7709462c71e668
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/13/2019
ms.locfileid: "29982021"
---
# <a name="tutorial-create-custom-functions-in-excel-preview"></a><span data-ttu-id="e0aac-103">チュートリアル: Excel でのカスタム関数の作成 (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="e0aac-103">Tutorial: Create custom functions in Excel (preview)</span></span>

<span data-ttu-id="e0aac-104">カスタム関数では、関数をアドインの一部として JavaScript で定義することによって、Excel に新しい関数を追加できます。</span><span class="sxs-lookup"><span data-stu-id="e0aac-104">Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="e0aac-105">ユーザーは、Excel 内から、`SUM()` などの Excel のあらゆるネイティブ関数の場合と同じようにカスタム関数にアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="e0aac-105">Users within Excel can access custom functions as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="e0aac-106">計算のような単純なタスク、または Web からワークシートへのデータのリアルタイム ストリーミングのようなより複雑なタスクを実行するカスタム関数を作成できます。</span><span class="sxs-lookup"><span data-stu-id="e0aac-106">You can create custom functions that perform simple tasks like calculations or more complex tasks such as streaming real-time data from the web into a worksheet.</span></span>

<span data-ttu-id="e0aac-107">このチュートリアルの内容:</span><span class="sxs-lookup"><span data-stu-id="e0aac-107">In this tutorial, you will:</span></span>
> [!div class="checklist"]
> * <span data-ttu-id="e0aac-108">[Office アドイン用の Yeoman ジェネレーター](https://www.npmjs.com/package/generator-office)を使用して、カスタム関数アドインを作成します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-108">Create a custom function add-in using the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office).</span></span> 
> * <span data-ttu-id="e0aac-109">あらかじめ用意されているカスタム関数を使用し、単純な計算を実行します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-109">Use a prebuilt custom function to perform a simple calculation.</span></span>
> * <span data-ttu-id="e0aac-110">Web からデータを取得するカスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-110">Create a custom function that gets data from the web.</span></span>
> * <span data-ttu-id="e0aac-111">Web からデータをリアルタイムでストリーミングするカスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-111">Create a custom function that streams real-time data from the web.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="prerequisites"></a><span data-ttu-id="e0aac-112">前提条件</span><span class="sxs-lookup"><span data-stu-id="e0aac-112">Prerequisites</span></span>

* <span data-ttu-id="e0aac-113">[Node.js](https://nodejs.org/en/) (バージョン 8.0.0 以降)</span><span class="sxs-lookup"><span data-stu-id="e0aac-113">[Node.js](https://nodejs.org/en/) (version 8.0.0 or later)</span></span>

* <span data-ttu-id="e0aac-114">[Git バッシュ](https://git-scm.com/downloads) (または別の Git クライアント)</span><span class="sxs-lookup"><span data-stu-id="e0aac-114">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

* <span data-ttu-id="e0aac-115">最新バージョンの [Yeoman](https://yeoman.io/) と [Office アドイン用の Yeoman ジェネレーター](https://www.npmjs.com/package/generator-office)。これらのツールをグローバルにインストールするには、コマンド プロンプトから次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-115">The latest version of [Yeoman](https://yeoman.io/) and the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office). To install these tools globally, run the following command via the command prompt:</span></span>

    ```
    npm install -g yo generator-office
    ```

    > [!NOTE]
    > <span data-ttu-id="e0aac-116">以前に Yeoman ジェネレーターをインストールしている場合でも、npm からパッケージを最新バージョンに更新することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="e0aac-116">Even if you have previously installed the Yeoman generator, we recommend updating your package to the latest version from npm.</span></span>

* <span data-ttu-id="e0aac-117">Windows 版 Excel (64 ビット バージョン 1810 以降) または Excel Online</span><span class="sxs-lookup"><span data-stu-id="e0aac-117">Excel for Windows (64-bit version 1810 or later) or Excel Online</span></span>

* <span data-ttu-id="e0aac-118">[Office Insider プログラム](https://products.office.com/office-insider)に加入する (**Insider** レベル -- 以前は "Insider Fast" と呼ばれていたもの)</span><span class="sxs-lookup"><span data-stu-id="e0aac-118">Join the [Office Insider program](https://products.office.com/office-insider) (**Insider** level -- formerly called "Insider Fast")</span></span>

## <a name="create-a-custom-functions-project"></a><span data-ttu-id="e0aac-119">カスタム関数プロジェクトを作成する</span><span class="sxs-lookup"><span data-stu-id="e0aac-119">Create a custom functions project</span></span>

 <span data-ttu-id="e0aac-120">まず、カスタム関数アドインをビルドするコード プロジェクトを作成します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-120">To start, you'll create the code project to build your custom function add-in.</span></span> <span data-ttu-id="e0aac-121">[Yeoman Office アドイン用の Yeoman ジェネレーター](https://www.npmjs.com/package/generator-office)を使用すると、プロジェクトをセットアップして、いくつかの初期カスタム関数を試すことができます。</span><span class="sxs-lookup"><span data-stu-id="e0aac-121">The [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) will set up your project with some initial custom functions that you can try out.</span></span>

1. <span data-ttu-id="e0aac-122">次のコマンドを実行し、以下のようにプロンプトに応答します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-122">Run the following command and then answer the prompts as follows.</span></span>
    
    ```
    yo office
    ```
    
    * <span data-ttu-id="e0aac-123">Choose a project type (プロジェクトの種類を選択): `Excel Custom Functions Add-in project (...)`</span><span class="sxs-lookup"><span data-stu-id="e0aac-123">Choose a project type: `Excel Custom Functions Add-in project (...)`</span></span>
    * <span data-ttu-id="e0aac-124">Choose a script type (スクリプトの種類を選択): `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="e0aac-124">Choose a script type: `JavaScript`</span></span>
    * <span data-ttu-id="e0aac-125">What would you want to name your add-in? (アドインの名前を何にしますか)</span><span class="sxs-lookup"><span data-stu-id="e0aac-125">What do you want to name your add-in?</span></span> `stock-ticker`
    
    ![カスタム関数の Office アドイン用の Yeoman ジェネレーターのプロンプト](../images/12-10-fork-cf-pic.jpg)
    
    <span data-ttu-id="e0aac-127">Yeoman ジェネレーターはプロジェクト ファイルを作成し、サポートしている Node.js コンポーネントをインストールします。</span><span class="sxs-lookup"><span data-stu-id="e0aac-127">The Yeoman generator creates the project files and installs supporting Node.js components.</span></span>

2. <span data-ttu-id="e0aac-128">プロジェクト フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-128">Go to the project folder.</span></span>
    
    ```
    cd stock-ticker
    ```

3. <span data-ttu-id="e0aac-129">このプロジェクトを実行するために必要な自己署名証明書を信頼します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-129">Trust the self-signed certificate that is needed to run this project.</span></span> <span data-ttu-id="e0aac-130">Windows または Mac についての詳細な手順については、「[自己署名証明書を信頼済みルート証明書として追加する](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e0aac-130">For detailed instructions for either Windows or Mac, see [Adding Self Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span></span>  

4. <span data-ttu-id="e0aac-131">プロジェクトをビルドします。</span><span class="sxs-lookup"><span data-stu-id="e0aac-131">Build the project.</span></span>
    
    ```
    npm run build
    ```

5. <span data-ttu-id="e0aac-132">Node.js で実行しているローカル Web サーバーを開始します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-132">Start the local web server, which runs in Node.js.</span></span> <span data-ttu-id="e0aac-133">Windows 用 Excel または Excel Online で、カスタム関数アドインを試すことができます。</span><span class="sxs-lookup"><span data-stu-id="e0aac-133">You can try out the custom function add-in in Excel for Windows, or Excel Online.</span></span>

# <a name="excel-for-windowstabexcel-windows"></a>[<span data-ttu-id="e0aac-134">Windows 用 Excel</span><span class="sxs-lookup"><span data-stu-id="e0aac-134">Excel for Windows</span></span>](#tab/excel-windows)

<span data-ttu-id="e0aac-135">次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-135">Run the following command.</span></span>

```
npm run start
```

<span data-ttu-id="e0aac-136">このコマンドは、Web サーバーを開始し、カスタム関数アドインを Windows 用 Excel にサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="e0aac-136">This command starts the web server, and sideloads your custom function add-in into Excel for Windows.</span></span>

> [!NOTE]
> <span data-ttu-id="e0aac-137">アドインが読み込まれない場合は、手順 3 が正しく完了しているか確認してください。</span><span class="sxs-lookup"><span data-stu-id="e0aac-137">If your add-in does not load, check that you have completed step 3 properly.</span></span> <span data-ttu-id="e0aac-138">**[実行時のログ](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)** を、任意のインストールまたは実行時の問題と同様に、アドインの XML マニフェスト ファイルの問題をトラブルシューティングすることもできます。</span><span class="sxs-lookup"><span data-stu-id="e0aac-138">You can also enable **[runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)** to troubleshoot issues with your add-in's XML manifest file, as well as any installation or runtime problems.</span></span> <span data-ttu-id="e0aac-139">実行時のログの書き込み`console.log`ステートメントを検索して、問題を解決するためにログ ファイルにします。</span><span class="sxs-lookup"><span data-stu-id="e0aac-139">Runtime logging writes `console.log` statements to a log file to help you find and fix issues.</span></span>

# <a name="excel-onlinetabexcel-online"></a>[<span data-ttu-id="e0aac-140">Excel Online</span><span class="sxs-lookup"><span data-stu-id="e0aac-140">Excel Online</span></span>](#tab/excel-online)

<span data-ttu-id="e0aac-141">次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-141">Run the following command.</span></span>

```
npm run start-web
```

<span data-ttu-id="e0aac-142">このコマンドは、Web サーバーを開始します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-142">This command starts the web server.</span></span> <span data-ttu-id="e0aac-143">アドインをサイドロードするには、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-143">Use the following steps to sideload your add-in.</span></span>

<ol type="a">
   <li><span data-ttu-id="e0aac-144">Excel Online で、<strong>[挿入]</strong> タブを選択して、<strong>[アドイン]</strong> を選択します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-144">In Excel Online, choose the <strong>Insert</strong> tab and then choose <strong>Add-ins</strong>.</span></span><br/>
   <img src="../images/excel-cf-online-register-add-in-1.png" alt="Insert ribbon in Excel Online with the My Add-ins icon highlighted"></li>
   <li><span data-ttu-id="e0aac-145"><strong>[マイ アドインの管理]</strong> を選択し、<strong>[マイ アドインのアップロード]</strong> を選択します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-145">Choose <strong>Manage My Add-ins</strong> and select <strong>Upload My Add-in</strong>.</span></span></li> 
   <li><span data-ttu-id="e0aac-146"><strong>[参照...]</strong> を選択し、Yeoman ジェネレーターによって作成されたプロジェクトのルート ディレクトリに移動します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-146">Choose <strong>Browse...</strong> and navigate to the root directory of the project that the Yeoman generator created.</span></span></li> 
   <li><span data-ttu-id="e0aac-147"><strong>manifest.xml</strong> ファイルを選択し、<strong>[開く]</strong> を選択し、<strong>[アップロード]</strong> を選択します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-147">Select the file <strong>manifest.xml</strong> and choose <strong>Open</strong>, then choose <strong>Upload</strong>.</span></span></li>
</ol>

> [!NOTE]
> <span data-ttu-id="e0aac-148">アドインが読み込まれない場合は、手順 3 が正しく完了しているか確認してください。</span><span class="sxs-lookup"><span data-stu-id="e0aac-148">If your add-in does not load, check that you have completed step 3 properly.</span></span>

--- 
    
## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="e0aac-149">あらかじめ用意されているカスタム関数を試す</span><span class="sxs-lookup"><span data-stu-id="e0aac-149">Try out a prebuilt custom function</span></span>

<span data-ttu-id="e0aac-150">既に作成したカスタム関数のプロジェクトには、ADD と INCREMENT という名前のあらかじめ用意されている 2 つのカスタム機能があります。</span><span class="sxs-lookup"><span data-stu-id="e0aac-150">The custom functions project that you created alrady has two prebuilt custom functions named ADD and INCREMENT.</span></span> <span data-ttu-id="e0aac-151">これらのあらかじめ用意されている関数のコードは、**src/customfunctions.js** ファイルにあります。</span><span class="sxs-lookup"><span data-stu-id="e0aac-151">The code for these prebuilt functions is in the  **src/customfunctions.js** file.</span></span> <span data-ttu-id="e0aac-152">**./manifest.xml** ファイルによって、カスタム関数はすべて `CONTOSO` 名前空間に属することが指定されます。</span><span class="sxs-lookup"><span data-stu-id="e0aac-152">The **./manifest.xml** file specifies that all custom functions belong to the `CONTOSO` namespace.</span></span> <span data-ttu-id="e0aac-153">Excel でカスタム関数にアクセスするには、CONTOSO 名前空間を使用します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-153">You'll use the CONTOSO namespace to access the custom functions in Excel.</span></span>

<span data-ttu-id="e0aac-154">その後、次の手順を実行し、`ADD` カスタム関数を試します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-154">Next you'll try out the `ADD` custom function by completing the following steps:</span></span>

1. <span data-ttu-id="e0aac-155">Excel で、任意のセルに移動し、`=CONTOSO` と入力します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-155">In Excel, go to any cell and enter `=CONTOSO`.</span></span> <span data-ttu-id="e0aac-156">`CONTOSO` 名前空間にあるすべての関数がオートコンプリート メニューに一覧表示されます。</span><span class="sxs-lookup"><span data-stu-id="e0aac-156">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="e0aac-157">セル内で値 `=CONTOSO.ADD(10,200)` を入力して Enter キーを押し、入力パラメーターとして数値 `10` と `200` を指定して、`CONTOSO.ADD` 関数を実行します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-157">Run the `CONTOSO.ADD` function, with numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="e0aac-158">`ADD` カスタム関数によって、指定した 2 つの数字の合計が計算され、**210** という結果が返されます。</span><span class="sxs-lookup"><span data-stu-id="e0aac-158">The `ADD` custom function computes the sum of the two numbers that you provided and returns the result of **210**.</span></span>

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a><span data-ttu-id="e0aac-159">Web からデータを要求するカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="e0aac-159">Create a custom function that requests data from the web</span></span>

<span data-ttu-id="e0aac-160">Web からデータを統合することは、カスタム関数を使用して Excel を拡張する優れた方法です。</span><span class="sxs-lookup"><span data-stu-id="e0aac-160">Integrating data from the Web is a great way to extend Excel through custom functions.</span></span> <span data-ttu-id="e0aac-161">次に、Web API から株価情報を取得し、ワークシートのセルに結果を返す、`stockPrice` というカスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-161">Next you’ll create a custom function named `stockPrice` that gets a stock quote from a Web API and returns the result to the cell of a worksheet.</span></span> <span data-ttu-id="e0aac-162">IEX Trading API を使用します。これは無料であり、認証を必要としません。</span><span class="sxs-lookup"><span data-stu-id="e0aac-162">You’ll use the IEX Trading API, which is free and does not require authentication.</span></span>

1. <span data-ttu-id="e0aac-163">**銘柄コード**プロジェクトで **src/customfunctions.js** ファイルを見つけ、それをコード エディターで開きます。</span><span class="sxs-lookup"><span data-stu-id="e0aac-163">In the **stock-ticker** project, find the file **src/customfunctions.js** and open it in your code editor.</span></span>

2. <span data-ttu-id="e0aac-164">**customfunctions.js** で、`increment` 関数を見つけ、その関数の直後に次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-164">In **customfunctions.js**, locate the `increment` function and add the following code immediately after that function.</span></span>

    ```js
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

> [!NOTE]
> In the January Insiders 1901 Build, there is a bug preventing fetch calls from executing which will result in #VALUE!.
> To workaround this please use the [XMLHTTPRequest API](https://docs.microsoft.com/en-us/office/dev/add-ins/excel/custom-functions-runtime#requesting-external-data) to make the web request.

3. In **customfunctions.js**, locate the line `CustomFunctions.associate("INCREMENT", increment);`. Add the following line of code immediately after that line, and save the file.

    ```js
    CustomFunctions.associate("STOCKPRICE", stockprice);
    ```

    <span data-ttu-id="e0aac-165">`CustomFunctions.associate` コードは、JavaScript で関数の `id` と `increment` の関数アドレスを関連付けて、Excel により関数を呼び出せるようにします。</span><span class="sxs-lookup"><span data-stu-id="e0aac-165">The `CustomFunctions.associate` code associates the `id` of the function with the function address of `increment` in JavaScript so that Excel can call your function.</span></span>

    <span data-ttu-id="e0aac-166">Excel でカスタム関数を使用できるようにするには、その前にメタデータを使用してそれを記述する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e0aac-166">Before Excel can use your custom function, you need to describe it using metadata.</span></span> <span data-ttu-id="e0aac-167">以前に `associate` メソッドで使用した `id` を、他のいくつかのメタデータと共に定義する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e0aac-167">You need to define the `id` used in the `associate` method previously, along with some other metadata.</span></span>


4. <span data-ttu-id="e0aac-168">**config/customfunctions.json** ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="e0aac-168">Open the **config/customfunctions.json** file.</span></span> <span data-ttu-id="e0aac-169">'関数' 配列に次の JSON オブジェクトを追加し、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-169">Add the following JSON object to the 'functions' array and save the file.</span></span>

    ```JSON
    {
        "id": "STOCKPRICE",
        "name": "STOCKPRICE",
        "description": "Fetches current stock price",
        "helpUrl": "http://www.contoso.com/help",
        "result": {
            "type": "number",
            "dimensionality": "scalar"
        },  
        "parameters": [
            {
                "name": "ticker",
                "description": "stock symbol",
                "type": "string",
                "dimensionality": "scalar"
            }
        ]
    }
    ```

    <span data-ttu-id="e0aac-170">この JSON は、`stockPrice` 関数、そのパラメーター、それによって返される結果の種類を記述します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-170">This JSON describes the `stockPrice` function, its parameters, and the type of result it returns.</span></span>

5. <span data-ttu-id="e0aac-171">新しい関数を使用できるようにするには、Excel でアドインを再登録します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-171">Re-register the add-in in Excel so that the new function is available.</span></span> 

# <a name="excel-for-windowstabexcel-windows"></a>[<span data-ttu-id="e0aac-172">Windows 用 Excel</span><span class="sxs-lookup"><span data-stu-id="e0aac-172">Excel for Windows</span></span>](#tab/excel-windows)

1. <span data-ttu-id="e0aac-173">Excel を閉じて再び開きます。</span><span class="sxs-lookup"><span data-stu-id="e0aac-173">Close Excel and then reopen Excel.</span></span>

2. <span data-ttu-id="e0aac-174">Excel で [**挿入**] タブを選択し、[**個人用アドイン**] の右にある下向き矢印を選択します。![[個人用アドイン] 矢印が強調表示されている Windows 版 Excel の [挿入] リボン](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="e0aac-174">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

3. <span data-ttu-id="e0aac-175">使用可能なアドインの一覧から **[開発者向けアドイン]** セクションを見つけ、**銘柄コード** アドインを選択して登録します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-175">In the list of available add-ins, find the **Developer Add-ins** section and select the **stock-ticker** add-in to register it.</span></span>
    <span data-ttu-id="e0aac-176">![[個人用アドイン] 一覧で [Excel カスタム関数] アドインが強調表示されている Windows 用 Excel の [挿入] リボン](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="e0aac-176">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

# <a name="excel-onlinetabexcel-online"></a>[<span data-ttu-id="e0aac-177">Excel Online</span><span class="sxs-lookup"><span data-stu-id="e0aac-177">Excel Online</span></span>](#tab/excel-online)

1. <span data-ttu-id="e0aac-178">Excel Online で **[挿入]** タブを選択し、**[アドイン]** を選択します。![[個人用アドイン] アイコンが強調表示されている Excel Online の [挿入] リボン](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="e0aac-178">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

2. <span data-ttu-id="e0aac-179">**[マイ アドインの管理]** を選択し、**[マイ アドインのアップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-179">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

3. <span data-ttu-id="e0aac-180">**[参照...]** を選択し、Yeoman ジェネレーターによって作成されたプロジェクトのルート ディレクトリに移動します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-180">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span> 

4. <span data-ttu-id="e0aac-181">**manifest.xml** ファイルを選択し、**[開く]** を選択し、**[アップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-181">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

--- 

<ol start="6">
<li> <span data-ttu-id="e0aac-182">新しい関数をお試しください。</span><span class="sxs-lookup"><span data-stu-id="e0aac-182">Try out the new function.</span></span> <span data-ttu-id="e0aac-183">セル <strong>B1</strong> に <strong>=CONTOSO.STOCKPRICE("MSFT")</strong> と入力し、Enter キーを押します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-183">In cell <strong>B1</strong>, type the text <strong>=CONTOSO.STOCKPRICE("MSFT")</strong> and press enter.</span></span> <span data-ttu-id="e0aac-184">セル <strong>B1</strong> の結果が Microsoft の最新株価になっているはずです。</span><span class="sxs-lookup"><span data-stu-id="e0aac-184">You should see that the result in cell <strong>B1</strong> is the current stock price for one share of Microsoft stock.</span></span></li>
</ol>

## <a name="create-a-streaming-asynchronous-custom-function"></a><span data-ttu-id="e0aac-185">非同期でデータをストリーミングするカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="e0aac-185">Create a streaming asynchronous custom function</span></span>

<span data-ttu-id="e0aac-186">`stockPrice` 関数では、特定の時点での株価が返されますが、株価は常に変動するものです。</span><span class="sxs-lookup"><span data-stu-id="e0aac-186">The `stockPrice` function returns the price of a stock at a specific moment in time, but stock prices are always changing.</span></span> <span data-ttu-id="e0aac-187">次に、1000 ミリ秒ごと株価を取得する、`stockPriceStream` という名前のカスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-187">Next you’ll create a custom function named `stockPriceStream` that gets the price of a stock every 1000 milliseconds.</span></span>

1. <span data-ttu-id="e0aac-188">**銘柄コード**プロジェクトで、次のコードを **src/customfunctions.js** に追加し、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-188">In the **stock-ticker** project, add the following code to **src/customfunctions.js** and save the file.</span></span>

    ```js
    function stockPriceStream(ticker, handler) {
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
                    handler.setResult(parseFloat(text));
                })
                .catch(function(error) {
                    handler.setResult(error);
                })
                .then(function() {
                    isPending = false;
                });
        }, updateFrequency);

        handler.onCanceled = () => {
            clearInterval(timer);
        };
    }
    
    CustomFunctions.associate("STOCKPRICESTREAM", stockpricestream);
    ```
    
    <span data-ttu-id="e0aac-189">Excel でカスタム関数を使用できるようにするには、その前にメタデータを使用してそれを記述する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e0aac-189">Before Excel can use your custom function, you need to describe it using metadata.</span></span>
    
2. <span data-ttu-id="e0aac-190">**銘柄コード**プロジェクトで、**config/customfunctions.json** ファイル内の `functions` 配列に次のオブジェクトを追加し、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-190">In the **stock-ticker** project add the following object to the `functions` array within the **config/customfunctions.json** file and save the file.</span></span>
    
    ```json
    { 
        "id": "STOCKPRICESTREAM",
        "name": "STOCKPRICESTREAM",
        "description": "Streams real time stock price",
        "helpUrl": "http://www.contoso.com/help",
        "result": {
            "type": "number",
            "dimensionality": "scalar"
        },  
        "parameters": [
            {
                "name": "ticker",
                "description": "stock symbol",
                "type": "string",
                "dimensionality": "scalar"
            }
        ],
        "options": {
            "stream": true,
            "cancelable": true
        }
    }
    ```

    <span data-ttu-id="e0aac-191">この JSON は、`stockPriceStream` 関数を記述します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-191">This JSON describes the `stockPriceStream` function.</span></span> <span data-ttu-id="e0aac-192">ストリーミング関数の場合、このコード サンプルで示すように、`options` オブジェクト内で `stream` プロパティと `cancelable` プロパティを `true` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e0aac-192">For any streaming function, the `stream` property and the `cancelable` property must be set to `true` within the `options` object, as shown in this code sample.</span></span>

3. <span data-ttu-id="e0aac-193">新しい関数を使用できるようにするには、Excel でアドインを再登録します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-193">Re-register the add-in in Excel so that the new function is available.</span></span>

# <a name="excel-for-windowstabexcel-windows"></a>[<span data-ttu-id="e0aac-194">Windows 用 Excel</span><span class="sxs-lookup"><span data-stu-id="e0aac-194">Excel for Windows</span></span>](#tab/excel-windows)

1. <span data-ttu-id="e0aac-195">Excel を閉じて再び開きます。</span><span class="sxs-lookup"><span data-stu-id="e0aac-195">Close Excel and then reopen Excel.</span></span>

2. <span data-ttu-id="e0aac-196">Excel で [**挿入**] タブを選択し、[**個人用アドイン**] の右にある下向き矢印を選択します。![[個人用アドイン] 矢印が強調表示されている Windows 版 Excel の [挿入] リボン](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="e0aac-196">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

3. <span data-ttu-id="e0aac-197">使用可能なアドインの一覧から **[開発者向けアドイン]** セクションを見つけ、**銘柄コード** アドインを選択して登録します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-197">In the list of available add-ins, find the **Developer Add-ins** section and select the **stock-ticker** add-in to register it.</span></span>
    <span data-ttu-id="e0aac-198">![[個人用アドイン] 一覧で [Excel カスタム関数] アドインが強調表示されている Windows 用 Excel の [挿入] リボン](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="e0aac-198">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

# <a name="excel-onlinetabexcel-online"></a>[<span data-ttu-id="e0aac-199">Excel Online</span><span class="sxs-lookup"><span data-stu-id="e0aac-199">Excel Online</span></span>](#tab/excel-online)

1. <span data-ttu-id="e0aac-200">Excel Online で **[挿入]** タブを選択し、**[アドイン]** を選択します。![[個人用アドイン] アイコンが強調表示されている Excel Online の [挿入] リボン](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="e0aac-200">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

2. <span data-ttu-id="e0aac-201">**[マイ アドインの管理]** を選択し、**[マイ アドインのアップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-201">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="e0aac-202">**[参照...]** を選択し、Yeoman ジェネレーターによって作成されたプロジェクトのルート ディレクトリに移動します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-202">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="e0aac-203">**manifest.xml** ファイルを選択し、**[開く]** を選択し、**[アップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-203">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

--- 

<ol start="4">
<li><span data-ttu-id="e0aac-204">新しい関数をお試しください。</span><span class="sxs-lookup"><span data-stu-id="e0aac-204">Try out the new function.</span></span> <span data-ttu-id="e0aac-205">セル <strong>C1</strong> に <strong>=CONTOSO.STOCKPRICESTREAM("MSFT")</strong> と入力し、Enter キーを押します。</span><span class="sxs-lookup"><span data-stu-id="e0aac-205">In cell <strong>C1</strong>, type the text <strong>=CONTOSO.STOCKPRICESTREAM("MSFT")</strong> and press enter.</span></span> <span data-ttu-id="e0aac-206">株式市場が開いている場合、セル <strong>C1</strong> の結果が継続的に更新され、Microsoft の株価がリアルタイムで反映されます。</span><span class="sxs-lookup"><span data-stu-id="e0aac-206">Provided that the stock market is open, you should see that the result in cell <strong>C1</strong> is constantly updated to reflect the real-time price for one share of Microsoft stock.</span></span></li>
</ol>


## <a name="next-steps"></a><span data-ttu-id="e0aac-207">次の手順</span><span class="sxs-lookup"><span data-stu-id="e0aac-207">Next steps</span></span>

<span data-ttu-id="e0aac-208">おめでとうございます。</span><span class="sxs-lookup"><span data-stu-id="e0aac-208">Congratulations!</span></span> <span data-ttu-id="e0aac-209">新しいカスタム関数プロジェクトを作成し、あらかじめ用意されている関数を試し、Web にデータを要求するカスタム関数を作成し、Web からデータをリアルタイムでストリーミングするカスタム関数を作成しました。</span><span class="sxs-lookup"><span data-stu-id="e0aac-209">You've created a new custom functions project, tried out a prebuilt function, created a custom function that requests data from the web, and created a custom function that streams real-time data from the web.</span></span> <span data-ttu-id="e0aac-210">Excel のカスタム関数に関する詳細については、次の記事にお進みください。</span><span class="sxs-lookup"><span data-stu-id="e0aac-210">To learn more about custom functions in Excel, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="e0aac-211">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="e0aac-211">Create custom functions in Excel</span></span>](../excel/custom-functions-overview.md)

### <a name="legal-information"></a><span data-ttu-id="e0aac-212">法的情報</span><span class="sxs-lookup"><span data-stu-id="e0aac-212">Legal information</span></span>

<span data-ttu-id="e0aac-213">データは [IEX](https://iextrading.com/developer/) より無料提供されました。</span><span class="sxs-lookup"><span data-stu-id="e0aac-213">Data provided free by [IEX](https://iextrading.com/developer/).</span></span> <span data-ttu-id="e0aac-214">[IEX の利用規約](https://iextrading.com/api-exhibit-a/)をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="e0aac-214">View [IEX's Terms of Use](https://iextrading.com/api-exhibit-a/).</span></span> <span data-ttu-id="e0aac-215">Microsoft はこのチュートリアルで IEX API を教育目的でのみ使用しています。</span><span class="sxs-lookup"><span data-stu-id="e0aac-215">Microsoft's use of the IEX API in this tutorial is for educational purposes only.</span></span>


