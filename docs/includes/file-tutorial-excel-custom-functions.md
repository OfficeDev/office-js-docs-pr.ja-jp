# <a name="tutorial-create-custom-functions-in-excel"></a><span data-ttu-id="0b3fc-101">チュートリアル: Excel でのカスタム関数の作成</span><span class="sxs-lookup"><span data-stu-id="0b3fc-101">Create custom functions in Excel (Preview)</span></span>

## <a name="introduction"></a><span data-ttu-id="0b3fc-102">概要</span><span class="sxs-lookup"><span data-stu-id="0b3fc-102">Introduction</span></span>

<span data-ttu-id="0b3fc-103">カスタム関数では、関数をアドインの一部として JavaScript で定義することによって、Excel に新しい関数を追加できます。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-103">Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="0b3fc-104">ユーザーは Excel 内から、`SUM()` などの Excel のあらゆるネイティブ関数の場合と同じようにカスタム関数にアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-104">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="0b3fc-105">ユーザー定義の計算のような単純なタスク、または Web からワークシートへのデータのリアルタイム ストリーミングのようなより複雑なタスクを実行するカスタム関数を作成できます。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-105">You can create custom functions that perform simple tasks such as custom calculations or more complex tasks such as streaming real-time data from the web into a worksheet.</span></span>

<span data-ttu-id="0b3fc-106">このチュートリアルの内容:</span><span class="sxs-lookup"><span data-stu-id="0b3fc-106">In this tutorial you will use Visual Studio Code.</span></span>
> [!div class="checklist"]
> * <span data-ttu-id="0b3fc-107">Yo Office ジェネレーターを使用してカスタム関数プロジェクトを作成する</span><span class="sxs-lookup"><span data-stu-id="0b3fc-107">Create a custom functions project by using the Yo Office generator</span></span>
> * <span data-ttu-id="0b3fc-108">あらかじめ用意されているカスタム関数を使用し、単純な計算を実行する</span><span class="sxs-lookup"><span data-stu-id="0b3fc-108">Use a prebuilt custom function to perform a simple calculation</span></span>
> * <span data-ttu-id="0b3fc-109">Web からデータを要求するカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="0b3fc-109">Create a custom function that requests data from the web</span></span>
> * <span data-ttu-id="0b3fc-110">Web からデータをリアルタイムでストリーミングするカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="0b3fc-110">Create a custom function that streams real-time data from the web</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="prerequisites"></a><span data-ttu-id="0b3fc-111">前提条件</span><span class="sxs-lookup"><span data-stu-id="0b3fc-111">Prerequisites</span></span>

* [<span data-ttu-id="0b3fc-112">Node.js と npm</span><span class="sxs-lookup"><span data-stu-id="0b3fc-112">Node and npm</span></span>](https://nodejs.org/en/)

* <span data-ttu-id="0b3fc-113">[Git バッシュ](https://git-scm.com/downloads) (または別の Git クライアント)</span><span class="sxs-lookup"><span data-stu-id="0b3fc-113">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

* <span data-ttu-id="0b3fc-114">[Yeoman](http://yeoman.io/) と [Yo Office ジェネレーター](https://www.npmjs.com/package/generator-office)の最新版。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-114">The latest version of [Yeoman](http://yeoman.io/) and the [Yo Office generator](https://www.npmjs.com/package/generator-office).</span></span> <span data-ttu-id="0b3fc-115">以上のツールをグローバルにインストールするには、コマンド プロンプトから次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-115">To install these tools globally, run the following command via the command prompt:</span></span>

    ```bash
    npm install -g yo generator-office
    ```

* <span data-ttu-id="0b3fc-116">Windows 版 Excel (バージョン 1810 以降) または Excel Online</span><span class="sxs-lookup"><span data-stu-id="0b3fc-116">Excel for Windows (version 1810 or later) or Excel Online</span></span>

* <span data-ttu-id="0b3fc-117">[Office Insider プログラム](https://products.office.com/office-insider)に加入する (**Insider** レベル -- 以前は "Insider Fast" と呼ばれていたもの)</span><span class="sxs-lookup"><span data-stu-id="0b3fc-117">Join the [Office Insider program](https://products.office.com/office-insider) (**Insider** level -- formerly called "Insider Fast")</span></span>

## <a name="create-a-custom-functions-project"></a><span data-ttu-id="0b3fc-118">カスタム関数プロジェクトを作成する</span><span class="sxs-lookup"><span data-stu-id="0b3fc-118">Create a custom functions project</span></span>

<span data-ttu-id="0b3fc-119">このチュートリアルでは最初に、Yo Office ジェネレーターを使用し、カスタム関数プロジェクトに必要なファイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-119">You’ll begin this tutorial by using the Yo Office generator to create the files that you need for your custom functions project.</span></span>

1. <span data-ttu-id="0b3fc-120">次のコマンドを実行し、以下のようにプロンプトに応答します。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-120">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    * <span data-ttu-id="0b3fc-121">Choose a project type (プロジェクトの種類を選択): `Excel Custom Functions Add-in project (...)`</span><span class="sxs-lookup"><span data-stu-id="0b3fc-121">Choose a project type  </span></span>
    * <span data-ttu-id="0b3fc-122">Choose a script type (スクリプトの種類を選択): `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="0b3fc-122">Choose a script type  </span></span>
    * <span data-ttu-id="0b3fc-123">What would you want to name your add-in? (アドインの名前を何にしますか)</span><span class="sxs-lookup"><span data-stu-id="0b3fc-123">What do you want to name your add-in?</span></span> `stock-ticker`

    ![カスタム関数の Yo Office バッシュ プロンプト](../images/yo-office-cfs-stock-ticker-3.png)

    <span data-ttu-id="0b3fc-125">ウィザードを完了すると、ジェネレーターによってプロジェクト ファイルが作成され、サポート ノード コンポーネントがインストールされます。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-125">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span> <span data-ttu-id="0b3fc-126">プロジェクト ファイルは [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) GitHub リポジトリにあります。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-126">The project files come from the [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) GitHub repository.</span></span>

2. <span data-ttu-id="0b3fc-127">プロジェクト フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-127">Navigate to the project folder:</span></span>

    ```bash
    cd stock-ticker
    ```

3. <span data-ttu-id="0b3fc-128">ローカル Web サーバーを開始します。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-128">Start the local web server.</span></span>

    * <span data-ttu-id="0b3fc-129">Windows 版 Excel を使用してカスタム関数をテストする場合、次のコマンドを実行してローカル Web サーバーを開始し、Excel を起動し、アドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-129">If you'll be using Excel for Windows to test your custom functions, run the following command to start the local web server, launch Excel, and sideload the add-in:</span></span>

        ```bash
        npm start
        ```

    * <span data-ttu-id="0b3fc-130">Excel Online を使用してカスタム関数をテストする場合、次のコマンドを実行してローカル Web サーバーを開始します。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-130">If you'll be using Excel Online to test your custom functions, run the following command to start the local web server:</span></span> 

        ```bash
        npm run start-web
        ```

## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="0b3fc-131">あらかじめ用意されているカスタム関数をテストする</span><span class="sxs-lookup"><span data-stu-id="0b3fc-131">Try out a prebuilt custom function</span></span>

<span data-ttu-id="0b3fc-132">Yo Office ジェネレーターで作成したカスタム関数プロジェクトには、あらかじめ用意されているカスタム関数がいくつか含まれており、**src/customfunction.js** ファイル内で定義されています。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-132">The custom functions project that you created by using the Yo Office generator contains some prebuilt custom functions, defined within the **src/customfunction.js** file.</span></span> <span data-ttu-id="0b3fc-133">プロジェクトのルート ディレクトリの **manifest.xml** ファイルによって、カスタム関数はすべて `CONTOSO` 名前空間に属することが指定されます。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-133">The **manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="0b3fc-134">あらかじめ用意されているカスタム関数を使用する前に、Excel でカスタム関数アドインを登録する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-134">Before you can use any of the prebuilt custom functions, you must register the custom functions add-in in Excel.</span></span> <span data-ttu-id="0b3fc-135">そのためには、このチュートリアルで使用しているプラットフォームの場合、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-135">Do so by completing steps for the platform that you'll be using in this tutorial.</span></span>

* <span data-ttu-id="0b3fc-136">Windows 版 Excel を使用してカスタム関数をテストする場合:</span><span class="sxs-lookup"><span data-stu-id="0b3fc-136">If you'll be using Excel for Windows to test your custom functions:</span></span>

    1. <span data-ttu-id="0b3fc-137">Excel で [**挿入**] タブを選択し、[**個人用アドイン**] の右にある下向き矢印を選択します。![[個人用アドイン] 矢印が強調表示されている Windows 版 Excel の [挿入] リボン](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="0b3fc-137">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

    2. <span data-ttu-id="0b3fc-138">使用可能なアドインの一覧から [**開発者向けアドイン**] を見つけ、[**Excel カスタム関数**] アドインを選択して登録します。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-138">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.</span></span>
        <span data-ttu-id="0b3fc-139">![[個人用アドイン] 一覧で [Excel カスタム関数] アドインが強調表示されている Windows 版 Excel の [挿入] リボン](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="0b3fc-139">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

* <span data-ttu-id="0b3fc-140">Excel Online を使用してカスタム関数をテストする場合:</span><span class="sxs-lookup"><span data-stu-id="0b3fc-140">If you'll be using Excel Online to test your custom functions:</span></span> 

    1. <span data-ttu-id="0b3fc-141">Excel Online で [**挿入**] タブを選択し、[**アドイン**] を選択します。![[個人用アドイン] アイコンが強調表示されている Excel Online の [挿入] リボン](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="0b3fc-141">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

    2. <span data-ttu-id="0b3fc-142">[**マイ アドインの管理**] を選択し、[**マイ アドインのアップロード**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-142">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

    3. <span data-ttu-id="0b3fc-143">[**参照...**] を選択し、Yo Office ジェネレーターによって作成されたプロジェクトのルート ディレクトリに移動します。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-143">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

    4. <span data-ttu-id="0b3fc-144">ファイル **manifest.xml** を選択し、[**開く**] を選択し、[**アップロード**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-144">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

<span data-ttu-id="0b3fc-145">この時点で、プロジェクトにあらかじめ用意されているカスタム関数が読み込まれており、Excel 内で使用できます。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-145">At this point, the prebuilt custom functions in your project are loaded and available within Excel.</span></span> <span data-ttu-id="0b3fc-146">Excel で次の手順を実行し、`ADD` カスタム関数を試してみてください。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-146">Try out the `ADD` custom function by completing the following steps in Excel:</span></span>

1. <span data-ttu-id="0b3fc-147">セル内に「**=CONTOSO**」と入力します。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-147">Within a cell, type **=CONTOSO**.</span></span> <span data-ttu-id="0b3fc-148">`CONTOSO` 名前空間にあるすべての関数がオートコンプリート メニューに一覧表示されます。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-148">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="0b3fc-149">`CONTOSO.ADD` 関数を実行します。入力パラメーターとして `10` と `200` をセル内で指定し、Enter キーを押します。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-149">Run the `CONTOSO.ADD` function, with numbers `10` and `200` as input parameters, by specifying the following value in the cell and pressing enter:</span></span>

    ```
    =CONTOSO.ADD(10,200)
    ```

<span data-ttu-id="0b3fc-150">`ADD` カスタム関数によって、入力パラメーターとして指定した 2 つの数字の合計が計算されます。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-150">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters.</span></span> <span data-ttu-id="0b3fc-151">「`=CONTOSO.ADD(10,200)`」と入力して Enter キーを押すと、**210** という結果が生成されるはずです。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-151">Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a><span data-ttu-id="0b3fc-152">Web からデータを要求するカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="0b3fc-152">Create a custom function that requests data from the web</span></span>

<span data-ttu-id="0b3fc-153">API に株価を要求し、ワークシートのセルに結果を表示する関数が必要になった場合、どうすればよいでしょうか。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-153">What if you needed a function that could request the price of a stock from an API and display the result in the cell of a worksheet?</span></span> <span data-ttu-id="0b3fc-154">カスタム関数は、Web にデータを非同期で簡単に要求できるように設計されています。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-154">Custom functions are designed so that you can easily request data from the web asynchronously.</span></span>

<span data-ttu-id="0b3fc-155">次の手順を実行し、銘柄コード (**MSFT** など) を受け取り、その株価を返す、`stockPrice` という名前のカスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-155">Complete the following steps to create a custom function named `stockPrice` that accepts a stock ticker (e.g., **MSFT**) and returns the price of that stock.</span></span> <span data-ttu-id="0b3fc-156">このカスタム関数では、IEX Trading API が使用されます。これは無料であり、認証を必要としません。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-156">This custom function uses the IEX Trading API, which is free and does not require authentication.</span></span>

1. <span data-ttu-id="0b3fc-157">Yo Office ジェネレーターによって作成された**銘柄コード** プロジェクトで、ファイル **src/customfunctions.js** を見つけ、それをコード エディターで開きます。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-157">In the **stock-ticker** project that the Yo Office generator created, find the file **src/customfunctions.js** and open it in your code editor.</span></span>

2. <span data-ttu-id="0b3fc-158">次のコードを **customfunctions.js** に追加し、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-158">Add the following code to **home.js** and save the file.</span></span>

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

    CustomFunctionMappings.STOCKPRICE = stockPrice;
    ```

3. <span data-ttu-id="0b3fc-159">Excel のエンドユーザーがこの新しい関数を使用できるようにするには、この関数について説明するメタデータを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-159">Before Excel can make this new function available to end-users, you must specify metadata that describes this function.</span></span> <span data-ttu-id="0b3fc-160">Yo Office ジェネレーターによって作成された**銘柄コード** プロジェクトで、ファイル **config/customfunctions.json** を見つけ、それをコード エディターで開きます。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-160">In the **stock-ticker** project that the Yo Office generator created, find the file **config/customfunctions.json** and open it in your code editor.</span></span> <span data-ttu-id="0b3fc-161">**config/customfunctions.json** ファイル内の `functions` 配列に次のオブジェクトを追加し、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-161">Add the following object to the `functions` array within the **config/customfunctions.json** file and save the file.</span></span>

    <span data-ttu-id="0b3fc-162">この JSON では、`stockPrice` 関数について説明しています。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-162">This JSON describes the `stockPrice` function.</span></span>

    ```json
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
                "description": "stock ticker name",
                "type": "string",
                "dimensionality": "scalar"
            }
        ]
    }
    ```

4. <span data-ttu-id="0b3fc-163">新しい関数をエンドユーザーが使用できるようにするには、Excel にアドインを登録する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-163">You must reregister the add-in in Excel in order for the new function to be available to end-users.</span></span> <span data-ttu-id="0b3fc-164">このチュートリアルで使用しているプラットフォームの場合、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-164">Complete the following steps for the platform that you're using in this tutorial.</span></span>

    * <span data-ttu-id="0b3fc-165">Windows 版 Excel を使用する場合:</span><span class="sxs-lookup"><span data-stu-id="0b3fc-165">If you're using Excel for Windows:</span></span>

        1. <span data-ttu-id="0b3fc-166">Excel を閉じて再び開きます。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-166">Close Excel and then reopen Excel.</span></span>

        2. <span data-ttu-id="0b3fc-167">Excel で [**挿入**] タブを選択し、[**個人用アドイン**] の右にある下向き矢印を選択します。![[個人用アドイン] 矢印が強調表示されている Windows 版 Excel の [挿入] リボン](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="0b3fc-167">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

        1. <span data-ttu-id="0b3fc-168">使用可能なアドインの一覧から [**開発者向けアドイン**] を見つけ、[**Excel カスタム関数**] アドインを選択して登録します。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-168">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.</span></span>
            <span data-ttu-id="0b3fc-169">![[個人用アドイン] 一覧で [Excel カスタム関数] アドインが強調表示されている Windows 版 Excel の [挿入] リボン](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="0b3fc-169">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

    * <span data-ttu-id="0b3fc-170">Excel Online を使用する場合:</span><span class="sxs-lookup"><span data-stu-id="0b3fc-170">If you're using Excel Online:</span></span> 

        1. <span data-ttu-id="0b3fc-171">Excel Online で [**挿入**] タブを選択し、[**アドイン**] を選択します。![[個人用アドイン] アイコンが強調表示されている Excel Online の [挿入] リボン](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="0b3fc-171">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

        2. <span data-ttu-id="0b3fc-172">[**マイ アドインの管理**] を選択し、[**マイ アドインのアップロード**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-172">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

        3. <span data-ttu-id="0b3fc-173">[**参照...**] を選択し、Yo Office ジェネレーターによって作成されたプロジェクトのルート ディレクトリに移動します。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-173">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

        4. <span data-ttu-id="0b3fc-174">ファイル **manifest.xml** を選択し、[**開く**] を選択し、[**アップロード**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-174">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

5. <span data-ttu-id="0b3fc-175">それでは、新しい関数を試してみましょう。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-175">Now, let's try out the new function.</span></span> <span data-ttu-id="0b3fc-176">セル **B1** にテキスト `=CONTOSO.STOCKPRICE("MSFT")` を入力し、Enter キーを押します。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-176">In cell **B1**, type the text `=CONTOSO.STOCKPRICE("MSFT")` and press enter.</span></span> <span data-ttu-id="0b3fc-177">セル **B1** の結果が Microsoft の最新株価になっているはずです。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-177">You should see that the result in cell **B1** is the current stock price for one share of Microsoft stock.</span></span>

## <a name="create-a-streaming-asynchronous-custom-function"></a><span data-ttu-id="0b3fc-178">非同期でデータをストリーミングするカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="0b3fc-178">Create a streaming asynchronous custom function</span></span>

<span data-ttu-id="0b3fc-179">作成した `stockPrice` 関数では、特定の時点での株価が返されますが、株価は常に変動するものです。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-179">The `stockPrice` function that you just created returns the price of a stock at a specific moment in time, but stock prices are always changing.</span></span> <span data-ttu-id="0b3fc-180">API からデータをストリーミングし、株価をリアルタイム更新するカスタム関数を作成しましょう。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-180">Let's create a custom function that streams data from an API to get real-time updates on a stock price.</span></span>

<span data-ttu-id="0b3fc-181">次の手順を実行し、(前の要求が完了しているという条件で) 1,000 ミリ秒ごとに指定の株価を要求する、`stockPriceStream` という名前のカスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-181">Complete the following steps to create a custom function named `stockPriceStream` that requests the price of the specified stock every 1000 milliseconds (provided that the previous request has completed).</span></span> <span data-ttu-id="0b3fc-182">最初の要求が進行中のとき、関数が呼び出されているセルに **#GETTING_DATA** というプレースホルダー値が表示されることがあります。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-182">While the initial request is in-progress, you may see the placeholder value **#GETTING_DATA** the cell where the function is being called.</span></span> <span data-ttu-id="0b3fc-183">関数によって値が返されると、そのセルの **#GETTING_DATA** がその値で置換されます。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-183">When a value is returned by the function, **#GETTING_DATA** will be replaced by that value in the cell.</span></span>

1. <span data-ttu-id="0b3fc-184">Yo Office ジェネレーターによって作成された**銘柄コード** プロジェクトで、次のコードを **src/customfunctions.js** に追加し、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-184">In the **stock-ticker** project that the Yo Office generator created, add the following code to **src/customfunctions.js** and save the file.</span></span>

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

    CustomFunctionMappings.STOCKPRICESTREAM = stockPriceStream;
    ```

2. <span data-ttu-id="0b3fc-185">Excel のエンドユーザーがこの新しい関数を使用できるようにするには、この関数について説明するメタデータを指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-185">Before Excel can make this new function available to end-users, you must specify metadata that describes this function.</span></span> <span data-ttu-id="0b3fc-186">Yo Office ジェネレーターによって作成された**銘柄コード** プロジェクトで、**config/customfunctions.json** ファイル内の `functions` 配列に次のオブジェクトを追加し、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-186">In the **stock-ticker** project that the Yo Office generator created, add the following object to the `functions` array within the **config/customfunctions.json** file and save the file.</span></span>

    <span data-ttu-id="0b3fc-187">この JSON では、`stockPriceStream` 関数について説明しています。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-187">This JSON describes the `stockPriceStream` function.</span></span> <span data-ttu-id="0b3fc-188">ストリーミング関数の場合、このコード サンプルで示すように、`options` オブジェクト内で `stream` プロパティと `cancelable` プロパティを `true` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-188">For any streaming function, the `stream` property and the `cancelable` property must be set to `true` within the `options` object, as shown in this code sample.</span></span>

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
                "description": "stock ticker name",
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

3. <span data-ttu-id="0b3fc-189">新しい関数をエンドユーザーが使用できるようにするには、Excel にアドインを登録する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-189">You must reregister the add-in in Excel in order for the new function to be available to end-users.</span></span> <span data-ttu-id="0b3fc-190">このチュートリアルで使用しているプラットフォームの場合、次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-190">Complete the following steps for the platform that you're using in this tutorial.</span></span>

    * <span data-ttu-id="0b3fc-191">Windows 版 Excel を使用する場合:</span><span class="sxs-lookup"><span data-stu-id="0b3fc-191">If you're using Excel for Windows:</span></span>

        1. <span data-ttu-id="0b3fc-192">Excel を閉じて再び開きます。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-192">Close Excel and then reopen Excel.</span></span>
        
        2. <span data-ttu-id="0b3fc-193">Excel で [**挿入**] タブを選択し、[**個人用アドイン**] の右にある下向き矢印を選択します。![[個人用アドイン] 矢印が強調表示されている Windows 版 Excel の [挿入] リボン](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="0b3fc-193">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

        3. <span data-ttu-id="0b3fc-194">使用可能なアドインの一覧から [**開発者向けアドイン**] を見つけ、[**Excel カスタム関数**] アドインを選択して登録します。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-194">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.</span></span>
            <span data-ttu-id="0b3fc-195">![[個人用アドイン] 一覧で [Excel カスタム関数] アドインが強調表示されている Windows 版 Excel の [挿入] リボン](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="0b3fc-195">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

    * <span data-ttu-id="0b3fc-196">Excel Online を使用する場合:</span><span class="sxs-lookup"><span data-stu-id="0b3fc-196">If you're using Excel Online:</span></span> 

        1. <span data-ttu-id="0b3fc-197">Excel Online で [**挿入**] タブを選択し、[**アドイン**] を選択します。![[個人用アドイン] アイコンが強調表示されている Excel Online の [挿入] リボン](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="0b3fc-197">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

        2. <span data-ttu-id="0b3fc-198">[**マイ アドインの管理**] を選択し、[**マイ アドインのアップロード**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-198">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

        3. <span data-ttu-id="0b3fc-199">[**参照...**] を選択し、Yo Office ジェネレーターによって作成されたプロジェクトのルート ディレクトリに移動します。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-199">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

        4. <span data-ttu-id="0b3fc-200">ファイル **manifest.xml** を選択し、[**開く**] を選択し、[**アップロード**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-200">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

4. <span data-ttu-id="0b3fc-201">それでは、新しい関数を試してみましょう。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-201">Now, let's try out the new function.</span></span> <span data-ttu-id="0b3fc-202">セル **C1** にテキスト `=CONTOSO.STOCKPRICESTREAM("MSFT")` を入力し、Enter キーを押します。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-202">In cell **C1**, type the text `=CONTOSO.STOCKPRICESTREAM("MSFT")` and press enter.</span></span> <span data-ttu-id="0b3fc-203">株式市場が開いている場合、セル **C1** の結果が継続的に更新され、Microsoft の株価がリアルタイムで反映されます。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-203">Provided that the stock market is open, you should see that the result in cell **C1** is constantly updated to reflect the real-time price for one share of Microsoft stock.</span></span>

## <a name="next-steps"></a><span data-ttu-id="0b3fc-204">次の手順</span><span class="sxs-lookup"><span data-stu-id="0b3fc-204">Next steps</span></span>

<span data-ttu-id="0b3fc-205">このチュートリアルでは、新しいカスタム関数プロジェクトを作成し、あらかじめ用意されている関数を試し、Web にデータを要求するカスタム関数を作成し、Web からデータをリアルタイムでストリーミングするカスタム関数を作成しました。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-205">In this tutorial, you've created a new custom functions project, tried out a prebuilt function, created a custom function that requests data from the web, and created a custom function that streams real-time data from the web.</span></span> <span data-ttu-id="0b3fc-206">Excel のカスタム関数に関する詳細については、次の記事にお進みください。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-206">To learn more about custom functions in Excel, continue to the following article:</span></span> 

> [!div class="nextstepaction"]
> [<span data-ttu-id="0b3fc-207">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="0b3fc-207">Create custom functions in Excel (Preview)</span></span>](../excel/custom-functions-overview.md)

## <a name="legal-information"></a><span data-ttu-id="0b3fc-208">法的情報</span><span class="sxs-lookup"><span data-stu-id="0b3fc-208">Legal information</span></span>

<span data-ttu-id="0b3fc-209">データは [IEX](https://iextrading.com/developer/) より無料提供されました。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-209">Data provided free by [IEX](https://iextrading.com/developer/).</span></span> <span data-ttu-id="0b3fc-210">[IEX の利用規約](https://iextrading.com/api-exhibit-a/)をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-210">View [IEX's Terms of Use](https://iextrading.com/api-exhibit-a/).</span></span> <span data-ttu-id="0b3fc-211">Microsoft はこのチュートリアルで IEX API を教育目的でのみ使用しています。</span><span class="sxs-lookup"><span data-stu-id="0b3fc-211">Microsoft's use of the IEX API in this tutorial is for educational purposes only.</span></span>
