# <a name="tutorial-create-custom-functions-in-excel"></a><span data-ttu-id="741d7-101">チュートリアル: Excel でカスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="741d7-101">Create custom functions in Excel (Preview)</span></span>

## <a name="introduction"></a><span data-ttu-id="741d7-102">概要</span><span class="sxs-lookup"><span data-stu-id="741d7-102">Introduction</span></span>

<span data-ttu-id="741d7-p101">カスタム関数とは、開発者がアドインの一部として、JavaScript でカスタム関数を定義して Excel に追加する新しい関数です。Excel 内のユーザーは、Excel の他のネイティブ関数 (`SUM()` など) と同様に、カスタム関数にアクセスできます。 ユーザー設定の計算などの単純なタスク、または Web からワークシートにリアルタイムにデータをストリーミングするといったより複雑なタスクを実行するカスタム関数を作成できます。</span><span class="sxs-lookup"><span data-stu-id="741d7-p101">Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in. Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`. You can create custom functions that perform simple tasks such as custom calculations or more complex tasks such as streaming real-time data from the web into a worksheet.</span></span>

<span data-ttu-id="741d7-106">このチュートリアルでは、以下の操作を実行します。</span><span class="sxs-lookup"><span data-stu-id="741d7-106">In this tutorial, you will use Visual Studio Code.</span></span>
> [!div class="checklist"]
> * <span data-ttu-id="741d7-107">Yo Office ジェネレーターを使用してカスタム関数プロジェクトを作成します。</span><span class="sxs-lookup"><span data-stu-id="741d7-107">Create a custom functions project by using the Yo Office generator</span></span>
> * <span data-ttu-id="741d7-108">作成済みのカスタム関数を使用して、単純な計算を実行するには</span><span class="sxs-lookup"><span data-stu-id="741d7-108">Use a prebuilt custom function to perform a simple calculation</span></span>
> * <span data-ttu-id="741d7-109">Web サイトからデータを要求するカスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="741d7-109">Create a custom function that requests data from the web</span></span>
> * <span data-ttu-id="741d7-110">Web サイトからのリアルタイムのデータをストリームするカスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="741d7-110">Create a custom function that streams real-time data from the web</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="prerequisites"></a><span data-ttu-id="741d7-111">前提条件</span><span class="sxs-lookup"><span data-stu-id="741d7-111">Prerequisites</span></span>

* [<span data-ttu-id="741d7-112">Node.js および npm</span><span class="sxs-lookup"><span data-stu-id="741d7-112">Node and npm</span></span>](https://nodejs.org/en/)

* <span data-ttu-id="741d7-113">[Git バッシュ](https://git-scm.com/downloads) (またはその他の Git クライアント)</span><span class="sxs-lookup"><span data-stu-id="741d7-113">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

* <span data-ttu-id="741d7-p102">[Yeoman](http://yeoman.io/) と [Yo Office  ジェネレーター](https://www.npmjs.com/package/generator-office)の最新バージョンです。グローバルにこれらのツールをインストールするには、コマンド プロンプトを使用して次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="741d7-p102">The latest version of [Yeoman](http://yeoman.io/) and the [Yo Office generator](https://www.npmjs.com/package/generator-office). To install these tools globally, run the following command via the command prompt:</span></span>

    ```bash
    npm install -g yo generator-office
    ```

* <span data-ttu-id="741d7-116"> Windows 版 Excel  (ビルド番号 10827 以降) または Excel Online</span><span class="sxs-lookup"><span data-stu-id="741d7-116">Excel for Windows (build number 10827 or later) or Excel Online</span></span>

* <span data-ttu-id="741d7-117">[  Office Insider プログラム](https://products.office.com/office-insider) に参加します。（\*\* Insider\*\* レベルは、以前は「Insider Fast」と呼ばれていました）</span><span class="sxs-lookup"><span data-stu-id="741d7-117">Join the [Office Insider program](https://products.office.com/office-insider) (**Insider** level -- formerly called "Insider Fast")</span></span>

## <a name="create-a-custom-functions-project"></a><span data-ttu-id="741d7-118">カスタム関数プロジェクトを作成します。</span><span class="sxs-lookup"><span data-stu-id="741d7-118">Create a custom functions project by using the Yo Office generator</span></span>

<span data-ttu-id="741d7-119">Yo Office ジェネレーターを使用するカスタム関数プロジェクトに必要なファイルを作成するこのチュートリアルを開始するでしょう。</span><span class="sxs-lookup"><span data-stu-id="741d7-119">You’ll begin this tutorial by using the Yo Office generator to create the files that you need for your custom functions project.</span></span>

1. <span data-ttu-id="741d7-120">次のコマンドを実行し、以下のプロンプトに応答します。</span><span class="sxs-lookup"><span data-stu-id="741d7-120">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    * <span data-ttu-id="741d7-121">プロジェクト タイプを選択してください `Excel Custom Functions Add-in project (...)`</span><span class="sxs-lookup"><span data-stu-id="741d7-121">Choose a project type:`Excel Custom Functions Add-in project (...)`</span></span>
    * <span data-ttu-id="741d7-122">スクリプト タイプを選択してください `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="741d7-122">Choose a script type:`JavaScript`</span></span>
    * <span data-ttu-id="741d7-123">アドインの名前を何にしますか？</span><span class="sxs-lookup"><span data-stu-id="741d7-123">What do you want to name your add-in?:</span></span> `stock-ticker`

    ![Yo Office バッシュは、カスタム関数のプロンプトを表示します](../images/yo-office-cfs-stock-ticker-3.png)

    <span data-ttu-id="741d7-p103">ウィザードを完了すると、ジェネレーターがプロジェクト ファイルを作成し、ノードのサポート コンポーネントをインストールします。プロジェクト ファイルは、 [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) の GitHub のリポジトリから取得されます。</span><span class="sxs-lookup"><span data-stu-id="741d7-p103">After you complete the wizard, the generator will create the project files and install supporting Node components. The project files come from the [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) GitHub repository.</span></span>

2. <span data-ttu-id="741d7-127">プロジェクト フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="741d7-127">Navigate to the project folder:</span></span>

    ```bash
    cd stock-ticker
    ```

3. <span data-ttu-id="741d7-128">ローカル Web サーバーを起動します。</span><span class="sxs-lookup"><span data-stu-id="741d7-128">Start the local web server.</span></span>

    * <span data-ttu-id="741d7-129">Windows 版 Excel のテストに使用する場合、ローカルの Web サーバーを起動するのには次のコマンドを実行、カスタム関数は Excel、およびアドインが Sideload を起動します。</span><span class="sxs-lookup"><span data-stu-id="741d7-129">If you'll be using Excel for Windows to test your custom functions, run the following command to start the local web server, launch Excel, and sideload the add-in:</span></span>

        ```bash
        npm start
        ```

    * <span data-ttu-id="741d7-130">Excel Online を使用してカスタム関数をテストする場合は、次のコマンドを実行してローカル Web サーバーを起動します。</span><span class="sxs-lookup"><span data-stu-id="741d7-130">If you'll be using Excel Online to test your custom functions, run the following command to start the local web server:</span></span> 

        ```bash
        npm run start-web
        ```

## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="741d7-131">作成済みのカスタム関数を試してみてください。</span><span class="sxs-lookup"><span data-stu-id="741d7-131">Try out a prebuilt custom function</span></span>

<span data-ttu-id="741d7-p104">Yo Office ジェネレーターを使用して作成したカスタム関数プロジェクトには、 **src/customfunction.js** ファイル内で定義されたいくつか作成済みのカスタム関数が含まれています。プロジェクトのルートディレクトリにある **manifest.xml** ファイルは、すべてのカスタム関数が `CONTOSO` 名前空間に属することを指定します。</span><span class="sxs-lookup"><span data-stu-id="741d7-p104">The custom functions project that you created by using the Yo Office generator contains some prebuilt custom functions, defined within the **src/customfunction.js** file. The **manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="741d7-p105">作成済みのカスタム関数のいずれかを使用する前に、Excel でカスタム関数アドインを登録する必要があります。このチュートリアルで使用するプラットフォームの手順を完了するようにします。</span><span class="sxs-lookup"><span data-stu-id="741d7-p105">Before you can use any of the prebuilt custom functions, you must register the custom functions add-in in Excel. Do so by completing steps for the platform that you'll be using in this tutorial.</span></span>

* <span data-ttu-id="741d7-136">カスタム関数をテストするには 、Windows 版 Excel を使用します。</span><span class="sxs-lookup"><span data-stu-id="741d7-136">If you'll be using Excel for Windows to test your custom functions:</span></span>

    1. <span data-ttu-id="741d7-137">Excel では、 **[挿入]** タブを選択し、 **[個人用アドイン]** の右にある下向き矢印を選択します。![[個人用アドイン]の矢印が強調表示された状態で Windows 版 Excel にリボンを挿入します。](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="741d7-137">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

    2. <span data-ttu-id="741d7-p106">利用可能なアドインの一覧で、  **アドインの開発者** のセクションを検索し、\*\* Excel カスタム関数\*\*   アドインを選択して登録します。 ![Excel カスタム関数アドインを[個人用アドイン] リストで強調表示して Windows 版 Excel にリボンを挿入します。](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="741d7-p106">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.  ![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

* <span data-ttu-id="741d7-140">カスタム関数をテストするには 、Excel Online を使用します。</span><span class="sxs-lookup"><span data-stu-id="741d7-140">If you'll be using Excel Online to test your custom functions:</span></span> 

    1. <span data-ttu-id="741d7-141">Excel Online で\*\* [挿入]\*\* ]タブを選択し、\*\* [アドイン]\*\* を選択します。![  [個人用アドイン]アイコンを強調表示して Excel Online でリボンを挿入します。](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="741d7-141">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

    2. <span data-ttu-id="741d7-142">\*\* [個人用アドインの管理] \*\* を選択し、 \*\* [個人用アドインのアップロード] \*\*を選択します。</span><span class="sxs-lookup"><span data-stu-id="741d7-142">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

    3. <span data-ttu-id="741d7-143">\*\* [参照... ]\*\*  を選択し、Yo Officeジェネレータが作成したプロジェクトのルートディレクトリに移動します。</span><span class="sxs-lookup"><span data-stu-id="741d7-143">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

    4. <span data-ttu-id="741d7-144">\*\* Manifest.xml**ファイルを選択して \*\* 開く** を選択し、\*\* アップロード\*\* を選択します。</span><span class="sxs-lookup"><span data-stu-id="741d7-144">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

<span data-ttu-id="741d7-p107">この時点で、プロジェクトの作成済みのカスタム関数がロードされて Excel 内で使用できます。Excel で次の手順を実行して`ADD` カスタム関数を試してみてください。</span><span class="sxs-lookup"><span data-stu-id="741d7-p107">At this point, the prebuilt custom functions in your project are loaded and available within Excel. Try out the `ADD` custom function by completing the following steps in Excel:</span></span>

1. <span data-ttu-id="741d7-p108">セル内には、**=CONTOSO** を入力します。オートコンプリートメニューには、 `CONTOSO` 名前空間内のすべての関数の一覧が表示されます。</span><span class="sxs-lookup"><span data-stu-id="741d7-p108">Within a cell, type **=CONTOSO**. Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="741d7-149">セルに次の値を指定し、Enter キーを押して、`CONTOSO.ADD` 数値 `10` および`200` 入力パラメータとして関数を実行します。</span><span class="sxs-lookup"><span data-stu-id="741d7-149">Run the `CONTOSO.ADD` function, with numbers `10` and `200` as input parameters, by specifying the following value in the cell and pressing enter:</span></span>

    ```
    =CONTOSO.ADD(10,200)
    ```

<span data-ttu-id="741d7-p109"> `ADD` カスタム関数は、入力パラメータとして指定されている 2 つの数値の合計を計算します。 `=CONTOSO.ADD(10,200)` を入力すると、Enter キーを押した後にセル内に *\*210** という結果が表示されます。</span><span class="sxs-lookup"><span data-stu-id="741d7-p109">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters. Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a><span data-ttu-id="741d7-152">Web サイトからデータを要求するカスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="741d7-152">Create a custom function that requests data from the web</span></span>

<span data-ttu-id="741d7-p110">API からの在庫の価格を要求し、その結果をワークシートのセルに表示する機能が必要な場合はどうなりますか？カスタム関数は、Web サイトから非同期にデータを簡単に要求できるように設計されています。</span><span class="sxs-lookup"><span data-stu-id="741d7-p110">What if you needed a function that could request the price of a stock from an API and display the result in the cell of a worksheet? Custom functions are designed so that you can easily request data from the web asynchronously.</span></span>

<span data-ttu-id="741d7-p111"> `stockPrice` というカスタム関数を作成し、株価表示（例：*\*MSFT** ）を受け取り、その株式の価格を返す次の手順を実行します。このカスタム関数は、無料で認証を必要としない IEX 取引 API を使用します。</span><span class="sxs-lookup"><span data-stu-id="741d7-p111">Complete the following steps to create a custom function named `stockPrice` that accepts a stock ticker (e.g., **MSFT**) and returns the price of that stock. This custom function uses the IEX Trading API, which is free and does not require authentication.</span></span>

1. <span data-ttu-id="741d7-157">Yo Office ジェネレーターが作成した **株価表示** プロジェクトでは、ファイル **src/customfunctions.js** を検索し、コード エディターで開きます。</span><span class="sxs-lookup"><span data-stu-id="741d7-157">In the **stock-ticker** project that the Yo Office generator created, find the file **src/customfunctions.js** and open it in your code editor.</span></span>

2. <span data-ttu-id="741d7-158">次のコードを **customfunctions.js** に追加して、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="741d7-158">Add the following code to **home.js** and save the file.</span></span>

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

3. <span data-ttu-id="741d7-p112">Excel でエンドユーザーがこの新しい機能を利用できるようにする前に、この機能を説明するメタデータを指定する必要があります。Yo Office ジェネレーターを作成した**株価表示** プロジェクトでは、ファイル **config/customfunctions.json** を検索し、コード エディターで開きます。  `functions` 次のオブジェクトを  **config/customfunctions.json**  ファイル内の配列に追加し、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="741d7-p112">Before Excel can make this new function available to end-users, you must specify metadata that describes this function. In the **stock-ticker** project that the Yo Office generator created, find the file **config/customfunctions.json** and open it in your code editor. Add the following object to the `functions` array within the **config/customfunctions.json** file and save the file.</span></span>

    <span data-ttu-id="741d7-162">この JSON は`stockPrice`関数を記述します。</span><span class="sxs-lookup"><span data-stu-id="741d7-162">This JSON describes the `stockPrice` function.</span></span>

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

4. <span data-ttu-id="741d7-p113">エンドユーザーが新しい機能を利用できるようにするには、アドインを Excel に再登録する必要があります。このチュートリアルで使用しているプラットフォームの次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="741d7-p113">You must reregister the add-in in Excel in order for the new function to be available to end-users. Complete the following steps for the platform that you're using in this tutorial.</span></span>

    * <span data-ttu-id="741d7-165">Windows 版 Excel の場合</span><span class="sxs-lookup"><span data-stu-id="741d7-165">If you're using Excel for Windows:</span></span>

        1. <span data-ttu-id="741d7-166">Excel を終了し、再度Excel を開きます。</span><span class="sxs-lookup"><span data-stu-id="741d7-166">Close Excel and then reopen Excel.</span></span>

        2. <span data-ttu-id="741d7-167">Excel では、 **[挿入]** タブを選択し、 **[個人用アドイン]** の右にある下向き矢印を選択します。![[個人用アドイン]の矢印が強調表示された状態で Windows 版 Excel にリボンを挿入します。](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="741d7-167">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

        1. <span data-ttu-id="741d7-p114">利用可能なアドインの一覧で、  **アドインの開発者** のセクションを検索し、\*\* Excel カスタム関数\*\*   アドインを選択して登録します。 ![Excel カスタム関数アドインを[個人用アドイン] リストで強調表示して Windows 版 Excel にリボンを挿入します。](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="741d7-p114">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.  ![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

    * <span data-ttu-id="741d7-170">Excel Online を使用している場合</span><span class="sxs-lookup"><span data-stu-id="741d7-170">If you're using Excel Online:</span></span> 

        1. <span data-ttu-id="741d7-171">Excel Online で\*\* [挿入]\*\* ]タブを選択し、\*\* [アドイン]\*\* を選択します。![  [個人用アドイン]アイコンを強調表示して Excel Online でリボンを挿入します。](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="741d7-171">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

        2. <span data-ttu-id="741d7-172">\*\* [個人用アドインの管理] \*\* を選択し、 \*\* [個人用アドインのアップロード] \*\*を選択します。</span><span class="sxs-lookup"><span data-stu-id="741d7-172">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

        3. <span data-ttu-id="741d7-173">\*\* [参照... ]\*\*  を選択し、Yo Officeジェネレータが作成したプロジェクトのルートディレクトリに移動します。</span><span class="sxs-lookup"><span data-stu-id="741d7-173">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

        4. <span data-ttu-id="741d7-174">\*\* Manifest.xml**ファイルを選択して \*\* 開く** を選択し、\*\* アップロード\*\* を選択します。</span><span class="sxs-lookup"><span data-stu-id="741d7-174">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

5. <span data-ttu-id="741d7-p115">ここで、新しい機能を試してみましょう。セル **B1**テキストを入力 `=CONTOSO.STOCKPRICE("MSFT")` し、Enter キーを押します。セル **B1** の結果は、Microsoft 株式の 1 株当たりの現在の株価であることがわかります。</span><span class="sxs-lookup"><span data-stu-id="741d7-p115">Now, let's try out the new function. In cell **B1**, type the text `=CONTOSO.STOCKPRICE("MSFT")` and press enter. You should see that the result in cell **B1** is the current stock price for one share of Microsoft stock.</span></span>

## <a name="create-a-streaming-asynchronous-custom-function"></a><span data-ttu-id="741d7-178">ストリーミング非同期のカスタム関数を作成します</span><span class="sxs-lookup"><span data-stu-id="741d7-178">Create a streaming asynchronous custom function</span></span>

<span data-ttu-id="741d7-p116"> `stockPrice` 関数を作成した時点で株式の価格を返しますが、株価は常に変化しています。株価のリアルタイム更新を取得するために、API からデータをストリーミングするカスタム関数を作成しましょう。</span><span class="sxs-lookup"><span data-stu-id="741d7-p116">The `stockPrice` function that you just created returns the price of a stock at a specific moment in time, but stock prices are always changing. Let's create a custom function that streams data from an API to get real-time updates on a stock price.</span></span>

<span data-ttu-id="741d7-p117">`stockPriceStream`という名前のカスタム関数を作成するには、次の手順を実行して、1000 ミリ秒ごとに指定した在庫の価格を要求します（前の要求が完了している場合）。最初の要求が進行中に、関数が呼び出されているセルのプレースホルダ値 **#GETTING_DATA** が表示される場合があります。関数によって値が返されると、**#GETTING_DATA**  はセル内の値に置き換えられます。</span><span class="sxs-lookup"><span data-stu-id="741d7-p117">Complete the following steps to create a custom function named `stockPriceStream` that requests the price of the specified stock every 1000 milliseconds (provided that the previous request has completed). While the initial request is in-progress, you may see the placeholder value **#GETTING_DATA** the cell where the function is being called. When a value is returned by the function, **#GETTING_DATA** will be replaced by that value in the cell.</span></span>

1. <span data-ttu-id="741d7-184">Yo Office ジェネレーターが作成した **株価表示** プロジェクトでは、 **src/customfunctions.js** に次のコードを追加し、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="741d7-184">In the **stock-ticker** project that the Yo Office generator created, add the following code to **src/customfunctions.js** and save the file.</span></span>

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

2. <span data-ttu-id="741d7-p118">Excel がこの新しい関数を使用できるようにエンドユーザーにする前にこの関数を記述するメタデータを指定してください。Yo Office ジェネレーターを作成した **株価表示** プロジェクトに `functions`  次のオブジェクトを \*\* config/customfunctions.json\*\*  ファイル内の配列に追加し、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="741d7-p118">Before Excel can make this new function available to end-users, you must specify metadata that describes this function. In the **stock-ticker** project that the Yo Office generator created, add the following object to the `functions` array within the **config/customfunctions.json** file and save the file.</span></span>

    <span data-ttu-id="741d7-p119">この JSON は`stockPriceStream`関数を記述します。ストリーミング関数の場合は、 `stream` プロパティと `cancelable`  プロパティは、`true` 次のコード例に示すように `options` オブジェクト 内 に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="741d7-p119">This JSON describes the `stockPriceStream` function. For any streaming function, the `stream` property and the `cancelable` property must be set to `true` within the `options` object, as shown in this code sample.</span></span>

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

3. <span data-ttu-id="741d7-p120">エンドユーザーが新しい機能を利用できるようにするには、アドインを Excel に再登録する必要があります。このチュートリアルで使用しているプラットフォームの次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="741d7-p120">You must reregister the add-in in Excel in order for the new function to be available to end-users. Complete the following steps for the platform that you're using in this tutorial.</span></span>

    * <span data-ttu-id="741d7-191">Windows 版 Excel の場合</span><span class="sxs-lookup"><span data-stu-id="741d7-191">If you're using Excel for Windows:</span></span>

        1. <span data-ttu-id="741d7-192">Excel を終了し、再度Excel を開きます。</span><span class="sxs-lookup"><span data-stu-id="741d7-192">Close Excel and then reopen Excel.</span></span>
        
        2. <span data-ttu-id="741d7-193">Excel では、 **[挿入]** タブを選択し、 **[個人用アドイン]** の右にある下向き矢印を選択します。![[個人用アドイン]の矢印が強調表示された状態で Windows 版 Excel にリボンを挿入します。](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="741d7-193">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

        3. <span data-ttu-id="741d7-p121">利用可能なアドインの一覧で、  **アドインの開発者** のセクションを検索し、\*\* Excel カスタム関数\*\*   アドインを選択して登録します。 ![Excel カスタム関数アドインを[個人用アドイン] リストで強調表示して Windows 版 Excel にリボンを挿入します。](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="741d7-p121">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.  ![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

    * <span data-ttu-id="741d7-196">Excel Online を使用している場合</span><span class="sxs-lookup"><span data-stu-id="741d7-196">If you're using Excel Online:</span></span> 

        1. <span data-ttu-id="741d7-197">Excel Online で\*\* [挿入]\*\* ]タブを選択し、\*\* [アドイン]\*\* を選択します。![  [個人用アドイン]アイコンを強調表示して Excel Online でリボンを挿入します。](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="741d7-197">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

        2. <span data-ttu-id="741d7-198">\*\* [個人用アドインの管理] \*\* を選択し、 \*\* [個人用アドインのアップロード] \*\*を選択します。</span><span class="sxs-lookup"><span data-stu-id="741d7-198">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

        3. <span data-ttu-id="741d7-199">\*\* [参照... ]\*\*  を選択し、Yo Officeジェネレータが作成したプロジェクトのルートディレクトリに移動します。</span><span class="sxs-lookup"><span data-stu-id="741d7-199">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

        4. <span data-ttu-id="741d7-200">\*\* Manifest.xml**ファイルを選択して \*\* 開く** を選択し、\*\* アップロード\*\* を選択します。</span><span class="sxs-lookup"><span data-stu-id="741d7-200">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

4. <span data-ttu-id="741d7-p122">ここで、新しい機能を試してみましょう。セル**C1** テキストを入力 し、`=CONTOSO.STOCKPRICESTREAM("MSFT")`Enter キーを押します。株式市場が開いている Microsoft の株 1 株のリアルタイムの価格を反映するようにセル **C1** の結果が常に更新されているはずです。</span><span class="sxs-lookup"><span data-stu-id="741d7-p122">Now, let's try out the new function. In cell **C1**, type the text `=CONTOSO.STOCKPRICESTREAM("MSFT")` and press enter. Provided that the stock market is open, you should see that the result in cell **C1** is constantly updated to reflect the real-time price for one share of Microsoft stock.</span></span>

## <a name="next-steps"></a><span data-ttu-id="741d7-204">次の手順</span><span class="sxs-lookup"><span data-stu-id="741d7-204">Next steps</span></span>

<span data-ttu-id="741d7-p123">このチュートリアルでは、新しいカスタム関数プロジェクトを作成し、事前ビルドされた関数を試し、Web からデータを要求するカスタム関数を作成し、Web からリアルタイムデータをストリームするカスタム関数を作成しました。Excel のカスタム関数の詳細については、次の記事を参照してください。</span><span class="sxs-lookup"><span data-stu-id="741d7-p123">In this tutorial, you've created a new custom functions project, tried out a prebuilt function, created a custom function that requests data from the web, and created a custom function that streams real-time data from the web. To learn more about custom functions in Excel, continue to the following article:</span></span> 

> [!div class="nextstepaction"]
> [<span data-ttu-id="741d7-207">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="741d7-207">Create custom functions in Excel (Preview)</span></span>](../excel/custom-functions-overview.md)

## <a name="legal-information"></a><span data-ttu-id="741d7-208">法的情報</span><span class="sxs-lookup"><span data-stu-id="741d7-208">Legal Information</span></span>

<span data-ttu-id="741d7-p124"> [IEX](https://iextrading.com/developer/)が無料で提供されるデータです。 [IEX の利用規約](https://iextrading.com/api-exhibit-a/)を表示します。このチュートリアルで Microsoft が IEX API を使用するのは、教育目的でのみです。</span><span class="sxs-lookup"><span data-stu-id="741d7-p124">Data provided free by [IEX](https://iextrading.com/developer/). View [IEX's Terms of Use](https://iextrading.com/api-exhibit-a/). Microsoft's use of the IEX API in this tutorial is for educational purposes only.</span></span>
