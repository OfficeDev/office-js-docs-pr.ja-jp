---
title: Excel カスタム関数のチュートリアル
description: このチュートリアルでは、計算の実行、Web データの要求、Web データのストリームが可能なカスタム関数を含む Excel アドインを作成します。
ms.date: 06/27/2019
ms.prod: excel
ms.topic: tutorial
localization_priority: Normal
ms.openlocfilehash: 1aa05581d1b0dfb1f5affa019e51b84126c8d199
ms.sourcegitcommit: 90c2d8236c6b30d80ac2b13950028a208ef60973
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/02/2019
ms.locfileid: "35454731"
---
# <a name="tutorial-create-custom-functions-in-excel"></a><span data-ttu-id="0e84d-103">チュートリアル: Excel でのカスタム関数の作成</span><span class="sxs-lookup"><span data-stu-id="0e84d-103">Tutorial: Create custom functions in Excel</span></span>

<span data-ttu-id="0e84d-104">カスタム関数では、関数をアドインの一部として JavaScript で定義することによって、Excel に新しい関数を追加できます。</span><span class="sxs-lookup"><span data-stu-id="0e84d-104">Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="0e84d-105">ユーザーは、Excel 内から、`SUM()` などの Excel のあらゆるネイティブ関数の場合と同じようにカスタム関数にアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="0e84d-105">Users within Excel can access custom functions as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="0e84d-106">計算のような単純なタスク、または Web からワークシートへのデータのリアルタイム ストリーミングのようなより複雑なタスクを実行するカスタム関数を作成できます。</span><span class="sxs-lookup"><span data-stu-id="0e84d-106">You can create custom functions that perform simple tasks like calculations or more complex tasks such as streaming real-time data from the web into a worksheet.</span></span>

<span data-ttu-id="0e84d-107">このチュートリアルの内容:</span><span class="sxs-lookup"><span data-stu-id="0e84d-107">In this tutorial, you will:</span></span>
> [!div class="checklist"]
> * <span data-ttu-id="0e84d-108">[Office アドイン用の Yeoman ジェネレーター](https://www.npmjs.com/package/generator-office)を使用して、カスタム関数アドインを作成します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-108">Create a custom function add-in using the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office).</span></span> 
> * <span data-ttu-id="0e84d-109">あらかじめ用意されているカスタム関数を使用し、単純な計算を実行します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-109">Use a prebuilt custom function to perform a simple calculation.</span></span>
> * <span data-ttu-id="0e84d-110">Web からデータを取得するカスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-110">Create a custom function that gets data from the web.</span></span>
> * <span data-ttu-id="0e84d-111">Web からデータをリアルタイムでストリーミングするカスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-111">Create a custom function that streams real-time data from the web.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="0e84d-112">前提条件</span><span class="sxs-lookup"><span data-stu-id="0e84d-112">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

* <span data-ttu-id="0e84d-113">Windows 上の Excel (バージョン1904以降、Office 365 サブスクリプションに接続されている) または web</span><span class="sxs-lookup"><span data-stu-id="0e84d-113">Excel on Windows (version 1904 or later, connected to Office 365 subscription) or on the web</span></span>

## <a name="create-a-custom-functions-project"></a><span data-ttu-id="0e84d-114">カスタム関数プロジェクトを作成する</span><span class="sxs-lookup"><span data-stu-id="0e84d-114">Create a custom functions project</span></span>

 <span data-ttu-id="0e84d-115">まず、カスタム関数アドインをビルドするコード プロジェクトを作成します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-115">To start, you'll create the code project to build your custom function add-in.</span></span> <span data-ttu-id="0e84d-116">[Office アドイン用の [ごみ箱] ジェネレーター](https://www.npmjs.com/package/generator-office)では、プロジェクトに事前に用意されているカスタム関数を使用してセットアップし、試すことができます。カスタム関数のクイックスタートを既に実行してプロジェクトを生成した場合は、そのプロジェクトを引き続き使用して、[この手順](#create-a-custom-function-that-requests-data-from-the-web)に進んでください。</span><span class="sxs-lookup"><span data-stu-id="0e84d-116">The [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) will set up your project with some prebuilt custom functions that you can try out. If you have already run the custom functions quick start and generated a project, continue to use that project and skip to [this step](#create-a-custom-function-that-requests-data-from-the-web) instead.</span></span>

1. <span data-ttu-id="0e84d-117">次のコマンドを実行し、以下のようにプロンプトに応答します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-117">Run the following command and then answer the prompts as follows.</span></span>
    
    ```command&nbsp;line
    yo office
    ```
    
    * <span data-ttu-id="0e84d-118">**Choose a project type: (プロジェクトの種類を選択)** `Excel Custom Functions Add-in project`</span><span class="sxs-lookup"><span data-stu-id="0e84d-118">**Choose a project type:** `Excel Custom Functions Add-in project`</span></span>
    * <span data-ttu-id="0e84d-119">**Choose a script type: (スクリプトの種類を選択)** `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="0e84d-119">**Choose a script type:** `JavaScript`</span></span>
    * <span data-ttu-id="0e84d-120">**What would you want to name your add-in?: (アドインの名前を何にしますか)**</span><span class="sxs-lookup"><span data-stu-id="0e84d-120">**What do you want to name your add-in?**</span></span> `starcount`

    ![カスタム関数の Office アドイン用の Yeoman ジェネレーターのプロンプト](../images/starcountPrompt.png)
    
    <span data-ttu-id="0e84d-122">Yeoman ジェネレーターはプロジェクト ファイルを作成し、サポートしているノード コンポーネントをインストールします。</span><span class="sxs-lookup"><span data-stu-id="0e84d-122">The Yeoman generator will create the project files and install supporting Node components.</span></span>

2. <span data-ttu-id="0e84d-123">プロジェクトのルート フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-123">Navigate to the root folder of the project.</span></span>
    
    ```command&nbsp;line
    cd starcount
    ```

3. <span data-ttu-id="0e84d-124">プロジェクトをビルドします。</span><span class="sxs-lookup"><span data-stu-id="0e84d-124">Build the project.</span></span>
    
    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > <span data-ttu-id="0e84d-125">開発の最中でも、OfficeアドインはHTTPではなくHTTPSを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0e84d-125">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="0e84d-126">`npm run build`の実行後に証明書をインストールするように指示が出された場合は、Yeomanジェネレーターが提供する証明書をインストールする手順に従ってください。</span><span class="sxs-lookup"><span data-stu-id="0e84d-126">If you are prompted to install a certificate after you run `npm run build`, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

4. <span data-ttu-id="0e84d-127">Node.js で実行しているローカル Web サーバーを開始します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-127">Start the local web server, which runs in Node.js.</span></span> <span data-ttu-id="0e84d-128">Web または Windows 上の Excel でカスタム関数アドインを試すことができます。</span><span class="sxs-lookup"><span data-stu-id="0e84d-128">You can try out the custom function add-in in Excel on the web or Windows.</span></span>

# <a name="excel-on-windowstabexcel-windows"></a>[<span data-ttu-id="0e84d-129">Windows 上の Excel</span><span class="sxs-lookup"><span data-stu-id="0e84d-129">Excel on Windows</span></span>](#tab/excel-windows)

<span data-ttu-id="0e84d-130">Windows の Excel でアドインをテストするには、次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-130">To test your add-in in Excel on Windows, run the following command.</span></span> <span data-ttu-id="0e84d-131">このコマンドを実行すると、ローカル web サーバーが起動し、アドインが読み込まれた状態で Excel が開きます。</span><span class="sxs-lookup"><span data-stu-id="0e84d-131">When you run this command, the local web server will start and Excel will open with your add-in loaded.</span></span>

```command&nbsp;line
npm run start:desktop
```

# <a name="excel-on-the-webtabexcel-online"></a>[<span data-ttu-id="0e84d-132">Excel on the web</span><span class="sxs-lookup"><span data-stu-id="0e84d-132">Excel on the web</span></span>](#tab/excel-online)

<span data-ttu-id="0e84d-133">ブラウザー上の Excel でアドインをテストするには、次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-133">To test your add-in in Excel on a browser, run the following command.</span></span> <span data-ttu-id="0e84d-134">このコマンドを実行すると、ローカル Web サーバーが起動します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-134">When you run this command, the local web server will start.</span></span>

```command&nbsp;line
npm run start:web
```

<span data-ttu-id="0e84d-135">カスタム関数アドインを使用するには、web 上の Excel で新しいブックを開きます。</span><span class="sxs-lookup"><span data-stu-id="0e84d-135">To use your custom functions add-in, open a new workbook in Excel on the web.</span></span> <span data-ttu-id="0e84d-136">このブックでは、次の手順を実行して、アドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="0e84d-136">In this workbook, complete the following steps to sideload your add-in.</span></span>

1. <span data-ttu-id="0e84d-137">Excel で、[**挿入**] タブを選択し、[**アドイン**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-137">In Excel, choose the **Insert** tab and then choose **Add-ins**.</span></span>

   ![[個人用アドイン] アイコンが強調表示されている web 上の Excel にリボンを挿入する](../images/excel-cf-online-register-add-in-1.png)
   
2. <span data-ttu-id="0e84d-139">**[マイ アドインの管理]** を選択し、**[マイ アドインのアップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-139">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="0e84d-140">**[参照...]** を選択し、Yeoman ジェネレーターによって作成されたプロジェクトのルート ディレクトリに移動します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-140">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="0e84d-141">**manifest.xml** ファイルを選択し、**[開く]** を選択し、**[アップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-141">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

--- 
    
## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="0e84d-142">あらかじめ用意されているカスタム関数を試す</span><span class="sxs-lookup"><span data-stu-id="0e84d-142">Try out a prebuilt custom function</span></span>

<span data-ttu-id="0e84d-143">作成したカスタム関数プロジェクトには、 **/src/functions/functions.js**ファイル内で定義されたあらかじめ用意されたカスタム関数がいくつか含まれています。</span><span class="sxs-lookup"><span data-stu-id="0e84d-143">The custom functions project that you created contains some prebuilt custom functions, defined within the **./src/functions/functions.js** file.</span></span> <span data-ttu-id="0e84d-144">**./manifest.xml** ファイルによって、カスタム関数はすべて `CONTOSO` 名前空間に属することが指定されます。</span><span class="sxs-lookup"><span data-stu-id="0e84d-144">The **./manifest.xml** file specifies that all custom functions belong to the `CONTOSO` namespace.</span></span> <span data-ttu-id="0e84d-145">Excel でカスタム関数にアクセスするには、CONTOSO 名前空間を使用します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-145">You'll use the CONTOSO namespace to access the custom functions in Excel.</span></span>

<span data-ttu-id="0e84d-146">その後、次の手順を実行し、`ADD` カスタム関数を試します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-146">Next you'll try out the `ADD` custom function by completing the following steps:</span></span>

1. <span data-ttu-id="0e84d-147">Excel で、任意のセルに移動し、`=CONTOSO` と入力します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-147">In Excel, go to any cell and enter `=CONTOSO`.</span></span> <span data-ttu-id="0e84d-148">`CONTOSO` 名前空間にあるすべての関数がオートコンプリート メニューに一覧表示されます。</span><span class="sxs-lookup"><span data-stu-id="0e84d-148">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="0e84d-149">セル内で値 `=CONTOSO.ADD(10,200)` を入力して Enter キーを押し、入力パラメーターとして数値 `10` と `200` を指定して、`CONTOSO.ADD` 関数を実行します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-149">Run the `CONTOSO.ADD` function, with numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="0e84d-150">`ADD` カスタム関数によって、指定した 2 つの数字の合計が計算され、**210** という結果が返されます。</span><span class="sxs-lookup"><span data-stu-id="0e84d-150">The `ADD` custom function computes the sum of the two numbers that you provided and returns the result of **210**.</span></span>

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a><span data-ttu-id="0e84d-151">Web からデータを要求するカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="0e84d-151">Create a custom function that requests data from the web</span></span>

<span data-ttu-id="0e84d-152">Web からデータを統合することは、カスタム関数を使用して Excel を拡張する優れた方法です。</span><span class="sxs-lookup"><span data-stu-id="0e84d-152">Integrating data from the Web is a great way to extend Excel through custom functions.</span></span> <span data-ttu-id="0e84d-153">次に、指定された Github `getStarCount`リポジトリのスター数を示すという名前のカスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-153">Next you’ll create a custom function named `getStarCount` that shows how many stars a given Github repository possesses.</span></span>

1. <span data-ttu-id="0e84d-154">**Starcount**プロジェクトで、 **/src/functions/functions.js**を見つけて、コードエディターで開きます。</span><span class="sxs-lookup"><span data-stu-id="0e84d-154">In the **starcount** project, find the file **./src/functions/functions.js** and open it in your code editor.</span></span> 

2. <span data-ttu-id="0e84d-155">**関数 .js**で、次のコードを追加します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-155">In **function.js**, add the following code:</span></span> 

```JS
 /**
   * Gets the star count for a given Github repository.
   * @customfunction 
   * @param {string} userName string name of Github user or organization.
   * @param {string} repoName string name of the Github repository.
   * @return {number} number of stars given to a Github repository.
   */
    async function getStarCount(userName, repoName) {
      try {
        //You can change this URL to any web request you want to work with.
        const url = "https://api.github.com/repos/" + userName + "/" + repoName;
        const response = await fetch(url);
        //Expect that status code is in 200-299 range
        if (!response.ok) {
          throw new Error(response.statusText)
        }
          const jsonResponse = await response.json();
          return jsonResponse.watchers_count;
      }
      catch (error) {
        return error;
      }
      }
    CustomFunctions.associate("GETSTARCOUNT", getStarCount);
```

<span data-ttu-id="0e84d-156">`CustomFunctions.associate` コードは、JavaScript で関数の `id` と `getStarCount` の関数アドレスを関連付けて、Excel により関数を呼び出せるようにします。</span><span class="sxs-lookup"><span data-stu-id="0e84d-156">The `CustomFunctions.associate` code associates the `id` of the function with the function address of `getStarCount` in JavaScript so that Excel can call your function.</span></span>

3. <span data-ttu-id="0e84d-157">次のコマンドを実行してプロジェクトを再構築します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-157">Run the following command to rebuild the project.</span></span>

    ```command&nbsp;line
    npm run build
    ```

4. <span data-ttu-id="0e84d-158">Excel on the web または Windows の場合は、次の手順を実行して、Excel でアドインを再登録します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-158">Complete the following steps (for either Excel on the web or Windows) to re-register the add-in in Excel.</span></span> <span data-ttu-id="0e84d-159">新しい関数を使用できるようにするには、これらの手順を完了する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0e84d-159">You must complete these steps before the new function will be available.</span></span>

### <a name="excel-on-windowstabexcel-windows"></a>[<span data-ttu-id="0e84d-160">Windows 上の Excel</span><span class="sxs-lookup"><span data-stu-id="0e84d-160">Excel on Windows</span></span>](#tab/excel-windows)

1. <span data-ttu-id="0e84d-161">Excel を閉じて再び開きます。</span><span class="sxs-lookup"><span data-stu-id="0e84d-161">Close Excel and then reopen Excel.</span></span>

2. <span data-ttu-id="0e84d-162">Excel で [**挿入**] タブを選択し、[**マイ**アドイン] の右側にある下向き矢印を選択します。 ![[個人用アドイン] 矢印が強調表示されている Windows 上の Excel でのリボンの挿入](../images/select-insert.png)</span><span class="sxs-lookup"><span data-stu-id="0e84d-162">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel on Windows with the My Add-ins arrow highlighted](../images/select-insert.png)</span></span>

3. <span data-ttu-id="0e84d-163">利用可能なアドインの一覧で、[**開発者用アドイン**] セクションを見つけ、 **starcount**アドインを選択して登録します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-163">In the list of available add-ins, find the **Developer Add-ins** section and select the **starcount** add-in to register it.</span></span>
    <span data-ttu-id="0e84d-164">![[個人用アドイン] ボックスの一覧で強調表示された Excel カスタム関数アドインを使用して、Excel の Excel にリボンを挿入する](../images/list-starcount.png)</span><span class="sxs-lookup"><span data-stu-id="0e84d-164">![Insert ribbon in Excel on Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/list-starcount.png)</span></span>


# <a name="excel-on-the-webtabexcel-online"></a>[<span data-ttu-id="0e84d-165">Excel on the web</span><span class="sxs-lookup"><span data-stu-id="0e84d-165">Excel on the web</span></span>](#tab/excel-online)

1. <span data-ttu-id="0e84d-166">Excel で、[**挿入**] タブを選択し、[**アドイン**] を選択します。 ![[個人用アドイン] アイコンが強調表示されている web 上の Excel にリボンを挿入する](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="0e84d-166">In Excel, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel on the web with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

2. <span data-ttu-id="0e84d-167">**[マイ アドインの管理]** を選択し、**[マイ アドインのアップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-167">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="0e84d-168">**[参照...]** を選択し、Yeoman ジェネレーターによって作成されたプロジェクトのルート ディレクトリに移動します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-168">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="0e84d-169">**manifest.xml** ファイルを選択し、**[開く]** を選択し、**[アップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-169">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

---

<ol start="5">
<li> <span data-ttu-id="0e84d-170">新しい関数をお試しください。</span><span class="sxs-lookup"><span data-stu-id="0e84d-170">Try out the new function.</span></span> <span data-ttu-id="0e84d-171">セル<strong>B1</strong>に、「CONTOSO」というテキストを入力し<strong>ます。GETSTARCOUNT ("OfficeDev", "Excel-ユーザー定義関数")</strong> 。 enter キーを押します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-171">In cell <strong>B1</strong>, type the text <strong>=CONTOSO.GETSTARCOUNT("OfficeDev", "Excel-Custom-Functions")</strong> and press enter.</span></span> <span data-ttu-id="0e84d-172">セル<strong>B1</strong>の結果は、 [Excel のカスタム機能である Github リポジトリ](https://github.com/OfficeDev/Excel-Custom-Functions)に与えられている現在の星数であることがわかります。</span><span class="sxs-lookup"><span data-stu-id="0e84d-172">You should see that the result in cell <strong>B1</strong> is the current number of stars given to the [Excel-Custom-Functions Github repository](https://github.com/OfficeDev/Excel-Custom-Functions).</span></span></li>
</ol>

## <a name="create-a-streaming-asynchronous-custom-function"></a><span data-ttu-id="0e84d-173">非同期でデータをストリーミングするカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="0e84d-173">Create a streaming asynchronous custom function</span></span>

<span data-ttu-id="0e84d-174">関数`getStarCount`は、特定の時点でリポジトリにある星の数を返します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-174">The `getStarCount` function returns the number of stars a repository has at a specific moment in time.</span></span> <span data-ttu-id="0e84d-175">カスタム関数は、絶えず変化するデータを返すこともできます。</span><span class="sxs-lookup"><span data-stu-id="0e84d-175">Custom functions can also return data that is continuously changing.</span></span> <span data-ttu-id="0e84d-176">これらの関数は、ストリーミング関数と呼ばれます。</span><span class="sxs-lookup"><span data-stu-id="0e84d-176">These functions are called streaming functions.</span></span> <span data-ttu-id="0e84d-177">これらには、 `invocation`関数が呼び出されたセルを参照するパラメーターを含める必要があります。</span><span class="sxs-lookup"><span data-stu-id="0e84d-177">They must include an `invocation` parameter which refers to the cell where the function was called from.</span></span> <span data-ttu-id="0e84d-178">`invocation`パラメーターは、セルの内容をいつでも更新するために使用されます。</span><span class="sxs-lookup"><span data-stu-id="0e84d-178">The `invocation` parameter is used to update the contents of the cell at any time.</span></span>  

<span data-ttu-id="0e84d-179">次のコードサンプルでは、 `currentTime`と`clock`の2つの関数があることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="0e84d-179">In the following code sample, you'll notice that there are two functions, `currentTime` and `clock`.</span></span> <span data-ttu-id="0e84d-180">関数`currentTime`は、ストリーミングを使用しない静的関数です。</span><span class="sxs-lookup"><span data-stu-id="0e84d-180">The `currentTime` function is a static function that does not use streaming.</span></span> <span data-ttu-id="0e84d-181">日付を文字列として返します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-181">It returns the date as a string.</span></span> <span data-ttu-id="0e84d-182">`clock`関数は、 `currentTime`関数を使用して、Excel のセルに対して2秒ごとに新しい時刻を提供します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-182">The `clock` function uses the `currentTime` function to provide the new time every second to a cell in Excel.</span></span> <span data-ttu-id="0e84d-183">を使用`invocation.setResult`して、Excel セルに時刻を提供`invocation.onCanceled`し、関数がキャンセルされたときに発生する処理を処理します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-183">It uses `invocation.setResult` to deliver the time to the Excel cell and `invocation.onCanceled` to handle what occurs when the function is canceled.</span></span>

1. <span data-ttu-id="0e84d-184">**Starcount**プロジェクトで、次のコードを **/src/functions/functions.js**に追加し、ファイルを保存します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-184">In the **starcount** project, add the following code to **./src/functions/functions.js** and save the file.</span></span>

```JS
/**
 * Returns the current time
 * @returns {string} String with the current time formatted for the current locale.
 */
function currentTime() {
  return new Date().toLocaleTimeString();
}

CustomFunctions.associate("CURRENTTIME", currentTime); 

 /**
 * Displays the current time once a second
 * @customfunction
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 */
function clock(invocation) {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}
CustomFunctions.associate("CLOCK", clock);
```

<span data-ttu-id="0e84d-185">`CustomFunctions.associate` コードは、JavaScript で関数の `id` と `CLOCK` の関数アドレスを関連付けて、Excel により関数を呼び出せるようにします。</span><span class="sxs-lookup"><span data-stu-id="0e84d-185">The `CustomFunctions.associate` code associates the `id` of the function with the function address of `CLOCK` in JavaScript so that Excel can call your function.</span></span>

2. <span data-ttu-id="0e84d-186">次のコマンドを実行してプロジェクトを再構築します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-186">Run the following command to rebuild the project.</span></span>

    ```command&nbsp;line
    npm run build
    ```

3. <span data-ttu-id="0e84d-187">Excel on the web または Windows の場合は、次の手順を実行して、Excel でアドインを再登録します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-187">Complete the following steps (for either Excel on the web or Windows) to re-register the add-in in Excel.</span></span> <span data-ttu-id="0e84d-188">新しい関数を使用できるようにするには、これらの手順を完了する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0e84d-188">You must complete these steps before the new function will be available.</span></span> 

# <a name="excel-on-windowstabexcel-windows"></a>[<span data-ttu-id="0e84d-189">Windows 上の Excel</span><span class="sxs-lookup"><span data-stu-id="0e84d-189">Excel on Windows</span></span>](#tab/excel-windows)

1. <span data-ttu-id="0e84d-190">Excel を閉じて再び開きます。</span><span class="sxs-lookup"><span data-stu-id="0e84d-190">Close Excel and then reopen Excel.</span></span>

2. <span data-ttu-id="0e84d-191">Excel で [**挿入**] タブを選択し、[**マイ**アドイン] の右側にある下向き矢印を選択します。 ![[個人用アドイン] 矢印が強調表示されている Windows 上の Excel でのリボンの挿入](../images/select-insert.png)</span><span class="sxs-lookup"><span data-stu-id="0e84d-191">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel on Windows with the My Add-ins arrow highlighted](../images/select-insert.png)</span></span>

3. <span data-ttu-id="0e84d-192">利用可能なアドインの一覧で、[**開発者用アドイン**] セクションを見つけ、 **starcount**アドインを選択して登録します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-192">In the list of available add-ins, find the **Developer Add-ins** section and select the **starcount** add-in to register it.</span></span>
    <span data-ttu-id="0e84d-193">![[個人用アドイン] ボックスの一覧で強調表示された Excel カスタム関数アドインを使用して、Excel の Excel にリボンを挿入する](../images/list-starcount.png)</span><span class="sxs-lookup"><span data-stu-id="0e84d-193">![Insert ribbon in Excel on Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/list-starcount.png)</span></span>

# <a name="excel-on-the-webtabexcel-online"></a>[<span data-ttu-id="0e84d-194">Excel on the web</span><span class="sxs-lookup"><span data-stu-id="0e84d-194">Excel on the web</span></span>](#tab/excel-online)

1. <span data-ttu-id="0e84d-195">Excel で、[**挿入**] タブを選択し、[**アドイン**] を選択します。 ![[個人用アドイン] アイコンが強調表示されている web 上の Excel にリボンを挿入する](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="0e84d-195">In Excel, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel on the web with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

2. <span data-ttu-id="0e84d-196">**[マイ アドインの管理]** を選択し、**[マイ アドインのアップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-196">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="0e84d-197">**[参照...]** を選択し、Yeoman ジェネレーターによって作成されたプロジェクトのルート ディレクトリに移動します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-197">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="0e84d-198">**manifest.xml** ファイルを選択し、**[開く]** を選択し、**[アップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-198">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

--- 

<ol start="4">
<li><span data-ttu-id="0e84d-199">新しい関数をお試しください。</span><span class="sxs-lookup"><span data-stu-id="0e84d-199">Try out the new function.</span></span> <span data-ttu-id="0e84d-200">セル<strong>C1</strong>に、「CONTOSO」というテキストを入力し<strong>ます。CLOCK ())</strong>と入力し、enter キーを押します。</span><span class="sxs-lookup"><span data-stu-id="0e84d-200">In cell <strong>C1</strong>, type the text <strong>=CONTOSO.CLOCK())</strong> and press enter.</span></span> <span data-ttu-id="0e84d-201">現在の日付が表示され、1秒ごとに更新が流れます。</span><span class="sxs-lookup"><span data-stu-id="0e84d-201">You should see the current date, which streams an update every second.</span></span> <span data-ttu-id="0e84d-202">このクロックはループのタイマーにすぎませんが、リアルタイムデータに対する web 要求を行う、より複雑な関数でタイマーを設定するのと同じ概念を使用できます。</span><span class="sxs-lookup"><span data-stu-id="0e84d-202">While this clock is just a timer on a loop, you can use the same idea of setting a timer on more complex functions that make web requests for real-time data.</span></span></li>
</ol>

## <a name="next-steps"></a><span data-ttu-id="0e84d-203">次のステップ</span><span class="sxs-lookup"><span data-stu-id="0e84d-203">Next steps</span></span>

<span data-ttu-id="0e84d-204">おめでとうございます。</span><span class="sxs-lookup"><span data-stu-id="0e84d-204">Congratulations!</span></span> <span data-ttu-id="0e84d-205">新しいカスタム関数プロジェクトを作成し、あらかじめ作成された関数を試し、web からデータを要求するカスタム関数を作成し、データをストリーム処理するカスタム関数を作成しました。</span><span class="sxs-lookup"><span data-stu-id="0e84d-205">You've created a new custom functions project, tried out a prebuilt function, created a custom function that requests data from the web, and created a custom function that streams data.</span></span> <span data-ttu-id="0e84d-206">この関数のデバッグは[、カスタム関数のデバッグ手順](../excel/custom-functions-debugging.md)を使用して実行することもできます。</span><span class="sxs-lookup"><span data-stu-id="0e84d-206">You can also try out debugging this function using [the custom function debugging instructions](../excel/custom-functions-debugging.md).</span></span> <span data-ttu-id="0e84d-207">Excel のカスタム関数に関する詳細については、次の記事にお進みください。</span><span class="sxs-lookup"><span data-stu-id="0e84d-207">To learn more about custom functions in Excel, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="0e84d-208">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="0e84d-208">Create custom functions in Excel</span></span>](../excel/custom-functions-overview.md)