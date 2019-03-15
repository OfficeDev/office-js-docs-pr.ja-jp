---
ms.date: 03/06/2019
description: Excel クイックスタートガイドでのカスタム関数の開発。
title: カスタム関数クイックスタート (プレビュー)
localization_priority: Normal
ms.openlocfilehash: 9dd3e5a99f08ce0b931e705fac3312ab10c19e18
ms.sourcegitcommit: 8fb60c3a31faedaea8b51b46238eb80c590a2491
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/14/2019
ms.locfileid: "30632703"
---
# <a name="get-started-developing-excel-custom-functions"></a><span data-ttu-id="79136-103">Excel カスタム関数の開発を始める</span><span class="sxs-lookup"><span data-stu-id="79136-103">Get started developing Excel custom functions</span></span>

<span data-ttu-id="79136-104">カスタム関数を使用すると、開発者は、JavaScript または Typescript でアドインの一部として定義することによって、Excel に新しい関数を追加できるようになります。</span><span class="sxs-lookup"><span data-stu-id="79136-104">With custom functions, developers can now add new functions to Excel by defining them in JavaScript or Typescript as part of an add-in.</span></span> <span data-ttu-id="79136-105">excel ユーザーは、excel の任意のネイティブ関数の場合と同じように、カスタム`SUM()`関数にアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="79136-105">Excel users can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="79136-106">前提条件</span><span class="sxs-lookup"><span data-stu-id="79136-106">Prerequisites</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="79136-107">カスタム関数の作成を開始するには、次のツールと関連するリソースが必要です。</span><span class="sxs-lookup"><span data-stu-id="79136-107">You'll need the following tools and related resources to begin creating custom functions.</span></span>

- <span data-ttu-id="79136-108">[Node.js](https://nodejs.org/en/) (バージョン 8.0.0 以降)</span><span class="sxs-lookup"><span data-stu-id="79136-108">[Node.js](https://nodejs.org/en/) (version 8.0.0 or later)</span></span>

- <span data-ttu-id="79136-109">[Git バッシュ](https://git-scm.com/downloads) (または別の Git クライアント)</span><span class="sxs-lookup"><span data-stu-id="79136-109">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

- <span data-ttu-id="79136-110">最新バージョンの [Yeoman](https://yeoman.io/) と [Office アドイン用の Yeoman ジェネレーター](https://www.npmjs.com/package/generator-office)。これらのツールをグローバルにインストールするには、コマンド プロンプトから次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="79136-110">The latest version of [Yeoman](https://yeoman.io/) and the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office). To install these tools globally, run the following command via the command prompt:</span></span>

    ```
    npm install -g yo generator-office
    ```

    > [!NOTE]
    > <span data-ttu-id="79136-111">以前に一度使用したバージョンのジェネレーターをインストールしていた場合でも、パッケージを npm から最新バージョンに更新することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="79136-111">Even if you've previously installed the Yeoman generator, we recommend you update your package to the latest version from npm.</span></span>

## <a name="build-your-first-custom-functions-project"></a><span data-ttu-id="79136-112">最初のカスタム関数プロジェクトを作成する</span><span class="sxs-lookup"><span data-stu-id="79136-112">Build your first custom functions project</span></span>

<span data-ttu-id="79136-113">はじめに、Yeoman ジェネレーターを使って、カスタム関数プロジェクトを作成します。</span><span class="sxs-lookup"><span data-stu-id="79136-113">To start, you'll use the Yeoman generator to create the custom functions project.</span></span> <span data-ttu-id="79136-114">これにより、カスタム関数のコーディングを開始するための正しいフォルダー構造、ソース ファイル、依存関係によるプロジェクトがセットアップされます。</span><span class="sxs-lookup"><span data-stu-id="79136-114">This will set up your project with the correct folder structure, source files, and dependencies to begin coding your custom functions.</span></span>

1. <span data-ttu-id="79136-115">次のコマンドを実行し、以下のようにプロンプトに応答します。</span><span class="sxs-lookup"><span data-stu-id="79136-115">Run the following command and then answer the prompts as follows.</span></span>

    ```
    yo office
    ```

    - <span data-ttu-id="79136-116">Choose a project type (プロジェクトの種類を選択): `Excel Custom Functions Add-in project (...)`</span><span class="sxs-lookup"><span data-stu-id="79136-116">Choose a project type: `Excel Custom Functions Add-in project (...)`</span></span>

    - <span data-ttu-id="79136-117">Choose a script type (スクリプトの種類を選択): `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="79136-117">Choose a script type: `JavaScript`</span></span>

    - <span data-ttu-id="79136-118">What would you want to name your add-in? (アドインの名前を何にしますか)</span><span class="sxs-lookup"><span data-stu-id="79136-118">What do you want to name your add-in?</span></span> `stock-ticker`

    ![カスタム関数の Office アドイン用の Yeoman ジェネレーターのプロンプト](../images/12-10-fork-cf-pic.jpg)

    <span data-ttu-id="79136-120">Yeoman ジェネレーターはプロジェクト ファイルを作成し、サポートしているノード コンポーネントをインストールします。</span><span class="sxs-lookup"><span data-stu-id="79136-120">The Yeoman generator will create the project files and install supporting Node components.</span></span>

2. <span data-ttu-id="79136-121">作成したばかりのプロジェクトフォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="79136-121">Navigate to the project folder you just created.</span></span>

    ```
    cd stock-ticker
    ```

3. <span data-ttu-id="79136-122">このプロジェクトを実行するには、自己署名証明書を信頼する必要があります。</span><span class="sxs-lookup"><span data-stu-id="79136-122">Trust the self-signed certificate you need to run this project.</span></span> <span data-ttu-id="79136-123">Windows または Mac についての詳細な手順については、「[自己署名証明書を信頼済みルート証明書として追加する](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="79136-123">For detailed instructions for either Windows or Mac, see [Adding Self Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span></span>  

4. <span data-ttu-id="79136-124">プロジェクトをビルドします。</span><span class="sxs-lookup"><span data-stu-id="79136-124">Build the project.</span></span>

    ```
    npm run build
    ```

5. <span data-ttu-id="79136-125">Node.js で実行しているローカル Web サーバーを開始します。</span><span class="sxs-lookup"><span data-stu-id="79136-125">Start the local web server, which runs in Node.js.</span></span>

    - <span data-ttu-id="79136-126">Windows 版 excel を使用してカスタム関数をテストする場合は、次のコマンドを実行してローカル web サーバーを起動し、Excel を起動して、アドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="79136-126">If you use Excel for Windows to test your custom functions, run the following command to start the local web server, launch Excel, and sideload the add-in:</span></span>

        ```
         npm run start
        ```
        <span data-ttu-id="79136-127">このコマンドを実行すると、コマンドプロンプトに web サーバーの起動に関する詳細が表示されます。</span><span class="sxs-lookup"><span data-stu-id="79136-127">After running this command, your command prompt will show details about starting the web server.</span></span> <span data-ttu-id="79136-128">Excel は、アドインが読み込まれた状態で起動します。</span><span class="sxs-lookup"><span data-stu-id="79136-128">Excel will start with your add-in loaded.</span></span> <span data-ttu-id="79136-129">アドインが読み込まれない場合は、手順 3 が正しく完了しているか確認してください。</span><span class="sxs-lookup"><span data-stu-id="79136-129">If you add-in does not load, check that you have completed step 3 properly.</span></span>

    - <span data-ttu-id="79136-130">Excel Online を使用してカスタム関数をテストする場合は、次のコマンドを実行してローカル web サーバーを開始します。</span><span class="sxs-lookup"><span data-stu-id="79136-130">If you use Excel Online to test your custom functions, run the following command to start the local web server:</span></span>

        ```
        npm run start-web
        ```

         <span data-ttu-id="79136-131">このコマンドを実行すると、コマンドプロンプトに web サーバーの起動に関する詳細が表示されます。</span><span class="sxs-lookup"><span data-stu-id="79136-131">After running this command, your command prompt will show details about starting the web server.</span></span> <span data-ttu-id="79136-132">関数を使用するには、Excel Online で新しいブックを開きます。</span><span class="sxs-lookup"><span data-stu-id="79136-132">To use your functions, open a new workbook in Excel Online.</span></span> <span data-ttu-id="79136-133">このブックでは、アドインを読み込む必要があります。</span><span class="sxs-lookup"><span data-stu-id="79136-133">In this workbook, you'll need to load your add-in.</span></span> 

        <span data-ttu-id="79136-134">これを行うには、リボンの [**挿入**] タブを選択して、[アドインの**取得**] を選択します。生成された新しいウィンドウで、[**マイアドイン**] タブが表示されていることを確認します。次に、[**個人用アドインの管理 > [個人用**アドインのアップロード] を選択します。</span><span class="sxs-lookup"><span data-stu-id="79136-134">To do this, select the **Insert** tab on the ribbon and select **Get Add-ins**. In the resulting new window, ensure you are on the **My Add-ins** tab. Next, select **Manage My Add-ins > Upload My Add-in**.</span></span> <span data-ttu-id="79136-135">マニフェストファイルを参照してアップロードします。</span><span class="sxs-lookup"><span data-stu-id="79136-135">Browse for your manifest file and upload it.</span></span> <span data-ttu-id="79136-136">アドインが読み込まれない場合は、手順3が正しく完了していることを確認してください。</span><span class="sxs-lookup"><span data-stu-id="79136-136">If your add-in does not load, check you've completed step 3 correctly.</span></span>

## <a name="try-out-the-prebuilt-custom-functions"></a><span data-ttu-id="79136-137">あらかじめ用意されているカスタム関数を試してみる</span><span class="sxs-lookup"><span data-stu-id="79136-137">Try out the prebuilt custom functions</span></span>

<span data-ttu-id="79136-138">Yeoman ジェネレーターで作成したカスタム関数プロジェクトには、あらかじめ用意されているカスタム関数がいくつか含まれており、**src/customfunctions.js** ファイル内で定義されています。</span><span class="sxs-lookup"><span data-stu-id="79136-138">The custom functions project that you created by using the Yeoman generator contains some prebuilt custom functions, defined within the **src/customfunctions.js** file.</span></span> <span data-ttu-id="79136-139">プロジェクトのルート ディレクトリの **manifest.xml** ファイルによって、カスタム関数はすべて `CONTOSO` 名前空間に属することが指定されます。</span><span class="sxs-lookup"><span data-stu-id="79136-139">The **manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="79136-140">Excel ブックで、次の手順を`ADD`実行してカスタム関数を試してみます。</span><span class="sxs-lookup"><span data-stu-id="79136-140">In your Excel workbook, try out the `ADD` custom function by completing the following steps:</span></span>

1. <span data-ttu-id="79136-141">セルを選択し、 `=CONTOSO`テキストを入力します。</span><span class="sxs-lookup"><span data-stu-id="79136-141">Select a cell and type `=CONTOSO`.</span></span> <span data-ttu-id="79136-142">`CONTOSO` 名前空間にあるすべての関数がオートコンプリート メニューに一覧表示されます。</span><span class="sxs-lookup"><span data-stu-id="79136-142">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="79136-143">セルに`CONTOSO.ADD`値`=CONTOSO.ADD(10,200)`を入力し`10` 、 `200` enter キーを押して、数値と入力パラメーターを使用して、関数を実行します。</span><span class="sxs-lookup"><span data-stu-id="79136-143">Run the `CONTOSO.ADD` function, using numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="79136-144">`ADD` カスタム関数によって、入力パラメーターとして指定した 2 つの数字の合計が計算されます。</span><span class="sxs-lookup"><span data-stu-id="79136-144">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters.</span></span> <span data-ttu-id="79136-145">「`=CONTOSO.ADD(10,200)`」と入力して Enter キーを押すと、**210** という結果が生成されるはずです。</span><span class="sxs-lookup"><span data-stu-id="79136-145">Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="next-steps"></a><span data-ttu-id="79136-146">次の手順</span><span class="sxs-lookup"><span data-stu-id="79136-146">Next steps</span></span>

<span data-ttu-id="79136-147">おめでとうございます。 Excel アドインでカスタム関数が正常に作成されました。</span><span class="sxs-lookup"><span data-stu-id="79136-147">Congratulations, you've successfully created a custom function in an Excel add-in!</span></span> <span data-ttu-id="79136-148">次に、ストリーミングデータ機能を使用して、より複雑なアドインをビルドします。</span><span class="sxs-lookup"><span data-stu-id="79136-148">Next, build a more complex add-in with streaming data capability.</span></span> <span data-ttu-id="79136-149">次のリンクでは、「カスタム関数を使用した Excel アドインのチュートリアル」の次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="79136-149">The following link takes you through the next steps in the Excel add-in with custom functions tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="79136-150">Excel カスタム関数アドインのチュートリアル</span><span class="sxs-lookup"><span data-stu-id="79136-150">Excel custom functions add-in tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md#create-a-custom-function-that-requests-data-from-the-web
)

## <a name="see-also"></a><span data-ttu-id="79136-151">関連項目</span><span class="sxs-lookup"><span data-stu-id="79136-151">See also</span></span>

* [<span data-ttu-id="79136-152">カスタム関数の概要</span><span class="sxs-lookup"><span data-stu-id="79136-152">Custom functions overview</span></span>](../excel/custom-functions-overview.md)
* [<span data-ttu-id="79136-153">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="79136-153">Custom functions metadata</span></span>](../excel/custom-functions-json.md)
* [<span data-ttu-id="79136-154">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="79136-154">Runtime for Excel custom functions</span></span>](../excel/custom-functions-runtime.md)
* [<span data-ttu-id="79136-155">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="79136-155">Custom functions best practices</span></span>](../excel/custom-functions-best-practices.md)
