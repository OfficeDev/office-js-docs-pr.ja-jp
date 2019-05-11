---
ms.date: 05/08/2019
description: Excel クイックスタートガイドでのカスタム関数の開発。
title: カスタム関数のクイックスタート
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 233e1b608eda4a696b14d833fe4e071b2fcffd67
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952384"
---
# <a name="get-started-developing-excel-custom-functions"></a><span data-ttu-id="93a38-103">Excel カスタム関数の開発を始める</span><span class="sxs-lookup"><span data-stu-id="93a38-103">Get started developing Excel custom functions</span></span>

<span data-ttu-id="93a38-104">カスタム関数を使用すると、開発者は、JavaScript または Typescript でアドインの一部として定義することによって、Excel に新しい関数を追加できるようになります。</span><span class="sxs-lookup"><span data-stu-id="93a38-104">With custom functions, developers can now add new functions to Excel by defining them in JavaScript or Typescript as part of an add-in.</span></span> <span data-ttu-id="93a38-105">Excel ユーザーは、Excel の任意のネイティブ関数の場合と同じように、カスタム`SUM()`関数にアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="93a38-105">Excel users can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="93a38-106">前提条件</span><span class="sxs-lookup"><span data-stu-id="93a38-106">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

* <span data-ttu-id="93a38-107">Excel on Windows (64 ビットバージョン1810以降) または Excel Online</span><span class="sxs-lookup"><span data-stu-id="93a38-107">Excel on Windows (64-bit version 1810 or later) or Excel Online</span></span>

* <span data-ttu-id="93a38-108">[Office Insider プログラム](https://products.office.com/office-insider)に加入する (**Insider** レベル -- 以前は "Insider Fast" と呼ばれていたもの)</span><span class="sxs-lookup"><span data-stu-id="93a38-108">Join the [Office Insider program](https://products.office.com/office-insider) (**Insider** level -- formerly called "Insider Fast")</span></span>

## <a name="build-your-first-custom-functions-project"></a><span data-ttu-id="93a38-109">最初のカスタム関数プロジェクトを作成する</span><span class="sxs-lookup"><span data-stu-id="93a38-109">Build your first custom functions project</span></span>

<span data-ttu-id="93a38-110">はじめに、Yeoman ジェネレーターを使って、カスタム関数プロジェクトを作成します。</span><span class="sxs-lookup"><span data-stu-id="93a38-110">To start, you'll use the Yeoman generator to create the custom functions project.</span></span> <span data-ttu-id="93a38-111">これにより、カスタム関数のコーディングを開始するための正しいフォルダー構造、ソース ファイル、依存関係によるプロジェクトがセットアップされます。</span><span class="sxs-lookup"><span data-stu-id="93a38-111">This will set up your project with the correct folder structure, source files, and dependencies to begin coding your custom functions.</span></span>

1. <span data-ttu-id="93a38-112">任意のフォルダーで、次のコマンドを実行し、次のようにプロンプトに応答します。</span><span class="sxs-lookup"><span data-stu-id="93a38-112">In a folder of your choice, run the following command and then answer the prompts as follows.</span></span>

    ```command&nbsp;line
    yo office
    ```

    - <span data-ttu-id="93a38-113">**Choose a project type: (プロジェクトの種類を選択)** `Excel Custom Functions Add-in project (...)`</span><span class="sxs-lookup"><span data-stu-id="93a38-113">**Choose a project type:** `Excel Custom Functions Add-in project (...)`</span></span>
    - <span data-ttu-id="93a38-114">**Choose a script type: (スクリプトの種類を選択)** `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="93a38-114">**Choose a script type:** `JavaScript`</span></span>
    - <span data-ttu-id="93a38-115">**What would you want to name your add-in?: (アドインの名前を何にしますか)**</span><span class="sxs-lookup"><span data-stu-id="93a38-115">**What do you want to name your add-in?**</span></span> `stock-ticker`

    ![カスタム関数の Office アドイン用の Yeoman ジェネレーターのプロンプト](../images/yo-office-excel-cf.png)

    <span data-ttu-id="93a38-117">Yeoman ジェネレーターはプロジェクト ファイルを作成し、サポートしているノード コンポーネントをインストールします。</span><span class="sxs-lookup"><span data-stu-id="93a38-117">The Yeoman generator will create the project files and install supporting Node components.</span></span>

2. <span data-ttu-id="93a38-118">[ごみ箱] ジェネレーターでは、プロジェクトの処理に関するいくつかの命令がコマンドラインに表示されますが、それらは無視して、手順に従って続行します。</span><span class="sxs-lookup"><span data-stu-id="93a38-118">The Yeoman generator will give you some instructions in your command line about what to do with the project, but ignore them and continue to follow our instructions.</span></span> <span data-ttu-id="93a38-119">プロジェクトのルート フォルダーに移動します。</span><span class="sxs-lookup"><span data-stu-id="93a38-119">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd stock-ticker
    ```

3. <span data-ttu-id="93a38-120">プロジェクトをビルドします。</span><span class="sxs-lookup"><span data-stu-id="93a38-120">Build the project.</span></span> <span data-ttu-id="93a38-121">これにより、プロジェクトが正常に機能するために必要な証明書もインストールされます。</span><span class="sxs-lookup"><span data-stu-id="93a38-121">This will also install certificates that your project needs in order to function properly.</span></span> 

    ```command&nbsp;line
    npm run build
    ```

4. <span data-ttu-id="93a38-122">Node.js で実行しているローカル Web サーバーを開始します。</span><span class="sxs-lookup"><span data-stu-id="93a38-122">Start the local web server, which runs in Node.js.</span></span> <span data-ttu-id="93a38-123">カスタム関数アドインは、Windows または Excel Online で Excel で試すことができます。</span><span class="sxs-lookup"><span data-stu-id="93a38-123">You can try out the custom function add-in in Excel on Windows or Excel Online.</span></span> <span data-ttu-id="93a38-124">アドインの作業ウィンドウを開くように求められる場合がありますが、これはオプションです。</span><span class="sxs-lookup"><span data-stu-id="93a38-124">You may be prompted to open the add-in's task pane, although this is optional.</span></span> <span data-ttu-id="93a38-125">アドインの作業ウィンドウを開かなくても、カスタム関数を実行できます。</span><span class="sxs-lookup"><span data-stu-id="93a38-125">You can still run your custom functions without opening your add-in's task pane.</span></span>

> [!NOTE]
> <span data-ttu-id="93a38-126">Office アドインは、開発中であっても HTTP ではなく HTTPS を使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="93a38-126">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="93a38-127">を実行`npm run start:desktop`した後に証明書をインストールするように求めるメッセージが表示されたら、このメッセージに同意します。</span><span class="sxs-lookup"><span data-stu-id="93a38-127">If you are prompted to install a certificate after you run `npm run start:desktop`, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

# <a name="excel-on-windowstabexcel-windows"></a>[<span data-ttu-id="93a38-128">Windows 上の Excel</span><span class="sxs-lookup"><span data-stu-id="93a38-128">Excel on Windows</span></span>](#tab/excel-windows)

<span data-ttu-id="93a38-129">Windows の Excel でアドインをテストするには、次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="93a38-129">To test your add-in in Excel on Windows, run the following command.</span></span> <span data-ttu-id="93a38-130">このコマンドを実行すると、ローカル web サーバーが起動し、アドインが読み込まれた状態で Excel が開きます。</span><span class="sxs-lookup"><span data-stu-id="93a38-130">When you run this command, the local web server will start and Excel will open with your add-in loaded.</span></span>

```command&nbsp;line
npm run start:desktop
```

# <a name="excel-onlinetabexcel-online"></a>[<span data-ttu-id="93a38-131">Excel Online</span><span class="sxs-lookup"><span data-stu-id="93a38-131">Excel Online</span></span>](#tab/excel-online)

<span data-ttu-id="93a38-132">Excel Online でアドインをテストするには、次のコマンドを実行します。</span><span class="sxs-lookup"><span data-stu-id="93a38-132">To test your add-in in Excel Online, run the following command.</span></span> <span data-ttu-id="93a38-133">このコマンドを実行すると、ローカル Web サーバーが起動します。</span><span class="sxs-lookup"><span data-stu-id="93a38-133">When you run this command, the local web server will start.</span></span>

```command&nbsp;line
npm run start:web
```

> [!NOTE]
> <span data-ttu-id="93a38-134">Office アドインは、開発中であっても HTTP ではなく HTTPS を使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="93a38-134">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="93a38-135">を実行`npm run start:web`した後に証明書をインストールするように求めるメッセージが表示されたら、このメッセージに同意します。</span><span class="sxs-lookup"><span data-stu-id="93a38-135">If you are prompted to install a certificate after you run `npm run start:web`, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

<span data-ttu-id="93a38-136">カスタム関数アドインを使用するには、Excel Online で新しいブックを開きます。</span><span class="sxs-lookup"><span data-stu-id="93a38-136">To use your custom functions add-in, open a new workbook in Excel Online.</span></span> <span data-ttu-id="93a38-137">このブックでは、次の手順を実行して、アドインをサイドロードします。</span><span class="sxs-lookup"><span data-stu-id="93a38-137">In this workbook, complete the following steps to sideload your add-in.</span></span>

1. <span data-ttu-id="93a38-138">Excel Online で、**[挿入]** タブを選択して、**[アドイン]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="93a38-138">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.</span></span>

   ![[個人用アドイン] アイコンが強調表示された状態で Excel Online にリボンを挿入する](../images/excel-cf-online-register-add-in-1.png)
   
2. <span data-ttu-id="93a38-140">**[マイ アドインの管理]** を選択し、**[マイ アドインのアップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="93a38-140">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="93a38-141">**[参照...]** を選択し、Yeoman ジェネレーターによって作成されたプロジェクトのルート ディレクトリに移動します。</span><span class="sxs-lookup"><span data-stu-id="93a38-141">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="93a38-142">**manifest.xml** ファイルを選択し、**[開く]** を選択し、**[アップロード]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="93a38-142">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

---

## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="93a38-143">あらかじめ用意されているカスタム関数を試す</span><span class="sxs-lookup"><span data-stu-id="93a38-143">Try out a prebuilt custom function</span></span>

<span data-ttu-id="93a38-144">[ごみ箱] ジェネレーターを使用して作成したカスタム関数プロジェクトには、 **/src/functions/functions.js**ファイル内で定義されているいくつかのあらかじめ用意されたカスタム関数があります。</span><span class="sxs-lookup"><span data-stu-id="93a38-144">The custom functions project that you created by using the Yeoman generator contains some prebuilt custom functions, defined within the **./src/functions/functions.js** file.</span></span> <span data-ttu-id="93a38-145">プロジェクトのルートディレクトリの **./manifest¥ xml**ファイルは、すべてのカスタム関数が`CONTOSO`名前空間に属することを指定します。</span><span class="sxs-lookup"><span data-stu-id="93a38-145">The **./manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="93a38-146">Excel ブックで、次の手順を`ADD`実行してカスタム関数を試してみます。</span><span class="sxs-lookup"><span data-stu-id="93a38-146">In your Excel workbook, try out the `ADD` custom function by completing the following steps:</span></span>

1. <span data-ttu-id="93a38-147">セルを選択し、 `=CONTOSO`テキストを入力します。</span><span class="sxs-lookup"><span data-stu-id="93a38-147">Select a cell and type `=CONTOSO`.</span></span> <span data-ttu-id="93a38-148">`CONTOSO` 名前空間にあるすべての関数がオートコンプリート メニューに一覧表示されます。</span><span class="sxs-lookup"><span data-stu-id="93a38-148">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="93a38-149">セルに`CONTOSO.ADD`値`=CONTOSO.ADD(10,200)`を入力し`10` 、 `200` enter キーを押して、数値と入力パラメーターを使用して、関数を実行します。</span><span class="sxs-lookup"><span data-stu-id="93a38-149">Run the `CONTOSO.ADD` function, using numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="93a38-150">`ADD` カスタム関数によって、入力パラメーターとして指定した 2 つの数字の合計が計算されます。</span><span class="sxs-lookup"><span data-stu-id="93a38-150">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters.</span></span> <span data-ttu-id="93a38-151">「`=CONTOSO.ADD(10,200)`」と入力して Enter キーを押すと、**210** という結果が生成されるはずです。</span><span class="sxs-lookup"><span data-stu-id="93a38-151">Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="next-steps"></a><span data-ttu-id="93a38-152">次のステップ</span><span class="sxs-lookup"><span data-stu-id="93a38-152">Next steps</span></span>

<span data-ttu-id="93a38-153">おめでとうございます。 Excel アドインでカスタム関数が正常に作成されました。</span><span class="sxs-lookup"><span data-stu-id="93a38-153">Congratulations, you've successfully created a custom function in an Excel add-in!</span></span> <span data-ttu-id="93a38-154">次に、ストリーミングデータ機能を使用して、より複雑なアドインをビルドします。</span><span class="sxs-lookup"><span data-stu-id="93a38-154">Next, build a more complex add-in with streaming data capability.</span></span> <span data-ttu-id="93a38-155">次のリンクでは、「カスタム関数を使用した Excel アドインのチュートリアル」の次の手順を実行します。</span><span class="sxs-lookup"><span data-stu-id="93a38-155">The following link takes you through the next steps in the Excel add-in with custom functions tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="93a38-156">Excel カスタム関数アドインのチュートリアル</span><span class="sxs-lookup"><span data-stu-id="93a38-156">Excel custom functions add-in tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md#create-a-custom-function-that-requests-data-from-the-web
)

## <a name="see-also"></a><span data-ttu-id="93a38-157">関連項目</span><span class="sxs-lookup"><span data-stu-id="93a38-157">See also</span></span>

* [<span data-ttu-id="93a38-158">カスタム関数の概要</span><span class="sxs-lookup"><span data-stu-id="93a38-158">Custom functions overview</span></span>](../excel/custom-functions-overview.md)
* [<span data-ttu-id="93a38-159">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="93a38-159">Custom functions metadata</span></span>](../excel/custom-functions-json.md)
* [<span data-ttu-id="93a38-160">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="93a38-160">Runtime for Excel custom functions</span></span>](../excel/custom-functions-runtime.md)
* [<span data-ttu-id="93a38-161">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="93a38-161">Custom functions best practices</span></span>](../excel/custom-functions-best-practices.md)
