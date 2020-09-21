---
title: Web ページから Excel を開き、Office アドインを埋め込む
description: Web ページから Excel を開き、Office アドインを埋め込みます。
ms.date: 09/15/2020
localization_priority: Normal
ms.openlocfilehash: 49df253c714f3ad84d2523b87e7df894b9027355
ms.sourcegitcommit: ea03e4ea2e8537d5f6d52477816209f6c1a6579c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/21/2020
ms.locfileid: "48166930"
---
# <a name="open-excel-from-your-web-page-and-embed-your-office-add-in"></a><span data-ttu-id="797fe-103">Web ページから Excel を開き、Office アドインを埋め込む</span><span class="sxs-lookup"><span data-stu-id="797fe-103">Open Excel from your web page and embed your Office Add-in</span></span>

:::image type="content" source="../images/pnp-open-in-excel.png" alt-text="Web ページ上の [Excel] ボタンのイメージアドインが埋め込まれた新しい Excel ドキュメントを開き、自動的に開きます。":::

<span data-ttu-id="797fe-105">SaaS web アプリケーションを拡張して、顧客が web ページから Microsoft Excel に直接データを開くことができるようにします。</span><span class="sxs-lookup"><span data-stu-id="797fe-105">Extend your SaaS web application so that your customers can open their data from a web page directly to Microsoft Excel.</span></span> <span data-ttu-id="797fe-106">一般的なシナリオは、ユーザーが web アプリケーション内のデータを操作することです。</span><span class="sxs-lookup"><span data-stu-id="797fe-106">A common scenario is that customers will be working with data in your web application.</span></span> <span data-ttu-id="797fe-107">その後、データを Excel ドキュメントにコピーします。</span><span class="sxs-lookup"><span data-stu-id="797fe-107">Then they’ll want to copy the data into an Excel document.</span></span> <span data-ttu-id="797fe-108">たとえば、Excel を使用して追加の分析を実行したい場合があります。</span><span class="sxs-lookup"><span data-stu-id="797fe-108">For example, they may want to perform additional analysis using Excel.</span></span> <span data-ttu-id="797fe-109">通常、お客様はデータをファイル (.csv ファイルなど) にエクスポートしてから、データを Excel にインポートする必要があります。</span><span class="sxs-lookup"><span data-stu-id="797fe-109">Typically, the customer is required to export the data to a file, such as a .csv file, and then import that data into Excel.</span></span> <span data-ttu-id="797fe-110">また、Office アドインをドキュメントに手動で追加する必要があります。</span><span class="sxs-lookup"><span data-stu-id="797fe-110">They also have to manually add your Office Add-in to the document.</span></span>

<span data-ttu-id="797fe-111">Excel ドキュメントを生成して開く、web ページ上の1回のボタンクリックに対して実行する手順の数を減らします。</span><span class="sxs-lookup"><span data-stu-id="797fe-111">Reduce the number of steps to a single button click on your web page that generates and opens the Excel document.</span></span> <span data-ttu-id="797fe-112">また、ドキュメントの内部に Office アドインを埋め込んで、ドキュメントを開いたときに表示することもできます。</span><span class="sxs-lookup"><span data-stu-id="797fe-112">You can also embed your Office Add-in inside the document and display it when the document opens.</span></span> <span data-ttu-id="797fe-113">これにより、お客様は引き続きアプリケーション機能にアクセスできるようになります。</span><span class="sxs-lookup"><span data-stu-id="797fe-113">This ensures the customer still has access to your application features.</span></span> <span data-ttu-id="797fe-114">ドキュメントが開いたときに、お客様が選択したデータが、Office アドインを引き続き使用できるようになります。</span><span class="sxs-lookup"><span data-stu-id="797fe-114">When the document opens, the data the customer selected, and your Office Add-in is already available for them to continue working.</span></span>

<span data-ttu-id="797fe-115">この記事では、独自の SaaS web アプリケーションでこのシナリオを実装するためのコードと手法について説明します。</span><span class="sxs-lookup"><span data-stu-id="797fe-115">This article shows you code and techniques for implementing this scenario in your own SaaS web application.</span></span>

## <a name="create-a-new-excel-document-and-embed-an-office-add-in"></a><span data-ttu-id="797fe-116">新しい Excel ドキュメントを作成し、Office アドインを埋め込む</span><span class="sxs-lookup"><span data-stu-id="797fe-116">Create a new Excel document and embed an Office Add-in</span></span>

<span data-ttu-id="797fe-117">最初に、web ページから Excel ドキュメントを作成し、アドインをドキュメントに埋め込む方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="797fe-117">First, let’s learn how to create an Excel document from a web page, and embed an add-in into the document.</span></span> <span data-ttu-id="797fe-118">[Office OOXML Embed アドインのコードサンプル](https://github.com/OfficeDev/Office-OOXML-EmbedAddin)は、[スクリプトラボアドイン](https://appsource.microsoft.com/product/office/wa104380862)を新しい Office ドキュメントに埋め込む方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="797fe-118">The [Office OOXML Embed Add-in code sample](https://github.com/OfficeDev/Office-OOXML-EmbedAddin) shows how to embed the [Script Lab add-in](https://appsource.microsoft.com/product/office/wa104380862) into a new Office document.</span></span> <span data-ttu-id="797fe-119">このサンプルは任意の Office ドキュメントで動作しますが、この記事の Excel スプレッドシートに重点を置いて説明します。</span><span class="sxs-lookup"><span data-stu-id="797fe-119">Although the sample works with any Office document, we’ll just focus on Excel spreadsheets in this article.</span></span> <span data-ttu-id="797fe-120">サンプルをビルドして実行するには、次の手順を使用します。</span><span class="sxs-lookup"><span data-stu-id="797fe-120">Use the following steps to build and run the sample.</span></span>

1. <span data-ttu-id="797fe-121">サンプルコードを  https://github.com/OfficeDev/Office-OOXML-EmbedAddin/archive/master.zip コンピューターのフォルダーに抽出します。</span><span class="sxs-lookup"><span data-stu-id="797fe-121">Extract the sample code from  https://github.com/OfficeDev/Office-OOXML-EmbedAddin/archive/master.zip into a folder on your computer.</span></span>
2. <span data-ttu-id="797fe-122">サンプルをビルドして実行するには、に記載されている手順に従って、readme の「プロジェクト」セクション **を使用** します。</span><span class="sxs-lookup"><span data-stu-id="797fe-122">To build and run the sample, follow the steps in the **To use the project** section of the readme.</span></span>
3. <span data-ttu-id="797fe-123">サンプルを実行すると、次のスクリーンショットに似た web ページが表示されます。</span><span class="sxs-lookup"><span data-stu-id="797fe-123">When you run the sample it will display a web page similar to the following screen shot.</span></span> <span data-ttu-id="797fe-124">Web ページを使用して、スクリプトラボが含まれる新しい Excel ドキュメントを作成します。</span><span class="sxs-lookup"><span data-stu-id="797fe-124">Use the web page to create a new Excel document that contains Script Lab when it opens.</span></span>
:::image type="content" source="../images/embed-script-lab-sample-ui.png" alt-text="埋め込みスクリプトラボサンプルが表示する web ページのスクリーンショットには、Excel ファイルを選択して、スクリプトラボアドインを埋め込むことができます。":::

### <a name="how-the-sample-works"></a><span data-ttu-id="797fe-126">サンプルの動作方法</span><span class="sxs-lookup"><span data-stu-id="797fe-126">How the sample works</span></span>

<span data-ttu-id="797fe-127">サンプルコードでは、OOXML SDK を使用して、選択した Excel ドキュメントにスクリプトラボアドインを埋め込みます。</span><span class="sxs-lookup"><span data-stu-id="797fe-127">The sample code uses the OOXML SDK to embed the Script Lab add-in to the Excel document that you choose.</span></span> <span data-ttu-id="797fe-128">次の情報は、readme ファイルの [ [**コードについて** ] セクション](https://github.com/OfficeDev/Office-OOXML-EmbedAddin/blob/master/README.md) から取得されています。</span><span class="sxs-lookup"><span data-stu-id="797fe-128">The following Information is taken from the [**About the code** section](https://github.com/OfficeDev/Office-OOXML-EmbedAddin/blob/master/README.md) in the readme file.</span></span>

<span data-ttu-id="797fe-129">ファイル **Home.aspx.cs**:</span><span class="sxs-lookup"><span data-stu-id="797fe-129">The file **Home.aspx.cs**:</span></span>

- <span data-ttu-id="797fe-130">ボタンイベントハンドラーと基本的な UI 操作を提供します。</span><span class="sxs-lookup"><span data-stu-id="797fe-130">Provides the button event handlers and basic UI manipulation.</span></span>
- <span data-ttu-id="797fe-131">は、標準の ASP.NET 技法を使用してファイルをアップロードおよびダウンロードします。</span><span class="sxs-lookup"><span data-stu-id="797fe-131">Uses standard ASP.NET techniques to upload and download the file.</span></span>
- <span data-ttu-id="797fe-132">アップロードされたファイルのファイル名拡張子 (.xlsx、.docx、または .pptx) を使用して、ファイルの種類を特定します。</span><span class="sxs-lookup"><span data-stu-id="797fe-132">Uses the file name extension of the uploaded file (xlsx, docx, or pptx) to determine the type of file.</span></span> <span data-ttu-id="797fe-133">通常、Open XML SDK には、ファイルの種類ごとに個別の Api が含まれているため、これを最初に実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="797fe-133">This needs to be done at the outset because the Open XML SDK generally has distinct APIs for each type of file.</span></span>
- <span data-ttu-id="797fe-134">**Ooxmlhelper**を呼び出してファイルを検証し、 **AddInEmbedder**を呼び出してスクリプトラボをファイルに埋め込み、自動的に開くように設定します。</span><span class="sxs-lookup"><span data-stu-id="797fe-134">Calls into the **OOXMLHelper** to validate the file and calls into the **AddInEmbedder** to embed Script Lab in the file and set to automatically open.</span></span>

<span data-ttu-id="797fe-135">ファイル **AddInEmbedder.cs**:</span><span class="sxs-lookup"><span data-stu-id="797fe-135">The file **AddInEmbedder.cs**:</span></span>

- <span data-ttu-id="797fe-136">主要なビジネスロジックを提供します。このサンプルでは、スクリプトラボを埋め込むメソッドを示します。</span><span class="sxs-lookup"><span data-stu-id="797fe-136">Provides the main business logic, which in this sample is a method that embeds Script Lab.</span></span>
- <span data-ttu-id="797fe-137">ファイルの種類に基づいて、OOXML ヘルパーに呼び出しを行います。</span><span class="sxs-lookup"><span data-stu-id="797fe-137">Makes calls into the OOXML helper based on the type of the file.</span></span>

<span data-ttu-id="797fe-138">ファイル **OOXMLHelper.cs**:</span><span class="sxs-lookup"><span data-stu-id="797fe-138">The file **OOXMLHelper.cs**:</span></span>

- <span data-ttu-id="797fe-139">詳細な OOXML 操作をすべて提供します。</span><span class="sxs-lookup"><span data-stu-id="797fe-139">Provides all the detailed OOXML manipulation.</span></span>
- <span data-ttu-id="797fe-140">Office ファイルを検証するための標準的な手法を使用します。この方法では、単に **ドキュメントの Open** メソッドを呼び出すことができます。</span><span class="sxs-lookup"><span data-stu-id="797fe-140">Uses a standard technique for validating the Office file, which is simply to call the **Document.Open** method on it.</span></span> <span data-ttu-id="797fe-141">ファイルが無効な場合、メソッドは例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="797fe-141">If the file is invalid, the method throws an exception.</span></span>
- <span data-ttu-id="797fe-142">Open xml 2.5 SDK 生産性ツールで生成された、 [OPEN xml 2.5 sdk](/office/open-xml/open-xml-sdk)のリンクで利用できる主なコードが含まれています。</span><span class="sxs-lookup"><span data-stu-id="797fe-142">Contains mainly code that was generated by the Open XML 2.5 SDK Productivity Tools which are available at the link for the [Open XML 2.5 SDK](/office/open-xml/open-xml-sdk).</span></span>

<span data-ttu-id="797fe-143">**OOXMLHelper.cs**ファイルの**GenerateWebExtensionPart1Content**メソッドは、Microsoft Appsource の Script Lab の ID への参照を設定します。</span><span class="sxs-lookup"><span data-stu-id="797fe-143">The **GenerateWebExtensionPart1Content** method in the **OOXMLHelper.cs** file sets the reference to the ID of Script Lab in Microsoft AppSource:</span></span>

```csharp
We.WebExtensionStoreReference webExtensionStoreReference1 = new We.WebExtensionStoreReference() { Id = "wa104380862", Version = "1.1.0.0", Store = "en-US", StoreType = "OMEX" };
```

- <span data-ttu-id="797fe-144">**Storetype**の値は、Microsoft appsource のエイリアスである "omex" です。</span><span class="sxs-lookup"><span data-stu-id="797fe-144">The **StoreType** value is "OMEX", an alias for Microsoft AppSource.</span></span>
- <span data-ttu-id="797fe-145">**Store**値は、スクリプトラボの Microsoft appsource culture セクションにある "en-us" です。</span><span class="sxs-lookup"><span data-stu-id="797fe-145">The **Store** value is "en-US" found in the Microsoft AppSource culture section for Script Lab.</span></span>
- <span data-ttu-id="797fe-146">**Id**値は、スクリプトラボの Microsoft appsource アセット Id です。</span><span class="sxs-lookup"><span data-stu-id="797fe-146">The **Id** value is the Microsoft AppSource asset ID for Script Lab.</span></span>

<span data-ttu-id="797fe-147">自動開きのファイル共有カタログからアドインを設定する場合は、別の値を使用します。</span><span class="sxs-lookup"><span data-stu-id="797fe-147">If you are setting up an add-in from a file share catalog for auto-open, you will use different values:</span></span>

<span data-ttu-id="797fe-148">**Storetype**の値は "FileSystem" です。</span><span class="sxs-lookup"><span data-stu-id="797fe-148">The **StoreType** value is "FileSystem".</span></span>

- <span data-ttu-id="797fe-149">**Store**値は、ネットワーク共有の URL です。たとえば、「 \\ \\ MyComputer \\ mysharedfolder」とします。</span><span class="sxs-lookup"><span data-stu-id="797fe-149">The **Store** value is the URL of the network share; for example, "\\\\MyComputer\\MySharedFolder".</span></span> <span data-ttu-id="797fe-150">これは、Office セキュリティセンターで、共有の信頼できるカタログアドレスとして表示される正確な URL である必要があります。</span><span class="sxs-lookup"><span data-stu-id="797fe-150">This should be the exact URL that appears as the share's Trusted Catalog Address in the Office Trust Center.</span></span>
- <span data-ttu-id="797fe-151">**Id**値は、アドインのマニフェストのアプリ id です。</span><span class="sxs-lookup"><span data-stu-id="797fe-151">The **Id** value is the app ID in the add-ins manifest.</span></span>
> [!NOTE]
> <span data-ttu-id="797fe-152">これらの属性の代替値の詳細については、「文書を使用して [作業ウィンドウを自動的に開く](../develop/automatically-open-a-task-pane-with-a-document.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="797fe-152">For more information about alternative values for these attributes, see [Automatically open a task pane with a document](../develop/automatically-open-a-task-pane-with-a-document.md).</span></span>

## <a name="use-the-fluent-ui"></a><span data-ttu-id="797fe-153">Fluent UI を使用する</span><span class="sxs-lookup"><span data-stu-id="797fe-153">Use the Fluent UI</span></span>

:::image type="content" source="../images/fluent-ui-wxp.png" alt-text="Word、Excel、および PowerPoint の Fluent UI アイコン。":::

<span data-ttu-id="797fe-155">ベストプラクティスとして、Fluent UI を使用して、ユーザーが Microsoft 製品間を移行できるようにします。</span><span class="sxs-lookup"><span data-stu-id="797fe-155">A best practice is to use the Fluent UI to help your users transition between Microsoft products.</span></span> <span data-ttu-id="797fe-156">Web ページから起動する Office アプリケーションを指定するには、常に Office アイコンを使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="797fe-156">You should always use an Office icon to indicate which Office application will be launched from your web page.</span></span> <span data-ttu-id="797fe-157">Excel のアイコンを使用して Excel アプリケーションを起動することを示すように、サンプルコードを変更してみましょう。</span><span class="sxs-lookup"><span data-stu-id="797fe-157">Let’s modify the sample code to use the Excel icon to indicate that it launches the Excel application.</span></span>

1. <span data-ttu-id="797fe-158">Visual Studio でサンプルを開きます。</span><span class="sxs-lookup"><span data-stu-id="797fe-158">Open the sample in Visual Studio.</span></span>
1. <span data-ttu-id="797fe-159">[ **Default.aspx** ] ページを開きます。</span><span class="sxs-lookup"><span data-stu-id="797fe-159">Open the **Home.aspx** page.</span></span>
1. <span data-ttu-id="797fe-160">フォーム上の [ダウンロード] ボタンである次のコードを検索します。</span><span class="sxs-lookup"><span data-stu-id="797fe-160">Find following code that is the download button on the form:</span></span>
    ```html
    <asp:Button ID="btnDownload" runat="server" Text="Download" OnClick="btnDownload_Click" /> 
    ```
1. <span data-ttu-id="797fe-161">ボタンコードを次のイメージタグに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="797fe-161">Replace the button code with the following image tag.</span></span>
    ```html
    <asp:Image  src="https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/excel_48x1.svg" width="48" height="48" ID="btnDownload" runat="server" OnClick="btnDownload_Click" AlternateText="Open in Microsoft Excel" role="button" ImageUrl=""/>  
    ```
1. <span data-ttu-id="797fe-162">**F5**キーを押します (または**デバッグ > デバッグを開始**します)。</span><span class="sxs-lookup"><span data-stu-id="797fe-162">Press **F5** (or **Debug > Start Debugging**).</span></span> <span data-ttu-id="797fe-163">ホームページが読み込まれると、アイコンが表示されます。</span><span class="sxs-lookup"><span data-stu-id="797fe-163">You'll see the icon appear when the home page loads.</span></span>

<span data-ttu-id="797fe-164">詳細については、「Fluent UI 開発者ポータルの [Office ブランドアイコン](https://developer.microsoft.com/fluentui#/styles/web/office-brand-icons) 」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="797fe-164">For more information, see [Office Brand Icons](https://developer.microsoft.com/fluentui#/styles/web/office-brand-icons) on the Fluent UI developer portal.</span></span>  

## <a name="upload-the-excel-document-to-microsoft-onedrive"></a><span data-ttu-id="797fe-165">Excel ドキュメントを Microsoft OneDrive にアップロードする</span><span class="sxs-lookup"><span data-stu-id="797fe-165">Upload the Excel document to Microsoft OneDrive</span></span>

<span data-ttu-id="797fe-166">お客様が OneDrive を使用している場合は、OneDrive に新しいドキュメントをアップロードすることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="797fe-166">We recommend uploading new documents to OneDrive if your customer uses OneDrive.</span></span> <span data-ttu-id="797fe-167">これにより、ドキュメントの検索と操作が容易になります。</span><span class="sxs-lookup"><span data-stu-id="797fe-167">This makes it easier for them to find and work with the documents.</span></span> <span data-ttu-id="797fe-168">新しいコードサンプルを作成し、Microsoft Graph SDK を使用して新しい Excel ドキュメントを OneDrive にアップロードする方法を確認しましょう。</span><span class="sxs-lookup"><span data-stu-id="797fe-168">Let’s create a new code sample and see how you can use the Microsoft Graph SDK to upload a new Excel document to OneDrive.</span></span>

### <a name="use-a-quick-start-to-build-a-new-microsoft-graph-web-application"></a><span data-ttu-id="797fe-169">クイックスタートを使用して新しい Microsoft Graph web アプリケーションを作成する</span><span class="sxs-lookup"><span data-stu-id="797fe-169">Use a quick-start to build a new Microsoft Graph web application</span></span>

1. <span data-ttu-id="797fe-170">に移動 [https://developer.microsoft.com/graph/quick-start](https://developer.microsoft.com/graph/quick-start) し、手順に従って、Office 365 サービスと対話するクイックスタートのコードサンプルを作成して開きます。</span><span class="sxs-lookup"><span data-stu-id="797fe-170">Go to [https://developer.microsoft.com/graph/quick-start](https://developer.microsoft.com/graph/quick-start) and follow the steps to create and open a quick start code sample that interacts with Office 365 services.</span></span>
1. <span data-ttu-id="797fe-171">[ **ステップ 1: 言語またはプラットフォームを選択**してください] で、[ **ASP.NET MVC**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="797fe-171">In **step 1: Pick you language or platform**, choose **ASP.NET MVC**.</span></span> <span data-ttu-id="797fe-172">この手順の手順では ASP.NET MVC オプションを使用していますが、手順は任意の言語またはプラットフォームに適用されるパターンに従います。</span><span class="sxs-lookup"><span data-stu-id="797fe-172">Although the steps in this procedure use the ASP.NET MVC option, the steps follow a pattern that apply to any language or platform.</span></span>
1. <span data-ttu-id="797fe-173">[ **手順 2: アプリ id とシークレットを取得する**] で、[ **アプリ id とシークレットを取得する**] を選択します。</span><span class="sxs-lookup"><span data-stu-id="797fe-173">In **step 2: Get an app ID and secret**, choose **Get an app ID and secret**.</span></span>
1. <span data-ttu-id="797fe-174">Microsoft 365 アカウントにサインインします。</span><span class="sxs-lookup"><span data-stu-id="797fe-174">Sign in to your Microsoft 365 account.</span></span>  
1. <span data-ttu-id="797fe-175">[ **アプリシークレットを保存** する] web ページで、アプリシークレットを、後で取得して使用できるファイルの場所に保存します。</span><span class="sxs-lookup"><span data-stu-id="797fe-175">On the **Please save your app secret** web page, save the app secret to a file location where you can retrieve and use it later.</span></span>
1. <span data-ttu-id="797fe-176">[ **取得] を選択して、クイックスタートに戻って**ください。</span><span class="sxs-lookup"><span data-stu-id="797fe-176">Choose **Got it, take me back to the quick start**.</span></span>
1. <span data-ttu-id="797fe-177">**手順 2: 登録に成功しました。**</span><span class="sxs-lookup"><span data-stu-id="797fe-177">In **step 2: Registration Successful!**</span></span> <span data-ttu-id="797fe-178">生成されたアプリシークレットを入力します。</span><span class="sxs-lookup"><span data-stu-id="797fe-178">Enter the generated app secret.</span></span>
1. <span data-ttu-id="797fe-179">**手順 3: コーディングを開始**するには、「 **SDK ベースのコードサンプルをダウンロードする**」を選択します。</span><span class="sxs-lookup"><span data-stu-id="797fe-179">In **step 3: Start coding**, choose **Download the SDK-based code sample**.</span></span>
1. <span data-ttu-id="797fe-180">ダウンロードした zip フォルダーをローカルフォルダーに展開します。</span><span class="sxs-lookup"><span data-stu-id="797fe-180">Extract the download zip folder into a local folder.</span></span>  
1. <span data-ttu-id="797fe-181">Visual Studio 2019 で graph-tutorial ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="797fe-181">Open the graph-tutorial.sln file in Visual Studio 2019.</span></span>
1. <span data-ttu-id="797fe-182">ソリューションをビルドして実行し、正しく動作していることを確認します。</span><span class="sxs-lookup"><span data-stu-id="797fe-182">Build and run the solution and confirm it is working correctly.</span></span> <span data-ttu-id="797fe-183">予定表 web ページを使用して、Microsoft 365 の予定表を表示できるようにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="797fe-183">You should be able to use the calendar web page to view your Microsoft 365 calendar.</span></span>

### <a name="upload-a-file-to-onedrive"></a><span data-ttu-id="797fe-184">OneDrive にファイルをアップロードする</span><span class="sxs-lookup"><span data-stu-id="797fe-184">Upload a file to OneDrive</span></span>

1. <span data-ttu-id="797fe-185">Visual Studio 2019 で **graph-tutorial** ソリューションを開き、 **PrivateSettings.config** ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="797fe-185">Open the **graph-tutorial.sln** solution in Visual Studio 2019, and open the **PrivateSettings.config** file.</span></span>
1. <span data-ttu-id="797fe-186">**Files.ReadWrite**   **Ida: appscopes**キーに新しいスコープファイルを追加して、次のコードのようにします。</span><span class="sxs-lookup"><span data-stu-id="797fe-186">Add a new scope **Files.ReadWrite** to the **ida:AppScopes** key so that it looks like the following code:</span></span>
    ```xml
    <add key="ida:AppScopes" value="User.Read Calendars.Read Files.ReadWrite " />
    ```
1. <span data-ttu-id="797fe-187">**人差し指**ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="797fe-187">Open the **Index.cshtml** file.</span></span>
1. <span data-ttu-id="797fe-188">OneDrive にファイルをアップロードするボタンを作成するには、次の ActionLink コードを挿入します。</span><span class="sxs-lookup"><span data-stu-id="797fe-188">Insert the following ActionLink code to create a button to upload a file to OneDrive.</span></span>
    ```razor
    @if (Request.IsAuthenticated)
    {
        <h4>Welcome @ViewBag.User.DisplayName!</h4>
        <p>Use the navigation bar at the top of the page to get started.</p>
        @Html.ActionLink("Click here to create a new file on OneDrive", "CreateOneDriveFile", "Home", new { area = "" }, new { @class = "btn btn-primary btn-large" })
    }
    ```
1. <span data-ttu-id="797fe-189">**HomeController.cs** ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="797fe-189">Open the **HomeController.cs** file.</span></span>
1. <span data-ttu-id="797fe-190">アクションリンクからの要求を処理するために、次のコードを挿入します。</span><span class="sxs-lookup"><span data-stu-id="797fe-190">Insert the following code to handle the request from the action link.</span></span>
    ```csharp
    public void CreateOneDriveFile()
        {
            using (var stream = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes("The contents of the file goes here.")))
            {
                var client = graph_tutorial.Helpers.GraphHelper.UploadFile("/test.txt", stream);
            }
        }
    ```
1. <span data-ttu-id="797fe-191">**GraphHelper.cs**ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="797fe-191">Open the **GraphHelper.cs** file.</span></span>
1. <span data-ttu-id="797fe-192">次のコードを挿入して、OneDrive に新しいファイルを作成するために Microsoft Graph API を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="797fe-192">Insert the following code to call the Microsoft Graph API to create a new file on OneDrive.</span></span>
    ```csharp
    public static async Task UploadFile(string fileName, System.IO.MemoryStream stream)
        {
           var graphClient = GetAuthenticatedClient();
            await graphClient.Me
                .Drive
                .Root
                .ItemWithPath(fileName)
                .Content
                .Request()
                .PutAsync<DriveItem>(stream);
            return;
        }
    ```
1. <span data-ttu-id="797fe-193">**F5**キーを押します (または**デバッグ > デバッグを開始**します)。</span><span class="sxs-lookup"><span data-stu-id="797fe-193">Press **F5** (or **Debug > Start Debugging**).</span></span> <span data-ttu-id="797fe-194">Web アプリケーションが起動します。</span><span class="sxs-lookup"><span data-stu-id="797fe-194">The web application will start.</span></span>
1. <span data-ttu-id="797fe-195">**[ここをクリックしてサインイン**] を選択し、サインインします。</span><span class="sxs-lookup"><span data-stu-id="797fe-195">Choose **Click here to sign in**, and sign in.</span></span>
1. <span data-ttu-id="797fe-196">**OneDrive で新しいファイルを作成するには、[ここをクリック**します] を選択します。</span><span class="sxs-lookup"><span data-stu-id="797fe-196">Choose **Click here to create a new file on OneDrive**.</span></span>
1. <span data-ttu-id="797fe-197">新しいブラウザーのタブを開いて、OneDrive アカウントにサインインします。</span><span class="sxs-lookup"><span data-stu-id="797fe-197">Open a new browser tab and sign in to your OneDrive account.</span></span> <span data-ttu-id="797fe-198">ルートフォルダーに test.txt ファイルが表示されます。</span><span class="sxs-lookup"><span data-stu-id="797fe-198">You'll see the test.txt file in the root folder.</span></span>

<span data-ttu-id="797fe-199">これで、ファイルを OneDrive にアップロードする方法を習得しました。このコードを再利用して、作成した Excel ドキュメントをアップロードすることができます。</span><span class="sxs-lookup"><span data-stu-id="797fe-199">Now that you've learned how to upload a file to OneDrive, you can reuse this code to upload any Excel document that you create.</span></span>

## <a name="additional-considerations-for-your-solution"></a><span data-ttu-id="797fe-200">ソリューションに関するその他の考慮事項</span><span class="sxs-lookup"><span data-stu-id="797fe-200">Additional considerations for your solution</span></span>

<span data-ttu-id="797fe-201">すべてのユーザーのソリューションは、テクノロジや方法によって異なります。</span><span class="sxs-lookup"><span data-stu-id="797fe-201">Everyone’s solution is different in terms of technologies and approaches.</span></span> <span data-ttu-id="797fe-202">次の考慮事項は、ソリューションを変更してドキュメントを開いたり、Office アドインを埋め込んだりする方法を計画するのに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="797fe-202">The following considerations will help you plan how to modify your solution to open documents and embed your Office Add-in.</span></span>

### <a name="create-a-new-excel-spreadsheet-from-the-web-page"></a><span data-ttu-id="797fe-203">Web ページから新しい Excel スプレッドシートを作成する</span><span class="sxs-lookup"><span data-stu-id="797fe-203">Create a new Excel spreadsheet from the web page</span></span>

<span data-ttu-id="797fe-204">このサンプルは、既存の Excel ドキュメントを変更します。</span><span class="sxs-lookup"><span data-stu-id="797fe-204">The sample modifies an existing Excel document.</span></span> <span data-ttu-id="797fe-205">一般的なシナリオでは、web ページから新しい Excel スプレッドシートを作成します。</span><span class="sxs-lookup"><span data-stu-id="797fe-205">A more common scenario is that you’ll create a new Excel spreadsheet from your web page.</span></span> <span data-ttu-id="797fe-206">新しいスプレッドシートを作成する方法については、「 **スプレッドシートドキュメントを作成** する」の「ファイル名を指定する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="797fe-206">You can find additional details on how to create a new spreadsheet in **Create a spreadsheet document** by providing a file name.</span></span> <span data-ttu-id="797fe-207">この記事では、ファイルをローカルで作成する方法について説明しますが、SpreadsheetDocument メソッドのオーバーロードを使用して、stream でファイルを作成することもできます。</span><span class="sxs-lookup"><span data-stu-id="797fe-207">This article shows how to create the file locally, but you can also create the file in a stream by using an overload on the SpreadsheetDocument.Create method.</span></span>

### <a name="read-custom-properties-when-your-add-in-starts"></a><span data-ttu-id="797fe-208">アドインの起動時にカスタムプロパティを読み取る</span><span class="sxs-lookup"><span data-stu-id="797fe-208">Read custom properties when your add-in starts</span></span>

<span data-ttu-id="797fe-209">このコードサンプルでは、OOXML SDK を使用して、新しい Excel ドキュメントにスニペット ID を格納します。</span><span class="sxs-lookup"><span data-stu-id="797fe-209">The code sample stores a snippet ID in the new Excel document using the OOXML SDK.</span></span> <span data-ttu-id="797fe-210">スクリプトラボは、Excel ドキュメントからスニペット ID を読み取り、開いたときにスニペットコードを表示します。</span><span class="sxs-lookup"><span data-stu-id="797fe-210">Script Lab reads the snippet ID from the Excel document and then displays that snippet code when it opens.</span></span> <span data-ttu-id="797fe-211">独自のアドイン (クエリ文字列、一時認証トークンなど) にカスタムプロパティを送信する必要がある場合があります。アドインを開始するときにカスタムプロパティを読み取る方法について詳しくは、「 **アドインの状態と設定を永続** 化する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="797fe-211">You may need to send custom properties to your own add-in (such as a query string, or temporary authentication token.) See **Persisting add-in state and settings** for complete details on how to read custom properties when your add-in starts.</span></span>

### <a name="initialize-the-excel-document-with-data"></a><span data-ttu-id="797fe-212">データを使用して Excel ドキュメントを初期化する</span><span class="sxs-lookup"><span data-stu-id="797fe-212">Initialize the Excel document with data</span></span>

<span data-ttu-id="797fe-213">通常、顧客が web サイトから Excel ドキュメントを開くと、ドキュメントに web サイトからのデータが含まれていると予想されます。</span><span class="sxs-lookup"><span data-stu-id="797fe-213">Typically, when the customer opens up an Excel document from your web site, they expect the document to contain some data from the web site.</span></span> <span data-ttu-id="797fe-214">ドキュメントにデータを書き込むには、いくつかの方法があります。</span><span class="sxs-lookup"><span data-stu-id="797fe-214">There are a couple of ways to write data into the document.</span></span>

- <span data-ttu-id="797fe-215">**OOXML SDK を使用**してデータを書き込みます。</span><span class="sxs-lookup"><span data-stu-id="797fe-215">**Use the OOXML SDK to write the data**.</span></span> <span data-ttu-id="797fe-216">SDK を使用して、ドキュメントに任意のデータを直接書き込むことができます。</span><span class="sxs-lookup"><span data-stu-id="797fe-216">You can use the SDK to directly write any data into the document.</span></span> <span data-ttu-id="797fe-217">この方法は、ドキュメントが開いているときにデータを使用できるようにする場合に便利です。</span><span class="sxs-lookup"><span data-stu-id="797fe-217">This approach is useful if you want the data to be available the instant the document is opened.</span></span>
- <span data-ttu-id="797fe-218">**カスタムクエリプロパティを Office アドインに渡し**ます。</span><span class="sxs-lookup"><span data-stu-id="797fe-218">**Pass a custom query property to your Office add-in**.</span></span> <span data-ttu-id="797fe-219">ドキュメントを生成するときに、必要なすべてのデータを取得するクエリ文字列を含む Office アドインのカスタムプロパティを埋め込みます。</span><span class="sxs-lookup"><span data-stu-id="797fe-219">When you generate the document, you embed a custom property for the Office add-in that contains a query string that retrieves all the required data.</span></span> <span data-ttu-id="797fe-220">アドインが開くと、クエリが取得され、クエリが実行され、Office JS API を使用してクエリの結果がドキュメントに挿入されます。</span><span class="sxs-lookup"><span data-stu-id="797fe-220">When your add-in opens, it retrieves the query, runs the query, and uses the Office JS API to insert the result of the query into the document.</span></span>

### <a name="working-with-the-ooxml-sdk"></a><span data-ttu-id="797fe-221">OOXML SDK を使用する</span><span class="sxs-lookup"><span data-stu-id="797fe-221">Working with the OOXML SDK</span></span>

<span data-ttu-id="797fe-222">OOXML SDK は .NET に基づいています。</span><span class="sxs-lookup"><span data-stu-id="797fe-222">The OOXML SDK is based on .NET.</span></span> <span data-ttu-id="797fe-223">Web アプリケーションが .NET に対応していない場合は、OOXML を操作するための別の方法を探す必要があります。</span><span class="sxs-lookup"><span data-stu-id="797fe-223">If your web application does not .NET, you’ll need to look for an alternative way to work with OOXML.</span></span>

<span data-ttu-id="797fe-224">Javascript 版の OOXML SDK は、 [javascript 用の OPEN XML sdk](https://archive.codeplex.com/?p=openxmlsdkjs)から入手できます。</span><span class="sxs-lookup"><span data-stu-id="797fe-224">There is a JavaScript version of the OOXML SDK available at [Open XML SDK for JavaScript](https://archive.codeplex.com/?p=openxmlsdkjs).</span></span>

<span data-ttu-id="797fe-225">OOXML コードを Azure 関数に配置して、.NET コードを web アプリケーションの他の部分と区別することができます。</span><span class="sxs-lookup"><span data-stu-id="797fe-225">You can place the OOXML code in an Azure function to separate the .NET code from the rest of your web application.</span></span> <span data-ttu-id="797fe-226">その後、Web アプリケーションから Azure 関数 (Excel ドキュメントを生成するため) を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="797fe-226">Then call the Azure function (to generate the Excel document) from your Web application.</span></span> <span data-ttu-id="797fe-227">Azure 関数の詳細については、「 [Azure 関数の概要](https://docs.microsoft.com/azure/azure-functions/functions-overview)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="797fe-227">For more information on Azure functions, see [An introduction to Azure Functions](https://docs.microsoft.com/azure/azure-functions/functions-overview).</span></span>

### <a name="simplify-authentication"></a><span data-ttu-id="797fe-228">認証を簡略化する</span><span class="sxs-lookup"><span data-stu-id="797fe-228">Simplify authentication</span></span>

<span data-ttu-id="797fe-229">通常、お客様は web アプリケーションでの作業時に認証され、サインインします。</span><span class="sxs-lookup"><span data-stu-id="797fe-229">Usually the customer will be authenticated and signed in when working in your web application.</span></span> <span data-ttu-id="797fe-230">ベストプラクティスとして、Office アドインを使用するために再度サインインする必要がないように、ドキュメントを開くときにサインインを続けることができます。</span><span class="sxs-lookup"><span data-stu-id="797fe-230">A best practice is to allow them to stay signed in when they open the document so they are not required to sign in again to use your Office Add-in.</span></span> <span data-ttu-id="797fe-231">このことを適切に処理するには、短時間の認証トークンをアドインに渡します。</span><span class="sxs-lookup"><span data-stu-id="797fe-231">A good way to handle this is pass a short-lived authentication token to the add-in.</span></span>

1. <span data-ttu-id="797fe-232">OOXML SDK を使用して、認証トークンをドキュメント内のカスタムプロパティとして保存します。</span><span class="sxs-lookup"><span data-stu-id="797fe-232">Use the OOXML SDK to save the authentication token as a custom property in the document.</span></span>
1. <span data-ttu-id="797fe-233">アドインの開始時に、ドキュメントからトークンを読み取ります。</span><span class="sxs-lookup"><span data-stu-id="797fe-233">Read the token from the document when the add-in starts.</span></span>
1. <span data-ttu-id="797fe-234">これで、アドインは顧客から追加の認証手順を必要とせずに、サービスに接続できます。</span><span class="sxs-lookup"><span data-stu-id="797fe-234">The add-in can then connect to your services without requiring any additional authentication steps from the customer.</span></span>

> [!WARNING]
> <span data-ttu-id="797fe-235">認証トークンをドキュメントに埋め込むと、承認されていないユーザーがトークンを入手できるセキュリティ上のリスクが生じます。</span><span class="sxs-lookup"><span data-stu-id="797fe-235">Embedding an authentication token in the document poses a security risk where an unauthorized user could obtain the token.</span></span> <span data-ttu-id="797fe-236">短時間の認証トークンを使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="797fe-236">We recommend you use a short-lived authentication token.</span></span> <span data-ttu-id="797fe-237">アドインが短時間トークンを使用している場合は、ドキュメントに保存されていない新しい認証トークンをすぐに要求する必要があります。</span><span class="sxs-lookup"><span data-stu-id="797fe-237">When the add-in uses the short-lived token, it should immediately request a new authentication token that is not saved in the document.</span></span>

## <a name="see-also"></a><span data-ttu-id="797fe-238">関連項目</span><span class="sxs-lookup"><span data-stu-id="797fe-238">See also</span></span>

- [<span data-ttu-id="797fe-239">Open XML SDK 2.5 for Office へようこそ</span><span class="sxs-lookup"><span data-stu-id="797fe-239">Welcome to the Open XML SDK 2.5 for Office</span></span>](/office/open-xml/open-xml-sdk)
- [<span data-ttu-id="797fe-240">ドキュメントで作業ウィンドウを自動的に開く</span><span class="sxs-lookup"><span data-stu-id="797fe-240">Automatically open a task pane with a document</span></span>](../develop/automatically-open-a-task-pane-with-a-document.md)
- [<span data-ttu-id="797fe-241">アドインの状態および設定を保持する</span><span class="sxs-lookup"><span data-stu-id="797fe-241">Persisting add-in state and settings</span></span>](../develop/persisting-add-in-state-and-settings.md)
- [<span data-ttu-id="797fe-242">ファイル名を指定してスプレッドシート ドキュメントを作成する</span><span class="sxs-lookup"><span data-stu-id="797fe-242">Create a spreadsheet document by providing a file name</span></span>](/office/open-xml/how-to-create-a-spreadsheet-document-by-providing-a-file-name)
