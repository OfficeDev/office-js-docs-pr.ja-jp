---
title: Web Excelファイルを開き、アドインOffice埋め込む
description: Web Excelを開き、アドインにOffice埋め込む。
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: 18f40b0030f4132a413a879e8b3419af49984b45
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349379"
---
# <a name="open-excel-from-your-web-page-and-embed-your-office-add-in"></a><span data-ttu-id="67972-103">Web Excelファイルを開き、アドインOffice埋め込む</span><span class="sxs-lookup"><span data-stu-id="67972-103">Open Excel from your web page and embed your Office Add-in</span></span>

:::image type="content" source="../images/pnp-open-in-excel.png" alt-text="アドインを埋Excel自動開く新しいドキュメントを開く web ページExcelボタンのイメージ。":::

<span data-ttu-id="67972-105">SaaS Web アプリケーションを拡張して、顧客が Web ページからユーザーに直接データを開Microsoft Excel。</span><span class="sxs-lookup"><span data-stu-id="67972-105">Extend your SaaS web application so that your customers can open their data from a web page directly to Microsoft Excel.</span></span> <span data-ttu-id="67972-106">一般的なシナリオは、顧客が Web アプリケーションのデータを操作することです。</span><span class="sxs-lookup"><span data-stu-id="67972-106">A common scenario is that customers will be working with data in your web application.</span></span> <span data-ttu-id="67972-107">次に、データを別のドキュメントにExcelします。</span><span class="sxs-lookup"><span data-stu-id="67972-107">Then they’ll want to copy the data into an Excel document.</span></span> <span data-ttu-id="67972-108">たとえば、このツールを使用して追加の分析を実行Excel。</span><span class="sxs-lookup"><span data-stu-id="67972-108">For example, they may want to perform additional analysis using Excel.</span></span> <span data-ttu-id="67972-109">通常、顧客はデータをファイル (.csv ファイルなど) にエクスポートし、そのデータを Excel にインポートする必要があります。</span><span class="sxs-lookup"><span data-stu-id="67972-109">Typically, the customer is required to export the data to a file, such as a .csv file, and then import that data into Excel.</span></span> <span data-ttu-id="67972-110">また、ドキュメントにアドインOffice手動で追加する必要があります。</span><span class="sxs-lookup"><span data-stu-id="67972-110">They also have to manually add your Office Add-in to the document.</span></span>

<span data-ttu-id="67972-111">ドキュメントを生成して開く Web ページのボタンを 1 回クリックする手順の数Excelします。</span><span class="sxs-lookup"><span data-stu-id="67972-111">Reduce the number of steps to a single button click on your web page that generates and opens the Excel document.</span></span> <span data-ttu-id="67972-112">また、ドキュメント内にOfficeアドインを埋め込み、ドキュメントが開くと表示できます。</span><span class="sxs-lookup"><span data-stu-id="67972-112">You can also embed your Office Add-in inside the document and display it when the document opens.</span></span> <span data-ttu-id="67972-113">これにより、顧客は引き続きアプリケーション機能にアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="67972-113">This ensures the customer still has access to your application features.</span></span> <span data-ttu-id="67972-114">ドキュメントが開くと、顧客が選択したデータと、Officeアドインが既に使用して作業を続行できます。</span><span class="sxs-lookup"><span data-stu-id="67972-114">When the document opens, the data the customer selected, and your Office Add-in is already available for them to continue working.</span></span>

<span data-ttu-id="67972-115">この記事では、このシナリオを独自の SaaS Web アプリケーションに実装するためのコードと手法について説明します。</span><span class="sxs-lookup"><span data-stu-id="67972-115">This article shows you code and techniques for implementing this scenario in your own SaaS web application.</span></span>

## <a name="create-a-new-excel-document-and-embed-an-office-add-in"></a><span data-ttu-id="67972-116">新しいドキュメントをExcelし、アドインOffice埋め込む</span><span class="sxs-lookup"><span data-stu-id="67972-116">Create a new Excel document and embed an Office Add-in</span></span>

<span data-ttu-id="67972-117">最初に、Web ページから Excelドキュメントを作成し、アドインをドキュメントに埋め込む方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="67972-117">First, let’s learn how to create an Excel document from a web page, and embed an add-in into the document.</span></span> <span data-ttu-id="67972-118">[次Office OOXML Embed アドイン](https://github.com/OfficeDev/Office-OOXML-EmbedAddin)のコード サンプルは、新[](https://appsource.microsoft.com/product/office/wa104380862)しいドキュメントにScript Labアドインを埋め込むOffice示しています。</span><span class="sxs-lookup"><span data-stu-id="67972-118">The [Office OOXML Embed Add-in code sample](https://github.com/OfficeDev/Office-OOXML-EmbedAddin) shows how to embed the [Script Lab add-in](https://appsource.microsoft.com/product/office/wa104380862) into a new Office document.</span></span> <span data-ttu-id="67972-119">このサンプルは、任意のドキュメントOffice機能しますが、この記事では、Excelスプレッドシートに焦点を当てる必要があります。</span><span class="sxs-lookup"><span data-stu-id="67972-119">Although the sample works with any Office document, we’ll just focus on Excel spreadsheets in this article.</span></span> <span data-ttu-id="67972-120">次の手順を使用して、サンプルをビルドして実行します。</span><span class="sxs-lookup"><span data-stu-id="67972-120">Use the following steps to build and run the sample.</span></span>

1. <span data-ttu-id="67972-121">サンプル コードをコンピューター上  https://github.com/OfficeDev/Office-OOXML-EmbedAddin/archive/master.zip のフォルダーに抽出します。</span><span class="sxs-lookup"><span data-stu-id="67972-121">Extract the sample code from  https://github.com/OfficeDev/Office-OOXML-EmbedAddin/archive/master.zip into a folder on your computer.</span></span>
2. <span data-ttu-id="67972-122">サンプルをビルドして実行するには、readme の 「プロジェクトを使用するには **」セクションの** 手順に従います。</span><span class="sxs-lookup"><span data-stu-id="67972-122">To build and run the sample, follow the steps in the **To use the project** section of the readme.</span></span>
3. <span data-ttu-id="67972-123">サンプルを実行すると、次のスクリーンショットのような Web ページが表示されます。</span><span class="sxs-lookup"><span data-stu-id="67972-123">When you run the sample it will display a web page similar to the following screenshot.</span></span> <span data-ttu-id="67972-124">Web ページを使用して、開Excelを含む新Script Labドキュメントを作成します。</span><span class="sxs-lookup"><span data-stu-id="67972-124">Use the web page to create a new Excel document that contains Script Lab when it opens.</span></span>
:::image type="content" source="../images/embed-script-lab-sample-ui.png" alt-text="埋め込みスクリプト ラボ サンプルが表示する web ページのスクリーン ショットで、Excel ファイルを選択し、スクリプト ラボ アドインを埋め込む。":::

### <a name="how-the-sample-works"></a><span data-ttu-id="67972-126">サンプルの動作</span><span class="sxs-lookup"><span data-stu-id="67972-126">How the sample works</span></span>

<span data-ttu-id="67972-127">サンプル コードでは、OOXML SDK を使用して、Script Labを選択したドキュメントExcel埋め込みします。</span><span class="sxs-lookup"><span data-stu-id="67972-127">The sample code uses the OOXML SDK to embed the Script Lab add-in to the Excel document that you choose.</span></span> <span data-ttu-id="67972-128">次の情報は、readme ファイルの [コード [**について** ]](https://github.com/OfficeDev/Office-OOXML-EmbedAddin/blob/master/README.md) セクションから取得されます。</span><span class="sxs-lookup"><span data-stu-id="67972-128">The following information is taken from the [**About the code** section](https://github.com/OfficeDev/Office-OOXML-EmbedAddin/blob/master/README.md) in the readme file.</span></span>

<span data-ttu-id="67972-129">**Home.aspx.cs ファイル**:</span><span class="sxs-lookup"><span data-stu-id="67972-129">The file **Home.aspx.cs**:</span></span>

- <span data-ttu-id="67972-130">ボタン イベント ハンドラーと基本的な UI 操作を提供します。</span><span class="sxs-lookup"><span data-stu-id="67972-130">Provides the button event handlers and basic UI manipulation.</span></span>
- <span data-ttu-id="67972-131">標準の ASP.NET を使用して、ファイルをアップロードおよびダウンロードします。</span><span class="sxs-lookup"><span data-stu-id="67972-131">Uses standard ASP.NET techniques to upload and download the file.</span></span>
- <span data-ttu-id="67972-132">アップロードされたファイルのファイル名の拡張子 (xlsx、docx、または pptx) を使用して、ファイルの種類を決定します。</span><span class="sxs-lookup"><span data-stu-id="67972-132">Uses the file name extension of the uploaded file (xlsx, docx, or pptx) to determine the type of file.</span></span> <span data-ttu-id="67972-133">Open XML SDK は通常、ファイルの種類ごとに異なる API を持つため、これは最初に行う必要があります。</span><span class="sxs-lookup"><span data-stu-id="67972-133">This needs to be done at the outset because the Open XML SDK generally has distinct APIs for each type of file.</span></span>
- <span data-ttu-id="67972-134">**OOXMLHelper** を呼び出してファイルを検証し **、AddInEmbedder** を呼び出してファイルにScript Labを埋め込み、自動的に開く設定を行います。</span><span class="sxs-lookup"><span data-stu-id="67972-134">Calls into the **OOXMLHelper** to validate the file and calls into the **AddInEmbedder** to embed Script Lab in the file and set to automatically open.</span></span>

<span data-ttu-id="67972-135">ファイル **AddInEmbedder.cs:**</span><span class="sxs-lookup"><span data-stu-id="67972-135">The file **AddInEmbedder.cs**:</span></span>

- <span data-ttu-id="67972-136">主なビジネス ロジックを提供します。このサンプルでは、このロジックを埋め込むScript Lab。</span><span class="sxs-lookup"><span data-stu-id="67972-136">Provides the main business logic, which in this sample is a method that embeds Script Lab.</span></span>
- <span data-ttu-id="67972-137">ファイルの種類に基づいて OOXML ヘルパーを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="67972-137">Makes calls into the OOXML helper based on the type of the file.</span></span>

<span data-ttu-id="67972-138">ファイル **OOXMLHelper.cs**:</span><span class="sxs-lookup"><span data-stu-id="67972-138">The file **OOXMLHelper.cs**:</span></span>

- <span data-ttu-id="67972-139">すべての詳細な OOXML 操作を提供します。</span><span class="sxs-lookup"><span data-stu-id="67972-139">Provides all the detailed OOXML manipulation.</span></span>
- <span data-ttu-id="67972-140">標準の手法を使用して、Officeファイルを検証します。これは単に **Document.Open メソッドを呼** び出す方法です。</span><span class="sxs-lookup"><span data-stu-id="67972-140">Uses a standard technique for validating the Office file, which is simply to call the **Document.Open** method on it.</span></span> <span data-ttu-id="67972-141">ファイルが無効な場合、メソッドは例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="67972-141">If the file is invalid, the method throws an exception.</span></span>
- <span data-ttu-id="67972-142">Open [XML 2.5 SDK](/office/open-xml/open-xml-sdk)のリンクで使用できる Open XML 2.5 SDK 生産性向上ツールによって生成されたコードが主に含まれています。</span><span class="sxs-lookup"><span data-stu-id="67972-142">Contains mainly code that was generated by the Open XML 2.5 SDK Productivity Tools which are available at the link for the [Open XML 2.5 SDK](/office/open-xml/open-xml-sdk).</span></span>

<span data-ttu-id="67972-143">**OOXMLHelper.cs** ファイルの **GenerateWebExtensionPart1Content** メソッドは、Microsoft AppSource の id Script Labを設定します。</span><span class="sxs-lookup"><span data-stu-id="67972-143">The **GenerateWebExtensionPart1Content** method in the **OOXMLHelper.cs** file sets the reference to the ID of Script Lab in Microsoft AppSource:</span></span>

```csharp
We.WebExtensionStoreReference webExtensionStoreReference1 = new We.WebExtensionStoreReference() { Id = "wa104380862", Version = "1.1.0.0", Store = "en-US", StoreType = "OMEX" };
```

- <span data-ttu-id="67972-144">**StoreType 値** は、Microsoft AppSource のエイリアスである "OMEX" です。</span><span class="sxs-lookup"><span data-stu-id="67972-144">The **StoreType** value is "OMEX", an alias for Microsoft AppSource.</span></span>
- <span data-ttu-id="67972-145">ストア **の** 値は、Microsoft AppSource カルチャ セクションにある "en-US" Script Lab。</span><span class="sxs-lookup"><span data-stu-id="67972-145">The **Store** value is "en-US" found in the Microsoft AppSource culture section for Script Lab.</span></span>
- <span data-ttu-id="67972-146">Id **値** は、Microsoft AppSource アセット ID の値Script Lab。</span><span class="sxs-lookup"><span data-stu-id="67972-146">The **Id** value is the Microsoft AppSource asset ID for Script Lab.</span></span>

<span data-ttu-id="67972-147">ファイル共有カタログから自動開き用にアドインを設定する場合は、次の異なる値を使用します。</span><span class="sxs-lookup"><span data-stu-id="67972-147">If you are setting up an add-in from a file share catalog for auto-open, you will use different values:</span></span>

<span data-ttu-id="67972-148">**StoreType 値** は "FileSystem" です。</span><span class="sxs-lookup"><span data-stu-id="67972-148">The **StoreType** value is "FileSystem".</span></span>

- <span data-ttu-id="67972-149">Store **の** 値は、ネットワーク共有の URL です。たとえば \\ \\ 、「MyComputer \\ MySharedFolder」などです。</span><span class="sxs-lookup"><span data-stu-id="67972-149">The **Store** value is the URL of the network share; for example, "\\\\MyComputer\\MySharedFolder".</span></span> <span data-ttu-id="67972-150">これは、信頼センターで共有の信頼済みカタログ アドレスとして表示される正確な URL Office必要があります。</span><span class="sxs-lookup"><span data-stu-id="67972-150">This should be the exact URL that appears as the share's Trusted Catalog Address in the Office Trust Center.</span></span>
- <span data-ttu-id="67972-151">Id **値** は、アドイン マニフェストのアプリ ID です。</span><span class="sxs-lookup"><span data-stu-id="67972-151">The **Id** value is the app ID in the add-ins manifest.</span></span>
> [!NOTE]
> <span data-ttu-id="67972-152">これらの属性の代替値の詳細については、「ドキュメントを含む作業ウィンドウを自動的に [開く」を参照してください](../develop/automatically-open-a-task-pane-with-a-document.md)。</span><span class="sxs-lookup"><span data-stu-id="67972-152">For more information about alternative values for these attributes, see [Automatically open a task pane with a document](../develop/automatically-open-a-task-pane-with-a-document.md).</span></span>

## <a name="use-the-fluent-ui"></a><span data-ttu-id="67972-153">UI のFluentする</span><span class="sxs-lookup"><span data-stu-id="67972-153">Use the Fluent UI</span></span>

:::image type="content" source="../images/fluent-ui-wxp.png" alt-text="FluentWord、Excel、およびPowerPoint。":::

<span data-ttu-id="67972-155">ベスト プラクティスは、ユーザーが Microsoft 製品間で移行Fluent UI を使用する方法です。</span><span class="sxs-lookup"><span data-stu-id="67972-155">A best practice is to use the Fluent UI to help your users transition between Microsoft products.</span></span> <span data-ttu-id="67972-156">Web ページから起動するアプリケーションOfficeを示Officeアイコンを常に使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="67972-156">You should always use an Office icon to indicate which Office application will be launched from your web page.</span></span> <span data-ttu-id="67972-157">サンプル コードを変更して、Excelアイコンを使用して、アプリケーションを起動Excelします。</span><span class="sxs-lookup"><span data-stu-id="67972-157">Let’s modify the sample code to use the Excel icon to indicate that it launches the Excel application.</span></span>

1. <span data-ttu-id="67972-158">サンプルを [サンプル] で開Visual Studio。</span><span class="sxs-lookup"><span data-stu-id="67972-158">Open the sample in Visual Studio.</span></span>
1. <span data-ttu-id="67972-159">**[Home.aspx] ページを開** きます。</span><span class="sxs-lookup"><span data-stu-id="67972-159">Open the **Home.aspx** page.</span></span>
1. <span data-ttu-id="67972-160">フォームのダウンロード ボタンである次のコードを見つける。</span><span class="sxs-lookup"><span data-stu-id="67972-160">Find following code that is the download button on the form.</span></span>

    ```html
    <asp:Button ID="btnDownload" runat="server" Text="Download" OnClick="btnDownload_Click" /> 
    ```

1. <span data-ttu-id="67972-161">ボタン コードを次のイメージ タグに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="67972-161">Replace the button code with the following image tag.</span></span>

    ```html
    <asp:Image  src="https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/excel_48x1.svg" width="48" height="48" ID="btnDownload" runat="server" OnClick="btnDownload_Click" AlternateText="Open in Microsoft Excel" role="button" ImageUrl=""/>  
    ```

1. <span data-ttu-id="67972-162">**F5 キーを押** します (**または [デバッグ] >を開始します**)。</span><span class="sxs-lookup"><span data-stu-id="67972-162">Press **F5** (or **Debug > Start Debugging**).</span></span> <span data-ttu-id="67972-163">ホーム ページが読み込まれると、アイコンが表示されます。</span><span class="sxs-lookup"><span data-stu-id="67972-163">You'll see the icon appear when the home page loads.</span></span>

<span data-ttu-id="67972-164">詳細については、「UI 開発者[ポータルOfficeブランド](https://developer.microsoft.com/fluentui#/styles/web/office-brand-icons)アイコンFluent参照してください。</span><span class="sxs-lookup"><span data-stu-id="67972-164">For more information, see [Office Brand Icons](https://developer.microsoft.com/fluentui#/styles/web/office-brand-icons) on the Fluent UI developer portal.</span></span>  

## <a name="upload-the-excel-document-to-microsoft-onedrive"></a><span data-ttu-id="67972-165">アップロードドキュメントをExcelするMicrosoft OneDrive</span><span class="sxs-lookup"><span data-stu-id="67972-165">Upload the Excel document to Microsoft OneDrive</span></span>

<span data-ttu-id="67972-166">顧客がドキュメントを使用している場合は、OneDriveに新しいドキュメントをアップロードOneDrive。</span><span class="sxs-lookup"><span data-stu-id="67972-166">We recommend uploading new documents to OneDrive if your customer uses OneDrive.</span></span> <span data-ttu-id="67972-167">これにより、ドキュメントの検索と作業が容易になります。</span><span class="sxs-lookup"><span data-stu-id="67972-167">This makes it easier for them to find and work with the documents.</span></span> <span data-ttu-id="67972-168">新しいコード サンプルを作成し、Microsoft Graph SDK を使用して新しいドキュメントをExcelする方法OneDrive。</span><span class="sxs-lookup"><span data-stu-id="67972-168">Let’s create a new code sample and see how you can use the Microsoft Graph SDK to upload a new Excel document to OneDrive.</span></span>

### <a name="use-a-quick-start-to-build-a-new-microsoft-graph-web-application"></a><span data-ttu-id="67972-169">クイック スタートを使用して新しい Microsoft Graph Web アプリケーションを構築する</span><span class="sxs-lookup"><span data-stu-id="67972-169">Use a quick-start to build a new Microsoft Graph web application</span></span>

1. <span data-ttu-id="67972-170">手順に従って、クイック スタート コード サンプルを作成して開き、サービスを操作Office [https://developer.microsoft.com/graph/quick-start](https://developer.microsoft.com/graph/quick-start) します。</span><span class="sxs-lookup"><span data-stu-id="67972-170">Go to [https://developer.microsoft.com/graph/quick-start](https://developer.microsoft.com/graph/quick-start) and follow the steps to create and open a quick start code sample that interacts with Office services.</span></span>
1. <span data-ttu-id="67972-171">手順 **1: 言語またはプラットフォームを選択し、[MVC]** を ASP.NET **します**。</span><span class="sxs-lookup"><span data-stu-id="67972-171">In **step 1: Pick you language or platform**, choose **ASP.NET MVC**.</span></span> <span data-ttu-id="67972-172">この手順の手順では、MVC ASP.NET を使用しますが、この手順は、任意の言語またはプラットフォームに適用されるパターンに従います。</span><span class="sxs-lookup"><span data-stu-id="67972-172">Although the steps in this procedure use the ASP.NET MVC option, the steps follow a pattern that apply to any language or platform.</span></span>
1. <span data-ttu-id="67972-173">手順 **2: アプリ ID とシークレットを取得** し、[アプリ ID **とシークレットを取得する] を選択します**。</span><span class="sxs-lookup"><span data-stu-id="67972-173">In **step 2: Get an app ID and secret**, choose **Get an app ID and secret**.</span></span>
1. <span data-ttu-id="67972-174">アカウントにサインインMicrosoft 365します。</span><span class="sxs-lookup"><span data-stu-id="67972-174">Sign in to your Microsoft 365 account.</span></span>  
1. <span data-ttu-id="67972-175">[アプリ **シークレット Web ページを保存してください** ] で、アプリ シークレットを後で取得して使用できるファイルの場所に保存します。</span><span class="sxs-lookup"><span data-stu-id="67972-175">On the **Please save your app secret** web page, save the app secret to a file location where you can retrieve and use it later.</span></span>
1. <span data-ttu-id="67972-176">[Got **it] を選択し、クイック スタートに戻します**。</span><span class="sxs-lookup"><span data-stu-id="67972-176">Choose **Got it, take me back to the quick start**.</span></span>
1. <span data-ttu-id="67972-177">手順 **2: 登録が成功しました!**</span><span class="sxs-lookup"><span data-stu-id="67972-177">In **step 2: Registration Successful!**</span></span> <span data-ttu-id="67972-178">生成されたアプリ シークレットを入力します。</span><span class="sxs-lookup"><span data-stu-id="67972-178">Enter the generated app secret.</span></span>
1. <span data-ttu-id="67972-179">手順 **3: コーディングを開始し\*\*\*\*、[SDK ベースのコード サンプルをダウンロードする] を選択します**。</span><span class="sxs-lookup"><span data-stu-id="67972-179">In **step 3: Start coding**, choose **Download the SDK-based code sample**.</span></span>
1. <span data-ttu-id="67972-180">ダウンロード zip フォルダーをローカル フォルダーに展開します。</span><span class="sxs-lookup"><span data-stu-id="67972-180">Extract the download zip folder into a local folder.</span></span>  
1. <span data-ttu-id="67972-181">2019 年に graph-tutorial.sln ファイルを開Visual Studioします。</span><span class="sxs-lookup"><span data-stu-id="67972-181">Open the graph-tutorial.sln file in Visual Studio 2019.</span></span>
1. <span data-ttu-id="67972-182">ソリューションをビルドして実行し、正しく動作しているのを確認します。</span><span class="sxs-lookup"><span data-stu-id="67972-182">Build and run the solution and confirm it is working correctly.</span></span> <span data-ttu-id="67972-183">予定表 Web ページを使用して、予定表を表示Microsoft 365があります。</span><span class="sxs-lookup"><span data-stu-id="67972-183">You should be able to use the calendar web page to view your Microsoft 365 calendar.</span></span>

### <a name="upload-a-file-to-onedrive"></a><span data-ttu-id="67972-184">アップロードを作成するOneDrive</span><span class="sxs-lookup"><span data-stu-id="67972-184">Upload a file to OneDrive</span></span>

1. <span data-ttu-id="67972-185">2019 年 2019 年に **graph-tutorial.sln** ソリューションを開きVisual Studioファイル **をPrivateSettings.config** します。</span><span class="sxs-lookup"><span data-stu-id="67972-185">Open the **graph-tutorial.sln** solution in Visual Studio 2019, and open the **PrivateSettings.config** file.</span></span>
1. <span data-ttu-id="67972-186">次のコードのように、新しいスコープ **Files.ReadWrite** を   **ida:AppScopes** キーに追加します。</span><span class="sxs-lookup"><span data-stu-id="67972-186">Add a new scope **Files.ReadWrite** to the **ida:AppScopes** key so that it looks like the following code.</span></span>

    ```xml
    <add key="ida:AppScopes" value="User.Read Calendars.Read Files.ReadWrite " />
    ```

1. <span data-ttu-id="67972-187">**Index.cshtml ファイルを開** きます。</span><span class="sxs-lookup"><span data-stu-id="67972-187">Open the **Index.cshtml** file.</span></span>
1. <span data-ttu-id="67972-188">次の ActionLink コードを挿入して、ファイルをファイルにアップロードするボタンを作成OneDrive。</span><span class="sxs-lookup"><span data-stu-id="67972-188">Insert the following ActionLink code to create a button to upload a file to OneDrive.</span></span>

    ```razor
    @if (Request.IsAuthenticated)
    {
        <h4>Welcome @ViewBag.User.DisplayName!</h4>
        <p>Use the navigation bar at the top of the page to get started.</p>
        @Html.ActionLink("Click here to create a new file on OneDrive", "CreateOneDriveFile", "Home", new { area = "" }, new { @class = "btn btn-primary btn-large" })
    }
    ```

1. <span data-ttu-id="67972-189">**HomeController.cs** ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="67972-189">Open the **HomeController.cs** file.</span></span>
1. <span data-ttu-id="67972-190">アクション リンクからの要求を処理するには、次のコードを挿入します。</span><span class="sxs-lookup"><span data-stu-id="67972-190">Insert the following code to handle the request from the action link.</span></span>

    ```csharp
    public void CreateOneDriveFile()
        {
            using (var stream = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes("The contents of the file goes here.")))
            {
                var client = graph_tutorial.Helpers.GraphHelper.UploadFile("/test.txt", stream);
            }
        }
    ```

1. <span data-ttu-id="67972-191">**GraphHelper.cs ファイルを開** きます。</span><span class="sxs-lookup"><span data-stu-id="67972-191">Open the **GraphHelper.cs** file.</span></span>
1. <span data-ttu-id="67972-192">次のコードを挿入して、Microsoft Graph API を呼び出して、新しいファイルを作成OneDrive。</span><span class="sxs-lookup"><span data-stu-id="67972-192">Insert the following code to call the Microsoft Graph API to create a new file on OneDrive.</span></span>

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

1. <span data-ttu-id="67972-193">**F5 キーを押** します (**または [デバッグ] >を開始します**)。</span><span class="sxs-lookup"><span data-stu-id="67972-193">Press **F5** (or **Debug > Start Debugging**).</span></span> <span data-ttu-id="67972-194">Web アプリケーションが起動します。</span><span class="sxs-lookup"><span data-stu-id="67972-194">The web application will start.</span></span>
1. <span data-ttu-id="67972-195">[ **ここをクリックしてサインイン] を選択し**、サインインします。</span><span class="sxs-lookup"><span data-stu-id="67972-195">Choose **Click here to sign in**, and sign in.</span></span>
1. <span data-ttu-id="67972-196">[**ここをクリックして新しいファイルを作成する] をOneDrive。**</span><span class="sxs-lookup"><span data-stu-id="67972-196">Choose **Click here to create a new file on OneDrive**.</span></span>
1. <span data-ttu-id="67972-197">新しいブラウザー タブを開き、アカウントにサインインOneDriveします。</span><span class="sxs-lookup"><span data-stu-id="67972-197">Open a new browser tab and sign in to your OneDrive account.</span></span> <span data-ttu-id="67972-198">ルート フォルダーにtest.txtファイルが表示されます。</span><span class="sxs-lookup"><span data-stu-id="67972-198">You'll see the test.txt file in the root folder.</span></span>

<span data-ttu-id="67972-199">ファイルを OneDrive にアップロードする方法を学んだので、このコードを再利用して、作成Excelドキュメントをアップロードできます。</span><span class="sxs-lookup"><span data-stu-id="67972-199">Now that you've learned how to upload a file to OneDrive, you can reuse this code to upload any Excel document that you create.</span></span>

## <a name="additional-considerations-for-your-solution"></a><span data-ttu-id="67972-200">ソリューションに関するその他の考慮事項</span><span class="sxs-lookup"><span data-stu-id="67972-200">Additional considerations for your solution</span></span>

<span data-ttu-id="67972-201">すべてのユーザーのソリューションは、テクノロジとアプローチの点で異なります。</span><span class="sxs-lookup"><span data-stu-id="67972-201">Everyone’s solution is different in terms of technologies and approaches.</span></span> <span data-ttu-id="67972-202">次の考慮事項は、ソリューションを変更してドキュメントを開き、アドインを埋め込むOfficeに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="67972-202">The following considerations will help you plan how to modify your solution to open documents and embed your Office Add-in.</span></span>

### <a name="create-a-new-excel-spreadsheet-from-the-web-page"></a><span data-ttu-id="67972-203">Web ページからExcel新しいスプレッドシートを作成する</span><span class="sxs-lookup"><span data-stu-id="67972-203">Create a new Excel spreadsheet from the web page</span></span>

<span data-ttu-id="67972-204">このサンプルでは、既存のドキュメントをExcelします。</span><span class="sxs-lookup"><span data-stu-id="67972-204">The sample modifies an existing Excel document.</span></span> <span data-ttu-id="67972-205">より一般的なシナリオは、Web ページから新しいExcel作成する場合です。</span><span class="sxs-lookup"><span data-stu-id="67972-205">A more common scenario is that you’ll create a new Excel spreadsheet from your web page.</span></span> <span data-ttu-id="67972-206">新しいスプレッドシートを作成する方法の詳細については、「ファイル名を指定してスプレッドシート ドキュメントを作成 **する」を** 参照してください。</span><span class="sxs-lookup"><span data-stu-id="67972-206">You can find additional details on how to create a new spreadsheet in **Create a spreadsheet document** by providing a file name.</span></span> <span data-ttu-id="67972-207">この記事では、ファイルをローカルに作成する方法を示しますが、SpreadsheetDocument.Create メソッドでオーバーロードを使用して、ストリーム内にファイルを作成することもできます。</span><span class="sxs-lookup"><span data-stu-id="67972-207">This article shows how to create the file locally, but you can also create the file in a stream by using an overload on the SpreadsheetDocument.Create method.</span></span>

### <a name="read-custom-properties-when-your-add-in-starts"></a><span data-ttu-id="67972-208">アドインの起動時にカスタム プロパティを読み取る</span><span class="sxs-lookup"><span data-stu-id="67972-208">Read custom properties when your add-in starts</span></span>

<span data-ttu-id="67972-209">コード サンプルでは、OOXML SDK を使用して、新Excelドキュメントにスニペット ID を格納します。</span><span class="sxs-lookup"><span data-stu-id="67972-209">The code sample stores a snippet ID in the new Excel document using the OOXML SDK.</span></span> <span data-ttu-id="67972-210">Script Labドキュメントからスニペット ID を読みExcel、開くとそのスニペット コードが表示されます。</span><span class="sxs-lookup"><span data-stu-id="67972-210">Script Lab reads the snippet ID from the Excel document and then displays that snippet code when it opens.</span></span> <span data-ttu-id="67972-211">カスタム プロパティを独自のアドイン (クエリ文字列、一時的な認証トークンなど) に送信する必要がある場合があります。アドイン **の起動時にカスタム プロパティ** を読み取る方法の詳細については、「永続化アドインの状態と設定」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="67972-211">You may need to send custom properties to your own add-in (such as a query string, or temporary authentication token.) See **Persisting add-in state and settings** for complete details on how to read custom properties when your add-in starts.</span></span>

### <a name="initialize-the-excel-document-with-data"></a><span data-ttu-id="67972-212">データを使用Excelドキュメントを初期化する</span><span class="sxs-lookup"><span data-stu-id="67972-212">Initialize the Excel document with data</span></span>

<span data-ttu-id="67972-213">通常、顧客が Web サイトから Excelドキュメントを開くと、そのドキュメントに Web サイトのデータが含まれると予想されます。</span><span class="sxs-lookup"><span data-stu-id="67972-213">Typically, when the customer opens up an Excel document from your web site, they expect the document to contain some data from the web site.</span></span> <span data-ttu-id="67972-214">ドキュメントにデータを書き込むには、いくつかの方法があります。</span><span class="sxs-lookup"><span data-stu-id="67972-214">There are a couple of ways to write data into the document.</span></span>

- <span data-ttu-id="67972-215">**OOXML SDK を使用してデータを書き込む**。</span><span class="sxs-lookup"><span data-stu-id="67972-215">**Use the OOXML SDK to write the data**.</span></span> <span data-ttu-id="67972-216">SDK を使用すると、ドキュメントに任意のデータを直接書き込みできます。</span><span class="sxs-lookup"><span data-stu-id="67972-216">You can use the SDK to directly write any data into the document.</span></span> <span data-ttu-id="67972-217">この方法は、ドキュメントを開いた瞬間にデータを使用できる場合に便利です。</span><span class="sxs-lookup"><span data-stu-id="67972-217">This approach is useful if you want the data to be available the instant the document is opened.</span></span>
- <span data-ttu-id="67972-218">**カスタム クエリ プロパティをアドインOffice渡します**。</span><span class="sxs-lookup"><span data-stu-id="67972-218">**Pass a custom query property to your Office Add-in**.</span></span> <span data-ttu-id="67972-219">ドキュメントを生成するときに、必要なすべてのデータを取得するクエリ文字列を含む Office アドインのカスタム プロパティを埋め込む必要があります。</span><span class="sxs-lookup"><span data-stu-id="67972-219">When you generate the document, you embed a custom property for the Office Add-in that contains a query string that retrieves all the required data.</span></span> <span data-ttu-id="67972-220">アドインが開くと、クエリを取得し、クエリを実行し、Office JS API を使用してクエリの結果をドキュメントに挿入します。</span><span class="sxs-lookup"><span data-stu-id="67972-220">When your add-in opens, it retrieves the query, runs the query, and uses the Office JS API to insert the result of the query into the document.</span></span>

### <a name="working-with-the-ooxml-sdk"></a><span data-ttu-id="67972-221">OOXML SDK の操作</span><span class="sxs-lookup"><span data-stu-id="67972-221">Working with the OOXML SDK</span></span>

<span data-ttu-id="67972-222">OOXML SDK は .NET に基づいて作成されます。</span><span class="sxs-lookup"><span data-stu-id="67972-222">The OOXML SDK is based on .NET.</span></span> <span data-ttu-id="67972-223">Web アプリケーションが .NET を使用しない場合は、OOXML を使用する別の方法を探す必要があります。</span><span class="sxs-lookup"><span data-stu-id="67972-223">If your web application does not .NET, you’ll need to look for an alternative way to work with OOXML.</span></span>

<span data-ttu-id="67972-224">Open XML SDK for JavaScript には、OOXML SDK の [JavaScript バージョンが用意されています](https://archive.codeplex.com/?p=openxmlsdkjs)。</span><span class="sxs-lookup"><span data-stu-id="67972-224">There is a JavaScript version of the OOXML SDK available at [Open XML SDK for JavaScript](https://archive.codeplex.com/?p=openxmlsdkjs).</span></span>

<span data-ttu-id="67972-225">OOXML コードを Azure 関数に配置して、.NET コードを他の Web アプリケーションから分離できます。</span><span class="sxs-lookup"><span data-stu-id="67972-225">You can place the OOXML code in an Azure function to separate the .NET code from the rest of your web application.</span></span> <span data-ttu-id="67972-226">次に、Web アプリケーションから Azure 関数 (Excelドキュメントを生成する) を呼び出します。</span><span class="sxs-lookup"><span data-stu-id="67972-226">Then call the Azure function (to generate the Excel document) from your Web application.</span></span> <span data-ttu-id="67972-227">Azure 関数の詳細については、「Azure [Functions の概要」を参照してください](/azure/azure-functions/functions-overview)。</span><span class="sxs-lookup"><span data-stu-id="67972-227">For more information on Azure functions, see [An introduction to Azure Functions](/azure/azure-functions/functions-overview).</span></span>

### <a name="use-single-sign-on"></a><span data-ttu-id="67972-228">シングル サインオンの使用</span><span class="sxs-lookup"><span data-stu-id="67972-228">Use single sign-on</span></span>

<span data-ttu-id="67972-229">認証を簡略化するために、アドインでシングル サインオンを実装することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="67972-229">To simplify authentication, we recommend your add-in implements single sign-on.</span></span> <span data-ttu-id="67972-230">詳細については、「Enable [single sign-on for Office アドイン」を参照してください。](../develop/sso-in-office-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="67972-230">For more information, see [Enable single sign-on for Office Add-ins](../develop/sso-in-office-add-ins.md)</span></span>

## <a name="see-also"></a><span data-ttu-id="67972-231">関連項目</span><span class="sxs-lookup"><span data-stu-id="67972-231">See also</span></span>

- [<span data-ttu-id="67972-232">Open XML SDK 2.5 for Office</span><span class="sxs-lookup"><span data-stu-id="67972-232">Welcome to the Open XML SDK 2.5 for Office</span></span>](/office/open-xml/open-xml-sdk)
- [<span data-ttu-id="67972-233">ドキュメントで作業ウィンドウを自動的に開く</span><span class="sxs-lookup"><span data-stu-id="67972-233">Automatically open a task pane with a document</span></span>](../develop/automatically-open-a-task-pane-with-a-document.md)
- [<span data-ttu-id="67972-234">アドインの状態および設定を保持する</span><span class="sxs-lookup"><span data-stu-id="67972-234">Persisting add-in state and settings</span></span>](../develop/persisting-add-in-state-and-settings.md)
- [<span data-ttu-id="67972-235">ファイル名を指定してスプレッドシート ドキュメントを作成する</span><span class="sxs-lookup"><span data-stu-id="67972-235">Create a spreadsheet document by providing a file name</span></span>](/office/open-xml/how-to-create-a-spreadsheet-document-by-providing-a-file-name)