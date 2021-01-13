---
title: Web ページから Excel を開き、アドインOffice埋め込む
description: Web ページから Excel を開き、アドインOffice埋め込む。
ms.date: 09/15/2020
localization_priority: Normal
ms.openlocfilehash: a88cc647fc1dba8ab6e6ddc0b504aab96517026a
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839867"
---
# <a name="open-excel-from-your-web-page-and-embed-your-office-add-in"></a><span data-ttu-id="6518b-103">Web ページから Excel を開き、アドインOffice埋め込む</span><span class="sxs-lookup"><span data-stu-id="6518b-103">Open Excel from your web page and embed your Office Add-in</span></span>

:::image type="content" source="../images/pnp-open-in-excel.png" alt-text="アドインを埋め込み、自動開きにした新しい Excel ドキュメントを開く Web ページ上の Excel ボタンの画像。":::

<span data-ttu-id="6518b-105">SaaS Web アプリケーションを拡張して、顧客が Web ページから Microsoft Excel に直接データを開くことができる。</span><span class="sxs-lookup"><span data-stu-id="6518b-105">Extend your SaaS web application so that your customers can open their data from a web page directly to Microsoft Excel.</span></span> <span data-ttu-id="6518b-106">一般的なシナリオは、顧客が Web アプリケーションのデータを操作することです。</span><span class="sxs-lookup"><span data-stu-id="6518b-106">A common scenario is that customers will be working with data in your web application.</span></span> <span data-ttu-id="6518b-107">次に、データを Excel ドキュメントにコピーします。</span><span class="sxs-lookup"><span data-stu-id="6518b-107">Then they’ll want to copy the data into an Excel document.</span></span> <span data-ttu-id="6518b-108">たとえば、Excel を使用して追加の分析を実行できます。</span><span class="sxs-lookup"><span data-stu-id="6518b-108">For example, they may want to perform additional analysis using Excel.</span></span> <span data-ttu-id="6518b-109">通常、顧客はデータを .csv ファイルなどのファイルにエクスポートし、そのデータを Excel にインポートする必要があります。</span><span class="sxs-lookup"><span data-stu-id="6518b-109">Typically, the customer is required to export the data to a file, such as a .csv file, and then import that data into Excel.</span></span> <span data-ttu-id="6518b-110">また、ドキュメントにアドインOffice手動で追加する必要があります。</span><span class="sxs-lookup"><span data-stu-id="6518b-110">They also have to manually add your Office Add-in to the document.</span></span>

<span data-ttu-id="6518b-111">Excel ドキュメントを生成して開く Web ページのボタンを 1 回クリックする手順の数を減らします。</span><span class="sxs-lookup"><span data-stu-id="6518b-111">Reduce the number of steps to a single button click on your web page that generates and opens the Excel document.</span></span> <span data-ttu-id="6518b-112">ドキュメント内にアドインOffice埋め込み、ドキュメントが開くと表示できます。</span><span class="sxs-lookup"><span data-stu-id="6518b-112">You can also embed your Office Add-in inside the document and display it when the document opens.</span></span> <span data-ttu-id="6518b-113">これにより、顧客は引き続きアプリケーション機能にアクセスできます。</span><span class="sxs-lookup"><span data-stu-id="6518b-113">This ensures the customer still has access to your application features.</span></span> <span data-ttu-id="6518b-114">ドキュメントが開くと、顧客が選択したデータとOfficeアドインは、引き続き作業できます。</span><span class="sxs-lookup"><span data-stu-id="6518b-114">When the document opens, the data the customer selected, and your Office Add-in is already available for them to continue working.</span></span>

<span data-ttu-id="6518b-115">この記事では、このシナリオを独自の SaaS Web アプリケーションに実装するためのコードと手法について説明します。</span><span class="sxs-lookup"><span data-stu-id="6518b-115">This article shows you code and techniques for implementing this scenario in your own SaaS web application.</span></span>

## <a name="create-a-new-excel-document-and-embed-an-office-add-in"></a><span data-ttu-id="6518b-116">新しい Excel ドキュメントを作成し、新Office埋め込む</span><span class="sxs-lookup"><span data-stu-id="6518b-116">Create a new Excel document and embed an Office Add-in</span></span>

<span data-ttu-id="6518b-117">最初に、Web ページから Excel ドキュメントを作成し、アドインをドキュメントに埋め込む方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="6518b-117">First, let’s learn how to create an Excel document from a web page, and embed an add-in into the document.</span></span> <span data-ttu-id="6518b-118">次 [Office OOXML Embed アドイン](https://github.com/OfficeDev/Office-OOXML-EmbedAddin) のコード サンプルは [、Script Lab](https://appsource.microsoft.com/product/office/wa104380862) アドインを新しいドキュメントに埋め込むOfficeしています。</span><span class="sxs-lookup"><span data-stu-id="6518b-118">The [Office OOXML Embed Add-in code sample](https://github.com/OfficeDev/Office-OOXML-EmbedAddin) shows how to embed the [Script Lab add-in](https://appsource.microsoft.com/product/office/wa104380862) into a new Office document.</span></span> <span data-ttu-id="6518b-119">このサンプルは、すべてのドキュメントOffice動作しますが、この記事では Excel スプレッドシートに焦点を当てるだけについて説明します。</span><span class="sxs-lookup"><span data-stu-id="6518b-119">Although the sample works with any Office document, we’ll just focus on Excel spreadsheets in this article.</span></span> <span data-ttu-id="6518b-120">次の手順を使用して、サンプルをビルドして実行します。</span><span class="sxs-lookup"><span data-stu-id="6518b-120">Use the following steps to build and run the sample.</span></span>

1. <span data-ttu-id="6518b-121">サンプル コードをコンピューター上  https://github.com/OfficeDev/Office-OOXML-EmbedAddin/archive/master.zip のフォルダーに抽出します。</span><span class="sxs-lookup"><span data-stu-id="6518b-121">Extract the sample code from  https://github.com/OfficeDev/Office-OOXML-EmbedAddin/archive/master.zip into a folder on your computer.</span></span>
2. <span data-ttu-id="6518b-122">サンプルをビルドして実行するには、readme の「プロジェクトを使用するには」セクションの手順に従います。</span><span class="sxs-lookup"><span data-stu-id="6518b-122">To build and run the sample, follow the steps in the **To use the project** section of the readme.</span></span>
3. <span data-ttu-id="6518b-123">サンプルを実行すると、次のスクリーン ショットのような Web ページが表示されます。</span><span class="sxs-lookup"><span data-stu-id="6518b-123">When you run the sample it will display a web page similar to the following screen shot.</span></span> <span data-ttu-id="6518b-124">Web ページを使用して、Script Lab を含む新しい Excel ドキュメントを作成します (開きます)。</span><span class="sxs-lookup"><span data-stu-id="6518b-124">Use the web page to create a new Excel document that contains Script Lab when it opens.</span></span>
:::image type="content" source="../images/embed-script-lab-sample-ui.png" alt-text="Excel ファイルを選択してスクリプト ラボ アドインを埋め込む目的で、埋め込みスクリプト ラボ サンプルに表示される Web ページのスクリーン ショット。":::

### <a name="how-the-sample-works"></a><span data-ttu-id="6518b-126">サンプルのしくみ</span><span class="sxs-lookup"><span data-stu-id="6518b-126">How the sample works</span></span>

<span data-ttu-id="6518b-127">サンプル コードでは、OOXML SDK を使用して、選択した Excel ドキュメントに Script Lab アドインを埋め込む方法を示します。</span><span class="sxs-lookup"><span data-stu-id="6518b-127">The sample code uses the OOXML SDK to embed the Script Lab add-in to the Excel document that you choose.</span></span> <span data-ttu-id="6518b-128">次の情報は、readme ファイルの [コード [**について** ]](https://github.com/OfficeDev/Office-OOXML-EmbedAddin/blob/master/README.md) セクションから取得されます。</span><span class="sxs-lookup"><span data-stu-id="6518b-128">The following Information is taken from the [**About the code** section](https://github.com/OfficeDev/Office-OOXML-EmbedAddin/blob/master/README.md) in the readme file.</span></span>

<span data-ttu-id="6518b-129">次の **ファイルHome.aspx.cs。**</span><span class="sxs-lookup"><span data-stu-id="6518b-129">The file **Home.aspx.cs**:</span></span>

- <span data-ttu-id="6518b-130">ボタン イベント ハンドラーと基本的な UI 操作を提供します。</span><span class="sxs-lookup"><span data-stu-id="6518b-130">Provides the button event handlers and basic UI manipulation.</span></span>
- <span data-ttu-id="6518b-131">標準的なASP.NETを使用して、ファイルをアップロードおよびダウンロードします。</span><span class="sxs-lookup"><span data-stu-id="6518b-131">Uses standard ASP.NET techniques to upload and download the file.</span></span>
- <span data-ttu-id="6518b-132">アップロードしたファイルのファイル名の拡張子 (xlsx、docx、pptx) を使用して、ファイルの種類を特定します。</span><span class="sxs-lookup"><span data-stu-id="6518b-132">Uses the file name extension of the uploaded file (xlsx, docx, or pptx) to determine the type of file.</span></span> <span data-ttu-id="6518b-133">Open XML SDK には通常、ファイルの種類ごとに異なる API が含まれるため、最初にこれを行う必要があります。</span><span class="sxs-lookup"><span data-stu-id="6518b-133">This needs to be done at the outset because the Open XML SDK generally has distinct APIs for each type of file.</span></span>
- <span data-ttu-id="6518b-134">**OOXMLHelper** を呼び出してファイルを検証し **、AddInEmbedder** を呼び出して Script Lab をファイルに埋め込み、自動的に開く設定を行います。</span><span class="sxs-lookup"><span data-stu-id="6518b-134">Calls into the **OOXMLHelper** to validate the file and calls into the **AddInEmbedder** to embed Script Lab in the file and set to automatically open.</span></span>

<span data-ttu-id="6518b-135">次の **ファイルAddInEmbedder.cs。**</span><span class="sxs-lookup"><span data-stu-id="6518b-135">The file **AddInEmbedder.cs**:</span></span>

- <span data-ttu-id="6518b-136">主要なビジネス ロジックを提供します。このサンプルでは、Script Lab を埋め込むメソッドです。</span><span class="sxs-lookup"><span data-stu-id="6518b-136">Provides the main business logic, which in this sample is a method that embeds Script Lab.</span></span>
- <span data-ttu-id="6518b-137">ファイルの種類に基づいて OOXML ヘルパーを呼び出します。</span><span class="sxs-lookup"><span data-stu-id="6518b-137">Makes calls into the OOXML helper based on the type of the file.</span></span>

<span data-ttu-id="6518b-138">次の **ファイルOOXMLHelper.cs。**</span><span class="sxs-lookup"><span data-stu-id="6518b-138">The file **OOXMLHelper.cs**:</span></span>

- <span data-ttu-id="6518b-139">すべての詳細な OOXML 操作を提供します。</span><span class="sxs-lookup"><span data-stu-id="6518b-139">Provides all the detailed OOXML manipulation.</span></span>
- <span data-ttu-id="6518b-140">ファイルに対して **Document.Open** メソッドを呼びOfficeファイルを検証するための標準的な手法を使用します。</span><span class="sxs-lookup"><span data-stu-id="6518b-140">Uses a standard technique for validating the Office file, which is simply to call the **Document.Open** method on it.</span></span> <span data-ttu-id="6518b-141">ファイルが無効な場合、メソッドは例外をスローします。</span><span class="sxs-lookup"><span data-stu-id="6518b-141">If the file is invalid, the method throws an exception.</span></span>
- <span data-ttu-id="6518b-142">主に Open XML 2.5 SDK Productivity Tools によって生成されたコードが含まれています。このコードは [、Open XML 2.5 SDK](/office/open-xml/open-xml-sdk)のリンクから参照できます。</span><span class="sxs-lookup"><span data-stu-id="6518b-142">Contains mainly code that was generated by the Open XML 2.5 SDK Productivity Tools which are available at the link for the [Open XML 2.5 SDK](/office/open-xml/open-xml-sdk).</span></span>

<span data-ttu-id="6518b-143">OOXMLHelper.cs ファイルの **GenerateWebExtensionPart1Content** メソッドは、Microsoft AppSource の Script Lab の ID への参照を設定します。 </span><span class="sxs-lookup"><span data-stu-id="6518b-143">The **GenerateWebExtensionPart1Content** method in the **OOXMLHelper.cs** file sets the reference to the ID of Script Lab in Microsoft AppSource:</span></span>

```csharp
We.WebExtensionStoreReference webExtensionStoreReference1 = new We.WebExtensionStoreReference() { Id = "wa104380862", Version = "1.1.0.0", Store = "en-US", StoreType = "OMEX" };
```

- <span data-ttu-id="6518b-144">StoreType **値** は、Microsoft AppSource のエイリアスである "OMEX" です。</span><span class="sxs-lookup"><span data-stu-id="6518b-144">The **StoreType** value is "OMEX", an alias for Microsoft AppSource.</span></span>
- <span data-ttu-id="6518b-145">Store **の** 値は、Script Lab の Microsoft AppSource カルチャ セクションにある "en-US" です。</span><span class="sxs-lookup"><span data-stu-id="6518b-145">The **Store** value is "en-US" found in the Microsoft AppSource culture section for Script Lab.</span></span>
- <span data-ttu-id="6518b-146">Id **の** 値は、Script Lab の Microsoft AppSource アセット ID です。</span><span class="sxs-lookup"><span data-stu-id="6518b-146">The **Id** value is the Microsoft AppSource asset ID for Script Lab.</span></span>

<span data-ttu-id="6518b-147">自動開き用にファイル共有カタログからアドインをセットアップする場合は、次の値を使用します。</span><span class="sxs-lookup"><span data-stu-id="6518b-147">If you are setting up an add-in from a file share catalog for auto-open, you will use different values:</span></span>

<span data-ttu-id="6518b-148">**StoreType の値** は "FileSystem" です。</span><span class="sxs-lookup"><span data-stu-id="6518b-148">The **StoreType** value is "FileSystem".</span></span>

- <span data-ttu-id="6518b-149">Store **の** 値は、ネットワーク共有の URL です。たとえば \\ \\ 、「MyComputer \\ MySharedFolder」とします。</span><span class="sxs-lookup"><span data-stu-id="6518b-149">The **Store** value is the URL of the network share; for example, "\\\\MyComputer\\MySharedFolder".</span></span> <span data-ttu-id="6518b-150">これは、セキュリティ センターで共有の信頼済みカタログ アドレスとして表示される正確な URL Office必要があります。</span><span class="sxs-lookup"><span data-stu-id="6518b-150">This should be the exact URL that appears as the share's Trusted Catalog Address in the Office Trust Center.</span></span>
- <span data-ttu-id="6518b-151">Id **値** は、アドイン マニフェストのアプリ ID です。</span><span class="sxs-lookup"><span data-stu-id="6518b-151">The **Id** value is the app ID in the add-ins manifest.</span></span>
> [!NOTE]
> <span data-ttu-id="6518b-152">これらの属性の代替値の詳細については、「ドキュメントで作業ウィンドウを自動的に [開く」を参照してください](../develop/automatically-open-a-task-pane-with-a-document.md)。</span><span class="sxs-lookup"><span data-stu-id="6518b-152">For more information about alternative values for these attributes, see [Automatically open a task pane with a document](../develop/automatically-open-a-task-pane-with-a-document.md).</span></span>

## <a name="use-the-fluent-ui"></a><span data-ttu-id="6518b-153">Fluent UI を使用する</span><span class="sxs-lookup"><span data-stu-id="6518b-153">Use the Fluent UI</span></span>

:::image type="content" source="../images/fluent-ui-wxp.png" alt-text="Word、Excel、PowerPoint の Fluent UI アイコン。":::

<span data-ttu-id="6518b-155">ベスト プラクティスは、Fluent UI を使用して、ユーザーが Microsoft 製品間を移行する場合に役立ちます。</span><span class="sxs-lookup"><span data-stu-id="6518b-155">A best practice is to use the Fluent UI to help your users transition between Microsoft products.</span></span> <span data-ttu-id="6518b-156">Web ページから起動するOfficeを示Officeアイコンを常に使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="6518b-156">You should always use an Office icon to indicate which Office application will be launched from your web page.</span></span> <span data-ttu-id="6518b-157">サンプル コードを変更して Excel アイコンを使用し、Excel アプリケーションを起動します。</span><span class="sxs-lookup"><span data-stu-id="6518b-157">Let’s modify the sample code to use the Excel icon to indicate that it launches the Excel application.</span></span>

1. <span data-ttu-id="6518b-158">サンプルを次のVisual Studio。</span><span class="sxs-lookup"><span data-stu-id="6518b-158">Open the sample in Visual Studio.</span></span>
1. <span data-ttu-id="6518b-159">**Home.aspx ページを開** きます。</span><span class="sxs-lookup"><span data-stu-id="6518b-159">Open the **Home.aspx** page.</span></span>
1. <span data-ttu-id="6518b-160">フォームのダウンロード ボタンである次のコードを検索します。</span><span class="sxs-lookup"><span data-stu-id="6518b-160">Find following code that is the download button on the form:</span></span>
    ```html
    <asp:Button ID="btnDownload" runat="server" Text="Download" OnClick="btnDownload_Click" /> 
    ```
1. <span data-ttu-id="6518b-161">ボタンのコードを次のイメージ タグに置き換えます。</span><span class="sxs-lookup"><span data-stu-id="6518b-161">Replace the button code with the following image tag.</span></span>
    ```html
    <asp:Image  src="https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/excel_48x1.svg" width="48" height="48" ID="btnDownload" runat="server" OnClick="btnDownload_Click" AlternateText="Open in Microsoft Excel" role="button" ImageUrl=""/>  
    ```
1. <span data-ttu-id="6518b-162">**F5 キーを押** します (またはデバッグ **>を開始します**)。</span><span class="sxs-lookup"><span data-stu-id="6518b-162">Press **F5** (or **Debug > Start Debugging**).</span></span> <span data-ttu-id="6518b-163">ホーム ページが読み込まれるとアイコンが表示されます。</span><span class="sxs-lookup"><span data-stu-id="6518b-163">You'll see the icon appear when the home page loads.</span></span>

<span data-ttu-id="6518b-164">詳しくは、Fluent UI [Officeのブランド](https://developer.microsoft.com/fluentui#/styles/web/office-brand-icons) アイコンの詳細をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="6518b-164">For more information, see [Office Brand Icons](https://developer.microsoft.com/fluentui#/styles/web/office-brand-icons) on the Fluent UI developer portal.</span></span>  

## <a name="upload-the-excel-document-to-microsoft-onedrive"></a><span data-ttu-id="6518b-165">Excel ドキュメントを Microsoft OneDrive にアップロードする</span><span class="sxs-lookup"><span data-stu-id="6518b-165">Upload the Excel document to Microsoft OneDrive</span></span>

<span data-ttu-id="6518b-166">顧客が OneDrive を使用している場合は、OneDrive に新しいドキュメントをアップロードすることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="6518b-166">We recommend uploading new documents to OneDrive if your customer uses OneDrive.</span></span> <span data-ttu-id="6518b-167">これにより、ドキュメントの検索と作業が容易になります。</span><span class="sxs-lookup"><span data-stu-id="6518b-167">This makes it easier for them to find and work with the documents.</span></span> <span data-ttu-id="6518b-168">新しいコード サンプルを作成し、Microsoft Graph SDK を使用して新しい Excel ドキュメントを OneDrive にアップロードする方法を確認しましょう。</span><span class="sxs-lookup"><span data-stu-id="6518b-168">Let’s create a new code sample and see how you can use the Microsoft Graph SDK to upload a new Excel document to OneDrive.</span></span>

### <a name="use-a-quick-start-to-build-a-new-microsoft-graph-web-application"></a><span data-ttu-id="6518b-169">クイック スタートを使用して新しい Microsoft Graph Web アプリケーションを構築する</span><span class="sxs-lookup"><span data-stu-id="6518b-169">Use a quick-start to build a new Microsoft Graph web application</span></span>

1. <span data-ttu-id="6518b-170">手順に [https://developer.microsoft.com/graph/quick-start](https://developer.microsoft.com/graph/quick-start) 従って、365 サービスとやり取りするクイック スタート コード サンプルを作成Office開きます。</span><span class="sxs-lookup"><span data-stu-id="6518b-170">Go to [https://developer.microsoft.com/graph/quick-start](https://developer.microsoft.com/graph/quick-start) and follow the steps to create and open a quick start code sample that interacts with Office 365 services.</span></span>
1. <span data-ttu-id="6518b-171">手順 **1: 言語またはプラットフォームを選択し、MVC** ASP.NET **します**。</span><span class="sxs-lookup"><span data-stu-id="6518b-171">In **step 1: Pick you language or platform**, choose **ASP.NET MVC**.</span></span> <span data-ttu-id="6518b-172">この手順の手順では MVC オプションASP.NET使用しますが、この手順は任意の言語またはプラットフォームに適用されるパターンに従います。</span><span class="sxs-lookup"><span data-stu-id="6518b-172">Although the steps in this procedure use the ASP.NET MVC option, the steps follow a pattern that apply to any language or platform.</span></span>
1. <span data-ttu-id="6518b-173">手順 **2: アプリ ID とシークレットを取得** し、[アプリ ID とシークレットの **取得] を選択します**。</span><span class="sxs-lookup"><span data-stu-id="6518b-173">In **step 2: Get an app ID and secret**, choose **Get an app ID and secret**.</span></span>
1. <span data-ttu-id="6518b-174">Microsoft 365 アカウントにサインインします。</span><span class="sxs-lookup"><span data-stu-id="6518b-174">Sign in to your Microsoft 365 account.</span></span>  
1. <span data-ttu-id="6518b-175">アプリ シークレット **Web ページで** 、アプリ シークレットを後で取得して使用できるファイルの場所に保存します。</span><span class="sxs-lookup"><span data-stu-id="6518b-175">On the **Please save your app secret** web page, save the app secret to a file location where you can retrieve and use it later.</span></span>
1. <span data-ttu-id="6518b-176">Choose **Got it, take me back to the quick start**.</span><span class="sxs-lookup"><span data-stu-id="6518b-176">Choose **Got it, take me back to the quick start**.</span></span>
1. <span data-ttu-id="6518b-177">手順 **2: 登録に成功しました。**</span><span class="sxs-lookup"><span data-stu-id="6518b-177">In **step 2: Registration Successful!**</span></span> <span data-ttu-id="6518b-178">生成されたアプリ シークレットを入力します。</span><span class="sxs-lookup"><span data-stu-id="6518b-178">Enter the generated app secret.</span></span>
1. <span data-ttu-id="6518b-179">手順 **3: コーディングを開始し**、[SDK ベースのコード サンプルのダウンロード **] を選択します**。</span><span class="sxs-lookup"><span data-stu-id="6518b-179">In **step 3: Start coding**, choose **Download the SDK-based code sample**.</span></span>
1. <span data-ttu-id="6518b-180">ダウンロード zip フォルダーをローカル フォルダーに展開します。</span><span class="sxs-lookup"><span data-stu-id="6518b-180">Extract the download zip folder into a local folder.</span></span>  
1. <span data-ttu-id="6518b-181">Visual Studio 2019 で graph-tutorial.sln ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="6518b-181">Open the graph-tutorial.sln file in Visual Studio 2019.</span></span>
1. <span data-ttu-id="6518b-182">ソリューションをビルドして実行し、正常に動作しているのを確認します。</span><span class="sxs-lookup"><span data-stu-id="6518b-182">Build and run the solution and confirm it is working correctly.</span></span> <span data-ttu-id="6518b-183">予定表 Web ページを使用して Microsoft 365 の予定表を表示できる必要があります。</span><span class="sxs-lookup"><span data-stu-id="6518b-183">You should be able to use the calendar web page to view your Microsoft 365 calendar.</span></span>

### <a name="upload-a-file-to-onedrive"></a><span data-ttu-id="6518b-184">OneDrive にファイルをアップロードする</span><span class="sxs-lookup"><span data-stu-id="6518b-184">Upload a file to OneDrive</span></span>

1. <span data-ttu-id="6518b-185">Visual Studio 2019 で **graph-tutorial.sln** ソリューションを開き、PrivateSettings.config **します。**</span><span class="sxs-lookup"><span data-stu-id="6518b-185">Open the **graph-tutorial.sln** solution in Visual Studio 2019, and open the **PrivateSettings.config** file.</span></span>
1. <span data-ttu-id="6518b-186">次のコードのように、新しいスコープ **Files.ReadWrite** を   **ida:AppScopes** キーに追加します。</span><span class="sxs-lookup"><span data-stu-id="6518b-186">Add a new scope **Files.ReadWrite** to the **ida:AppScopes** key so that it looks like the following code:</span></span>
    ```xml
    <add key="ida:AppScopes" value="User.Read Calendars.Read Files.ReadWrite " />
    ```
1. <span data-ttu-id="6518b-187">**Index.cshtml ファイルを開** きます。</span><span class="sxs-lookup"><span data-stu-id="6518b-187">Open the **Index.cshtml** file.</span></span>
1. <span data-ttu-id="6518b-188">次の ActionLink コードを挿入して、ファイルを OneDrive にアップロードするボタンを作成します。</span><span class="sxs-lookup"><span data-stu-id="6518b-188">Insert the following ActionLink code to create a button to upload a file to OneDrive.</span></span>
    ```razor
    @if (Request.IsAuthenticated)
    {
        <h4>Welcome @ViewBag.User.DisplayName!</h4>
        <p>Use the navigation bar at the top of the page to get started.</p>
        @Html.ActionLink("Click here to create a new file on OneDrive", "CreateOneDriveFile", "Home", new { area = "" }, new { @class = "btn btn-primary btn-large" })
    }
    ```
1. <span data-ttu-id="6518b-189">**HomeController.cs** ファイルを開きます。</span><span class="sxs-lookup"><span data-stu-id="6518b-189">Open the **HomeController.cs** file.</span></span>
1. <span data-ttu-id="6518b-190">アクション リンクからの要求を処理する次のコードを挿入します。</span><span class="sxs-lookup"><span data-stu-id="6518b-190">Insert the following code to handle the request from the action link.</span></span>
    ```csharp
    public void CreateOneDriveFile()
        {
            using (var stream = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes("The contents of the file goes here.")))
            {
                var client = graph_tutorial.Helpers.GraphHelper.UploadFile("/test.txt", stream);
            }
        }
    ```
1. <span data-ttu-id="6518b-191">ファイルを **GraphHelper.cs** します。</span><span class="sxs-lookup"><span data-stu-id="6518b-191">Open the **GraphHelper.cs** file.</span></span>
1. <span data-ttu-id="6518b-192">次のコードを挿入して Microsoft Graph API を呼び出し、OneDrive に新しいファイルを作成します。</span><span class="sxs-lookup"><span data-stu-id="6518b-192">Insert the following code to call the Microsoft Graph API to create a new file on OneDrive.</span></span>
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
1. <span data-ttu-id="6518b-193">**F5 キーを押** します (またはデバッグ **>を開始します**)。</span><span class="sxs-lookup"><span data-stu-id="6518b-193">Press **F5** (or **Debug > Start Debugging**).</span></span> <span data-ttu-id="6518b-194">Web アプリケーションが起動します。</span><span class="sxs-lookup"><span data-stu-id="6518b-194">The web application will start.</span></span>
1. <span data-ttu-id="6518b-195">Choose **Click here to sign in,** and sign in.</span><span class="sxs-lookup"><span data-stu-id="6518b-195">Choose **Click here to sign in**, and sign in.</span></span>
1. <span data-ttu-id="6518b-196">Choose **Click here to create a new file on OneDrive**.</span><span class="sxs-lookup"><span data-stu-id="6518b-196">Choose **Click here to create a new file on OneDrive**.</span></span>
1. <span data-ttu-id="6518b-197">新しいブラウザー タブを開き、OneDrive アカウントにサインインします。</span><span class="sxs-lookup"><span data-stu-id="6518b-197">Open a new browser tab and sign in to your OneDrive account.</span></span> <span data-ttu-id="6518b-198">ルート フォルダーにtest.txtファイルが表示されます。</span><span class="sxs-lookup"><span data-stu-id="6518b-198">You'll see the test.txt file in the root folder.</span></span>

<span data-ttu-id="6518b-199">OneDrive にファイルをアップロードする方法を学んだので、このコードを再利用して、作成した Excel ドキュメントをアップロードできます。</span><span class="sxs-lookup"><span data-stu-id="6518b-199">Now that you've learned how to upload a file to OneDrive, you can reuse this code to upload any Excel document that you create.</span></span>

## <a name="additional-considerations-for-your-solution"></a><span data-ttu-id="6518b-200">ソリューションに関するその他の考慮事項</span><span class="sxs-lookup"><span data-stu-id="6518b-200">Additional considerations for your solution</span></span>

<span data-ttu-id="6518b-201">テクノロジとアプローチの点では、すべてのユーザーのソリューションが異なります。</span><span class="sxs-lookup"><span data-stu-id="6518b-201">Everyone’s solution is different in terms of technologies and approaches.</span></span> <span data-ttu-id="6518b-202">次の考慮事項は、ソリューションを変更してドキュメントを開き、アドインを埋め込むOfficeに役立ちます。</span><span class="sxs-lookup"><span data-stu-id="6518b-202">The following considerations will help you plan how to modify your solution to open documents and embed your Office Add-in.</span></span>

### <a name="create-a-new-excel-spreadsheet-from-the-web-page"></a><span data-ttu-id="6518b-203">Web ページから新しい Excel スプレッドシートを作成する</span><span class="sxs-lookup"><span data-stu-id="6518b-203">Create a new Excel spreadsheet from the web page</span></span>

<span data-ttu-id="6518b-204">このサンプルでは、既存の Excel ドキュメントを変更します。</span><span class="sxs-lookup"><span data-stu-id="6518b-204">The sample modifies an existing Excel document.</span></span> <span data-ttu-id="6518b-205">より一般的なシナリオとして、Web ページから新しい Excel スプレッドシートを作成します。</span><span class="sxs-lookup"><span data-stu-id="6518b-205">A more common scenario is that you’ll create a new Excel spreadsheet from your web page.</span></span> <span data-ttu-id="6518b-206">新しいスプレッドシートを作成する方法の詳細については、「ファイル名を指定してスプレッドシート ドキュメントを作成する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6518b-206">You can find additional details on how to create a new spreadsheet in **Create a spreadsheet document** by providing a file name.</span></span> <span data-ttu-id="6518b-207">この記事では、ファイルをローカルに作成する方法を示しますが、SpreadsheetDocument.Create メソッドのオーバーロードを使用して、ファイルをストリームで作成することもできます。</span><span class="sxs-lookup"><span data-stu-id="6518b-207">This article shows how to create the file locally, but you can also create the file in a stream by using an overload on the SpreadsheetDocument.Create method.</span></span>

### <a name="read-custom-properties-when-your-add-in-starts"></a><span data-ttu-id="6518b-208">アドインの起動時にカスタム プロパティを読み取る</span><span class="sxs-lookup"><span data-stu-id="6518b-208">Read custom properties when your add-in starts</span></span>

<span data-ttu-id="6518b-209">コード サンプルでは、OOXML SDK を使用して新しい Excel ドキュメントにスニペット ID を格納します。</span><span class="sxs-lookup"><span data-stu-id="6518b-209">The code sample stores a snippet ID in the new Excel document using the OOXML SDK.</span></span> <span data-ttu-id="6518b-210">Script Lab は、Excel ドキュメントからスニペット ID を読み取り、そのスニペット コードを開くと表示します。</span><span class="sxs-lookup"><span data-stu-id="6518b-210">Script Lab reads the snippet ID from the Excel document and then displays that snippet code when it opens.</span></span> <span data-ttu-id="6518b-211">カスタム プロパティを独自のアドイン (クエリ文字列、一時認証トークンなど) に送信する必要がある場合があります。アドイン **の起動時にカスタム プロパティを読** み取る方法の詳細については、「アドインの状態と設定を保持する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6518b-211">You may need to send custom properties to your own add-in (such as a query string, or temporary authentication token.) See **Persisting add-in state and settings** for complete details on how to read custom properties when your add-in starts.</span></span>

### <a name="initialize-the-excel-document-with-data"></a><span data-ttu-id="6518b-212">データを使用して Excel ドキュメントを初期化する</span><span class="sxs-lookup"><span data-stu-id="6518b-212">Initialize the Excel document with data</span></span>

<span data-ttu-id="6518b-213">通常、顧客が Web サイトから Excel ドキュメントを開くと、そのドキュメントには Web サイトのデータが含まれると予想されます。</span><span class="sxs-lookup"><span data-stu-id="6518b-213">Typically, when the customer opens up an Excel document from your web site, they expect the document to contain some data from the web site.</span></span> <span data-ttu-id="6518b-214">ドキュメントにデータを書き込むには、いくつかの方法があります。</span><span class="sxs-lookup"><span data-stu-id="6518b-214">There are a couple of ways to write data into the document.</span></span>

- <span data-ttu-id="6518b-215">**OOXML SDK を使用してデータを書き込む**。</span><span class="sxs-lookup"><span data-stu-id="6518b-215">**Use the OOXML SDK to write the data**.</span></span> <span data-ttu-id="6518b-216">SDK を使用して、任意のデータをドキュメントに直接書き込みできます。</span><span class="sxs-lookup"><span data-stu-id="6518b-216">You can use the SDK to directly write any data into the document.</span></span> <span data-ttu-id="6518b-217">この方法は、ドキュメントを開いた時点でデータを使用する場合に便利です。</span><span class="sxs-lookup"><span data-stu-id="6518b-217">This approach is useful if you want the data to be available the instant the document is opened.</span></span>
- <span data-ttu-id="6518b-218">**カスタム クエリ プロパティをアドインにOffice渡します**。</span><span class="sxs-lookup"><span data-stu-id="6518b-218">**Pass a custom query property to your Office add-in**.</span></span> <span data-ttu-id="6518b-219">ドキュメントを生成するときに、必要なすべてのデータを取得するクエリ文字列を含む Office アドインのカスタム プロパティを埋め込む必要があります。</span><span class="sxs-lookup"><span data-stu-id="6518b-219">When you generate the document, you embed a custom property for the Office add-in that contains a query string that retrieves all the required data.</span></span> <span data-ttu-id="6518b-220">アドインが開くと、クエリを取得し、クエリを実行し、Office JS API を使用してクエリの結果をドキュメントに挿入します。</span><span class="sxs-lookup"><span data-stu-id="6518b-220">When your add-in opens, it retrieves the query, runs the query, and uses the Office JS API to insert the result of the query into the document.</span></span>

### <a name="working-with-the-ooxml-sdk"></a><span data-ttu-id="6518b-221">OOXML SDK の操作</span><span class="sxs-lookup"><span data-stu-id="6518b-221">Working with the OOXML SDK</span></span>

<span data-ttu-id="6518b-222">OOXML SDK は .NET に基づいて作成されています。</span><span class="sxs-lookup"><span data-stu-id="6518b-222">The OOXML SDK is based on .NET.</span></span> <span data-ttu-id="6518b-223">Web アプリケーションが .NET ではない場合は、OOXML を使用する別の方法を探す必要があります。</span><span class="sxs-lookup"><span data-stu-id="6518b-223">If your web application does not .NET, you’ll need to look for an alternative way to work with OOXML.</span></span>

<span data-ttu-id="6518b-224">[Open XML SDK for JavaScript には、OOXML SDK の JavaScript バージョンが用意されています](https://archive.codeplex.com/?p=openxmlsdkjs)。</span><span class="sxs-lookup"><span data-stu-id="6518b-224">There is a JavaScript version of the OOXML SDK available at [Open XML SDK for JavaScript](https://archive.codeplex.com/?p=openxmlsdkjs).</span></span>

<span data-ttu-id="6518b-225">OOXML コードを Azure 関数に配置して、.NET コードを Web アプリケーションの他の部分から分離できます。</span><span class="sxs-lookup"><span data-stu-id="6518b-225">You can place the OOXML code in an Azure function to separate the .NET code from the rest of your web application.</span></span> <span data-ttu-id="6518b-226">次に、Web アプリケーションから Azure 関数を呼び出します (Excel ドキュメントを生成します)。</span><span class="sxs-lookup"><span data-stu-id="6518b-226">Then call the Azure function (to generate the Excel document) from your Web application.</span></span> <span data-ttu-id="6518b-227">Azure 関数について詳しくは、「Azure 関数の概要 [」をご覧ください](/azure/azure-functions/functions-overview)。</span><span class="sxs-lookup"><span data-stu-id="6518b-227">For more information on Azure functions, see [An introduction to Azure Functions](/azure/azure-functions/functions-overview).</span></span>

### <a name="use-single-sign-on"></a><span data-ttu-id="6518b-228">シングル サインオンを使用する</span><span class="sxs-lookup"><span data-stu-id="6518b-228">Use single sign-on</span></span>

<span data-ttu-id="6518b-229">認証を簡略化するために、アドインにシングル サインオンを実装することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="6518b-229">To simplify authentication, we recommend your add-in implements single sign-on.</span></span> <span data-ttu-id="6518b-230">詳細については、「アドインのシングル サインオンを有効にする [Office参照してください。](../develop/sso-in-office-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="6518b-230">For more information, see [Enable single sign-on for Office Add-ins](../develop/sso-in-office-add-ins.md)</span></span>

## <a name="see-also"></a><span data-ttu-id="6518b-231">関連項目</span><span class="sxs-lookup"><span data-stu-id="6518b-231">See also</span></span>

- [<span data-ttu-id="6518b-232">Welcome to the Open XML SDK 2.5 for Office</span><span class="sxs-lookup"><span data-stu-id="6518b-232">Welcome to the Open XML SDK 2.5 for Office</span></span>](/office/open-xml/open-xml-sdk)
- [<span data-ttu-id="6518b-233">ドキュメントで作業ウィンドウを自動的に開く</span><span class="sxs-lookup"><span data-stu-id="6518b-233">Automatically open a task pane with a document</span></span>](../develop/automatically-open-a-task-pane-with-a-document.md)
- [<span data-ttu-id="6518b-234">アドインの状態および設定を保持する</span><span class="sxs-lookup"><span data-stu-id="6518b-234">Persisting add-in state and settings</span></span>](../develop/persisting-add-in-state-and-settings.md)
- [<span data-ttu-id="6518b-235">ファイル名を指定してスプレッドシート ドキュメントを作成する</span><span class="sxs-lookup"><span data-stu-id="6518b-235">Create a spreadsheet document by providing a file name</span></span>](/office/open-xml/how-to-create-a-spreadsheet-document-by-providing-a-file-name)