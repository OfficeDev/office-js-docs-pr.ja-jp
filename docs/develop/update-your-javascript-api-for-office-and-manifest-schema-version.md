---
title: Office ライブラリ用の最新 JavaScript API、およびバージョン 1.1 のアドイン マニフェスト スキーマへの更新
description: Office アドイン プロジェクトの JavaScript ファイル (Office.js ファイルとアプリに固有の .js ファイル) とアドイン マニフェスト検証ファイルをバージョン 1.1 に更新します。
ms.date: 12/12/2018
localization_priority: Normal
ms.openlocfilehash: 20c6c6362aa09926e967e52edfe6be69a09edb18
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/23/2019
ms.locfileid: "29387723"
---
# <a name="update-to-the-latest-javascript-api-for-office-library-and-version-11-add-in-manifest-schema"></a><span data-ttu-id="f2137-103">Office ライブラリ用の最新 JavaScript API、およびバージョン 1.1 のアドイン マニフェスト スキーマへの更新</span><span class="sxs-lookup"><span data-stu-id="f2137-103">Update to the latest JavaScript API for Office library and version 1.1 add-in manifest schema</span></span>

<span data-ttu-id="f2137-104">この記事では、Office アドイン プロジェクトに含まれる JavaScript ファイル (Office.js およびアプリに固有の .js ファイル) とアドイン マニフェスト検証ファイルをバージョン 1.1 に更新する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="f2137-104">This article describes how to update your JavaScript files (Office.js and app-specific .js files) and add-in manifest validation file in your Office Add-in project to version 1.1.</span></span>

> [!NOTE]
> <span data-ttu-id="f2137-105">Visual Studio 2017 で作成されたプロジェクトでは既にバージョン 1.1 が使用されています。</span><span class="sxs-lookup"><span data-stu-id="f2137-105">Projects created in Visual Studio 2017 will already use version 1.1.</span></span> <span data-ttu-id="f2137-106">ただし、バージョン 1.1 にはマイナー アップデートがときどきあります。これは、この記事に記載されている方法を使用して適用できます。</span><span class="sxs-lookup"><span data-stu-id="f2137-106">However there are occasional minor updates to version 1.1 that you can apply by using the techniques in this article.</span></span>

## <a name="use-the-most-up-to-date-project-files"></a><span data-ttu-id="f2137-107">最新のプロジェクト ファイルを使用する</span><span class="sxs-lookup"><span data-stu-id="f2137-107">Use the most up-to-date project files</span></span>

<span data-ttu-id="f2137-108">Visual Studio を使用してアドインを開発するときに、JavaScript API for Office の[最新の API メンバー](https://docs.microsoft.com/office/dev/add-ins/reference/what's-changed-in-the-javascript-api-for-office)と[アドイン マニフェスト v1.1 の機能](../develop/add-in-manifests.md) (offappmanifest-1.1.xsd に対して検証される) を使用する場合は、Visual Studio 2017 をダウンロードする必要があります。</span><span class="sxs-lookup"><span data-stu-id="f2137-108">If you use Visual Studio to develop your add-in, to use the [newest API members](https://docs.microsoft.com/office/dev/add-ins/reference/what's-changed-in-the-javascript-api-for-office) of the JavaScript API for Office and the [v1.1 features of the add-in manifest](../develop/add-in-manifests.md) (which is validated against offappmanifest-1.1.xsd), you need to download Visual Studio 2017.</span></span> <span data-ttu-id="f2137-109">Visual Studio 2017 をダウンロードするには、[Visual Studio IDE のページ](https://visualstudio.microsoft.com/vs/)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f2137-109">To download Visual Studio 2017, see the [Visual Studio IDE page](https://visualstudio.microsoft.com/vs/).</span></span> <span data-ttu-id="f2137-110">インストール時には、Office/SharePoint 開発ワークロードを選択する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f2137-110">During installation you'll need to select the Office/SharePoint development workload.</span></span>

<span data-ttu-id="f2137-111">テキスト エディター、または Visual Studio 以外の IDE を使用してアドインを開発する場合は、Office.js に対する CDN への参照と、アドインのマニフェストで参照するスキーマのバージョンを更新する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f2137-111">If you use a text editor or IDE other than Visual Studio to develop your add-in, you need to update the references to the CDN for Office.js and the version of schema referenced in your add-in's manifest.</span></span>

<span data-ttu-id="f2137-112">Office.js の新しい API や更新された API とアドインのマニフェスト機能を使用して開発したアドインを実行するには、ユーザー側で Office 2013 SP1 以降のオンプレミスの製品を実行し、該当する場合は SharePoint Server 2013 SP1 と関連するサーバー製品、Exchange Server 2013 Service Pack 1 (SP1)、または同等のオンライン ホスト製品である Office 365、SharePoint Online、および Exchange Online を実行している必要があります。</span><span class="sxs-lookup"><span data-stu-id="f2137-112">To run an add-in developed using new and updated Office.js API and add-in manifest features, your customers must be running Office 2013 SP1 or later version on-premises products, and where applicable, SharePoint Server 2013 SP1 and related server products, Exchange Server 2013 Service Pack 1 (SP1), or the equivalent online hosted products: Office 365, SharePoint Online, and Exchange Online.</span></span>

<span data-ttu-id="f2137-113">Office、SharePoint、Exchange SP1 の各製品をダウンロードするには、次を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f2137-113">To download Office, SharePoint, and Exchange SP1 products, see the following:</span></span>

- [<span data-ttu-id="f2137-114">Microsoft Office 2013 および関連のデスクトップ製品のすべての Service Pack 1 (SP1) の更新プログラムの一覧</span><span class="sxs-lookup"><span data-stu-id="f2137-114">List of all Service Pack 1 (SP1) updates for Microsoft Office 2013 and related desktop products</span></span>](https://support.microsoft.com/kb/2850036)
    
- [<span data-ttu-id="f2137-115">製品 Microsoft SharePoint Server 2013 と関連するサーバー製品のすべての Service Pack 1 (SP1) の更新プログラムの一覧</span><span class="sxs-lookup"><span data-stu-id="f2137-115">List of all Service Pack 1 (SP1) updates for Microsoft SharePoint Server 2013 and related server products</span></span>](https://support.microsoft.com/kb/2850035)
    
- [<span data-ttu-id="f2137-116">Exchange Server 2013 Service Pack 1 の説明</span><span class="sxs-lookup"><span data-stu-id="f2137-116">Description of Exchange Server 2013 Service Pack 1</span></span>](https://support.microsoft.com/kb/2926248)
    

## <a name="updating-an-office-add-in-project-created-with-visual-studio"></a><span data-ttu-id="f2137-117">Visual Studio で作成した Office アドイン プロジェクトを更新する</span><span class="sxs-lookup"><span data-stu-id="f2137-117">Updating an Office Add-in project created with Visual Studio</span></span>

<span data-ttu-id="f2137-118">JavaScript API for Office とアドイン マニフェスト スキーマの v1.1 のリリース前に作成されたプロジェクトの場合は、 **NuGet パッケージ マネージャー**を使用してプロジェクトのファイルを更新してから、それらを参照するようにアドインの HTML ページを更新できます。</span><span class="sxs-lookup"><span data-stu-id="f2137-118">For projects created before the release of v1.1 of the JavaScript API for Office and add-in manifest schema, you can update a project's files using the  **NuGet Package Manager**, and then update your add-in's HTML pages to reference them.</span></span> 

<span data-ttu-id="f2137-119">なお、この更新プロセスは _プロジェクトごと_ に適用する必要があることに注意してください。v1.1 の Office.js とアドイン マニフェスト スキーマを使用するアドイン プロジェクトごとに、この更新プロセスを繰り返します。</span><span class="sxs-lookup"><span data-stu-id="f2137-119">Note that the update process is applied on a  _per-project basis_ - you'll need to repeat the updating process for each add-in project in which you want to use v1.1 of Office.js and add-in manifest schema.</span></span>


### <a name="update-the-javascript-api-for-office-library-files-in-your-project-to-the-newest-release"></a><span data-ttu-id="f2137-120">プロジェクトの JavaScript API for Office ライブラリ ファイルを最新のリリースに更新する</span><span class="sxs-lookup"><span data-stu-id="f2137-120">Update the JavaScript API for Office library files in your project to the newest release</span></span>
<span data-ttu-id="f2137-121">次の手順では、Office のライブラリ ファイルを最新バージョンに更新します。</span><span class="sxs-lookup"><span data-stu-id="f2137-121">The following steps will update your Office library files to the latest version.</span></span> <span data-ttu-id="f2137-122">この手順では、Visual Studio 2017 を使用しますが、Visual Studio 2015 の手順とよく似ています。</span><span class="sxs-lookup"><span data-stu-id="f2137-122">The steps use Visual Studio 2017, but they are similar for Visual Studio 2015.</span></span>

1. <span data-ttu-id="f2137-123">Visual Studio 2017 で、**Office アドイン** プロジェクトを開くか新規作成します。</span><span class="sxs-lookup"><span data-stu-id="f2137-123">In Visual Studio 2017, open or create a new  **Office Add-in** project.</span></span>    
2. <span data-ttu-id="f2137-124">**[ツール]**、[**NuGet パッケージ マネージャー**]、[**ソリューションの Nuget パッケージの管理**] の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="f2137-124">Choose  **Tools** > **NuGet Package Manager** > **Manage Nuget Packages for Solution**.</span></span>
3. <span data-ttu-id="f2137-125">**[NuGet パッケージ マネージャー]** で **[パッケージ ソース]** に **[nuget.org]** を選択します。</span><span class="sxs-lookup"><span data-stu-id="f2137-125">In the  **NuGet Package Manager**, select  **nuget.org** for **Package source**.</span></span>
4. <span data-ttu-id="f2137-126">**[更新]** タブを選択します。</span><span class="sxs-lookup"><span data-stu-id="f2137-126">Choose the **Updates** tab.</span></span>
5. <span data-ttu-id="f2137-127">Microsoft.Office.js を選択します。</span><span class="sxs-lookup"><span data-stu-id="f2137-127">Select Microsoft.Office.js.</span></span>
6. <span data-ttu-id="f2137-128">左側のウィンドウで、**[更新]** を選択してパッケージの更新プロセスを完了します。</span><span class="sxs-lookup"><span data-stu-id="f2137-128">In the left pane, choose **Update** and complete the package update process.</span></span>

<span data-ttu-id="f2137-129">更新を完了するには、さらにいくつか手順を実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f2137-129">You'll need to take a few additional steps to complete the update.</span></span> <span data-ttu-id="f2137-130">アドインの HTML ページの **head** タグ内で、既存の office.js スクリプトに対する参照をすべてコメント アウトするか削除して、更新した JavaScript API for Office ライブラリを次のように参照します。</span><span class="sxs-lookup"><span data-stu-id="f2137-130">In the **head** tag of your add-in's HTML pages, comment out or delete any existing office.js script references, and reference the updated JavaScript API for Office library as follows:</span></span>
    
  ```html
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
  ```

   > [!NOTE] 
   > <span data-ttu-id="f2137-131">CDN URL の `office.js` の `/1/` は、Office.js のバージョン 1 内で最新の増分リリースを使用するように指定します。</span><span class="sxs-lookup"><span data-stu-id="f2137-131">The `/1/` in the `office.js` in the CDN URL specifies to use the latest incremental release within version 1 of Office.js.</span></span>


### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a><span data-ttu-id="f2137-132">プロジェクト内のマニフェスト ファイルがバージョン 1.1 のスキーマを使用するように更新する</span><span class="sxs-lookup"><span data-stu-id="f2137-132">Update the manifest file in your project to use schema version 1.1</span></span>

<span data-ttu-id="f2137-133">アドインのマニフェスト ファイルで、**OfficeApp** 要素の **xmlns** 属性のバージョン値を `1.1` に変更して更新します (**xmlns** 以外の属性は変更しません)。</span><span class="sxs-lookup"><span data-stu-id="f2137-133">In your add-in's manifest file, update the **xmlns** attribute of the **OfficeApp** element changing the version value to `1.1` (leaving attributes other than the **xmlns** attribute unchanged).</span></span>
    
```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp" 
    xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" 
    xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE] 
> <span data-ttu-id="f2137-134">アドイン マニフェスト スキーマのバージョンを 1.1 に更新したら、**Capabilities** 要素と**Capability** 要素を削除し、それらを [Hosts](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/hosts) 要素と [Host](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/host) 要素または [Requirements 要素と Requirement 要素](specify-office-hosts-and-api-requirements.md)に置き換える必要があります。</span><span class="sxs-lookup"><span data-stu-id="f2137-134">After updating the version of the add-in manifest schema to 1.1, you will need to remove the  **Capabilities** and **Capability** elements, and replace them with either the [Hosts](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/hosts) and [Host](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/host) elements or the [Requirements and Requirement elements](specify-office-hosts-and-api-requirements.md).</span></span>

## <a name="updating-an-office-add-in-project-created-with-a-text-editor-or-other-ide"></a><span data-ttu-id="f2137-135">テキスト エディターまたは他の IDE で作成した Office アドイン プロジェクトを更新する</span><span class="sxs-lookup"><span data-stu-id="f2137-135">Updating an Office Add-in project created with a text editor or other IDE</span></span>

<span data-ttu-id="f2137-136">JavaScript API for Office とアドイン マニフェスト スキーマの v1.1 のリリース前に作成されたプロジェクトについては、v1.1 のライブラリの CDN を参照するようにアドインの HTML ページを更新し、スキーマ v1.1 を使用するようにアドインのマニフェスト ファイルを更新する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f2137-136">For projects created before the release of v1.1 of the JavaScript API for Office and add-in manifest schema, you need to update your add-in's HTML pages to reference CDN of the v1.1 library, and update your add-in's manifest file to use schema v1.1.</span></span> 

<span data-ttu-id="f2137-137">この更新プロセスは_プロジェクトごと_に適用します。そのため、v1.1 の Office.js とアドイン マニフェスト スキーマを使用するアドイン プロジェクトごとに、この更新プロセスを繰り返す必要があります。</span><span class="sxs-lookup"><span data-stu-id="f2137-137">The update process is applied on a  _per-project basis_ - you'll need to repeat the updating process for each add-in project in which you want to use v1.1 of Office.js and add-in manifest schema.</span></span>

<span data-ttu-id="f2137-138">Office アドインを開発するために、JavaScript API for Office ファイル (Office.js とアプリ固有の .js ファイル) のローカル コピーを用意する必要はありません (Office.js の CDN を参照すれば、実行時に必要なファイルがダウンロードされます)。それでも、ライブラリ ファイルのローカル コピーが必要な場合は、[NuGet コマンド ライン ユーティリティ](https://docs.nuget.org/consume/installing-nuget)の `Install-Package Microsoft.Office.js` コマンドを使用してダウンロードしてください。</span><span class="sxs-lookup"><span data-stu-id="f2137-138">You don't need local copies of the JavaScript API for Office files (Office.js and app-specific .js files) to develop anOffice Add-in (referencing the CDN for Office.js downloads the necessary files at runtime), but if you want a local copy of the library files you can use the [NuGet Command-Line Utility](https://docs.nuget.org/consume/installing-nuget) and the `Install-Package Microsoft.Office.js` command to download them.</span></span>

> [!NOTE] 
> <span data-ttu-id="f2137-139">v1.1 アドイン マニフェストの XSD (XML スキーマ定義) のコピーの取得については、「[Office アドインのマニフェスト向けのスキーマ リファレンス (v1.1)](../develop/add-in-manifests.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f2137-139">To get a copy of the XSD (XML Schema Definition) for the v1.1 add-in manifest, see the listing in [Schema reference for Office Add-ins manifests (v1.1)](../develop/add-in-manifests.md).</span></span>


### <a name="update-the-javascript-api-for-office-library-files-in-your-project-to-use-the-newest-release"></a><span data-ttu-id="f2137-140">最新のリリースを使用するようにプロジェクトの JavaScript API for Office ライブラリ ファイルを更新する</span><span class="sxs-lookup"><span data-stu-id="f2137-140">Update the JavaScript API for Office library files in your project to use the newest release</span></span>

1. <span data-ttu-id="f2137-141">テキスト エディターまたは IDE でアドインの HTML ページを開きます。</span><span class="sxs-lookup"><span data-stu-id="f2137-141">Open the HTML pages for your add-in in your text editor or IDE.</span></span>
    
2. <span data-ttu-id="f2137-142">アドインの HTML ページの **head** タグ内で、既存の office.js スクリプトに対する参照をすべてコメント アウトするか削除して、更新した JavaScript API for Office ライブラリを次のように参照します。</span><span class="sxs-lookup"><span data-stu-id="f2137-142">In the **head** tag of your add-in's HTML pages, comment out or delete any existing office.js script references, and reference the updated JavaScript API for Office library as follows:</span></span>
    
    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
    ```

   > [!NOTE] 
   > <span data-ttu-id="f2137-143">CDN URL で `office.js` の前にある `/1/` は、Office.js のバージョン 1 内で最新の増分リリースを使用するように指定します。</span><span class="sxs-lookup"><span data-stu-id="f2137-143">The `/1/` in front of `office.js` in the CDN URL specifies to use the latest incremental release within version 1 of Office.js.</span></span>   

### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a><span data-ttu-id="f2137-144">プロジェクト内のマニフェスト ファイルがバージョン 1.1 のスキーマを使用するように更新する</span><span class="sxs-lookup"><span data-stu-id="f2137-144">Update the manifest file in your project to use schema version 1.1</span></span>

<span data-ttu-id="f2137-145">アドインのマニフェスト ファイルで、**OfficeApp** 要素の **xmlns** 属性のバージョン値を `1.1` に変更して更新します (**xmlns** 以外の属性は変更しません)。</span><span class="sxs-lookup"><span data-stu-id="f2137-145">In your add-in's manifest file, update the **xmlns** attribute of the **OfficeApp** element changing the version value to `1.1` (leaving attributes other than the **xmlns** attribute unchanged).</span></span>
    
```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp" 
    xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" 
    xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE] 
> <span data-ttu-id="f2137-146">アドイン マニフェスト スキーマのバージョンを 1.1 に更新したら、**Capabilities** 要素と**Capability** 要素を削除し、それらを [Hosts](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/hosts) 要素と [Host](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/host) 要素または [Requirements 要素と Requirement 要素](specify-office-hosts-and-api-requirements.md)に置き換える必要があります。</span><span class="sxs-lookup"><span data-stu-id="f2137-146">After updating the version of the add-in manifest schema to 1.1, you will need to remove the  **Capabilities** and **Capability** elements, and replace them with either the [Hosts](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/hosts) and [Host](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/host) elements or the [Requirements and Requirement elements](specify-office-hosts-and-api-requirements.md).</span></span>
    

## <a name="see-also"></a><span data-ttu-id="f2137-147">関連項目</span><span class="sxs-lookup"><span data-stu-id="f2137-147">See also</span></span>

- [<span data-ttu-id="f2137-148">Office のホストと API の要件を指定する</span><span class="sxs-lookup"><span data-stu-id="f2137-148">Specify Office hosts and API requirements</span></span>](specify-office-hosts-and-api-requirements.md) 
- [<span data-ttu-id="f2137-149">JavaScript API for Office について</span><span class="sxs-lookup"><span data-stu-id="f2137-149">Understanding the JavaScript API for Office</span></span>](understanding-the-javascript-api-for-office.md)    
- [<span data-ttu-id="f2137-150">JavaScript API for Office</span><span class="sxs-lookup"><span data-stu-id="f2137-150">JavaScript API for Office</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office)   
- [<span data-ttu-id="f2137-151">Office アドインのマニフェスト向けのスキーマ リファレンス (v1.1)</span><span class="sxs-lookup"><span data-stu-id="f2137-151">Schema reference for Office Add-ins manifests (v1.1)</span></span>](../develop/add-in-manifests.md)
    
