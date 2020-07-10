---
title: 最新の Office JavaScript API ライブラリおよびバージョン1.1 のアドインマニフェストスキーマへの更新
description: Office アドイン プロジェクトの JavaScript ファイル (Office.js ファイルとアプリに固有の .js ファイル) とアドイン マニフェスト検証ファイルをバージョン 1.1 に更新します。
ms.date: 10/11/2019
localization_priority: Normal
ms.openlocfilehash: 34127b3920af1309d4e4c2e1c265c676640a1c24
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093554"
---
# <a name="update-to-the-latest-office-javascript-api-library-and-version-11-add-in-manifest-schema"></a><span data-ttu-id="e3162-103">最新の Office JavaScript API ライブラリおよびバージョン1.1 のアドインマニフェストスキーマへの更新</span><span class="sxs-lookup"><span data-stu-id="e3162-103">Update to the latest Office JavaScript API library and version 1.1 add-in manifest schema</span></span>

<span data-ttu-id="e3162-104">この記事では、Office アドイン プロジェクトに含まれる JavaScript ファイル (Office.js およびアプリに固有の .js ファイル) とアドイン マニフェスト検証ファイルをバージョン 1.1 に更新する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="e3162-104">This article describes how to update your JavaScript files (Office.js and app-specific .js files) and add-in manifest validation file in your Office Add-in project to version 1.1.</span></span>

> [!NOTE]
> <span data-ttu-id="e3162-105">Visual Studio 2019 で作成されたプロジェクトは、既にバージョン1.1 を使用しています。</span><span class="sxs-lookup"><span data-stu-id="e3162-105">Projects created in Visual Studio 2019 will already use version 1.1.</span></span> <span data-ttu-id="e3162-106">ただし、バージョン 1.1 にはマイナー アップデートがときどきあります。これは、この記事に記載されている方法を使用して適用できます。</span><span class="sxs-lookup"><span data-stu-id="e3162-106">However there are occasional minor updates to version 1.1 that you can apply by using the techniques in this article.</span></span>

## <a name="use-the-most-up-to-date-project-files"></a><span data-ttu-id="e3162-107">最新のプロジェクト ファイルを使用する</span><span class="sxs-lookup"><span data-stu-id="e3162-107">Use the most up-to-date project files</span></span>

<span data-ttu-id="e3162-108">Visual Studio を使用してアドインを開発する場合は、Office JavaScript API の最新の API メンバーと[アドインマニフェストの v2.0 機能](../develop/add-in-manifests.md)(offappmanifest-に対して検証されます) を使用するには、visual studio 2019 をダウンロードする必要があります。</span><span class="sxs-lookup"><span data-stu-id="e3162-108">If you use Visual Studio to develop your add-in, to use the newest API members of the Office JavaScript API and the [v1.1 features of the add-in manifest](../develop/add-in-manifests.md) (which is validated against offappmanifest-1.1.xsd), you need to download Visual Studio 2019.</span></span> <span data-ttu-id="e3162-109">Visual Studio 2019 をダウンロードするには、 [Visual STUDIO IDE ページ](https://visualstudio.microsoft.com/vs/)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e3162-109">To download Visual Studio 2019, see the [Visual Studio IDE page](https://visualstudio.microsoft.com/vs/).</span></span> <span data-ttu-id="e3162-110">インストール時には、Office/SharePoint 開発ワークロードを選択する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e3162-110">During installation you'll need to select the Office/SharePoint development workload.</span></span>

<span data-ttu-id="e3162-111">テキスト エディター、または Visual Studio 以外の IDE を使用してアドインを開発する場合は、Office.js に対する CDN への参照と、アドインのマニフェストで参照するスキーマのバージョンを更新する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e3162-111">If you use a text editor or IDE other than Visual Studio to develop your add-in, you need to update the references to the CDN for Office.js and the version of schema referenced in your add-in's manifest.</span></span>

<span data-ttu-id="e3162-112">新規および更新された Office.js API およびアドインマニフェスト機能を使用して開発したアドインを実行するには、Office 2013 SP1 以降のバージョンのオンプレミス製品を実行している必要があり365ます。また、該当する場合は、SharePoint Server 2013 SP1 および関連するサーバー製品、Exchange Server 2013 Service Pack 1 (SP1)、またはそれと同等の</span><span class="sxs-lookup"><span data-stu-id="e3162-112">To run an add-in developed using new and updated Office.js API and add-in manifest features, your customers must be running Office 2013 SP1 or later version on-premises products, and where applicable, SharePoint Server 2013 SP1 and related server products, Exchange Server 2013 Service Pack 1 (SP1), or the equivalent online hosted products: Microsoft 365, SharePoint Online, and Exchange Online.</span></span>

<span data-ttu-id="e3162-113">Office、SharePoint、Exchange SP1 の各製品をダウンロードするには、次を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e3162-113">To download Office, SharePoint, and Exchange SP1 products, see the following:</span></span>

- [<span data-ttu-id="e3162-114">Microsoft Office 2013 および関連のデスクトップ製品のすべての Service Pack 1 (SP1) の更新プログラムの一覧</span><span class="sxs-lookup"><span data-stu-id="e3162-114">List of all Service Pack 1 (SP1) updates for Microsoft Office 2013 and related desktop products</span></span>](https://support.microsoft.com/kb/2850036)

- [<span data-ttu-id="e3162-115">製品 Microsoft SharePoint Server 2013 と関連するサーバー製品のすべての Service Pack 1 (SP1) の更新プログラムの一覧</span><span class="sxs-lookup"><span data-stu-id="e3162-115">List of all Service Pack 1 (SP1) updates for Microsoft SharePoint Server 2013 and related server products</span></span>](https://support.microsoft.com/kb/2850035)

- [<span data-ttu-id="e3162-116">Exchange Server 2013 Service Pack 1 の説明</span><span class="sxs-lookup"><span data-stu-id="e3162-116">Description of Exchange Server 2013 Service Pack 1</span></span>](https://support.microsoft.com/kb/2926248)


## <a name="updating-an-office-add-in-project-created-with-visual-studio"></a><span data-ttu-id="e3162-117">Visual Studio で作成した Office アドイン プロジェクトを更新する</span><span class="sxs-lookup"><span data-stu-id="e3162-117">Updating an Office Add-in project created with Visual Studio</span></span>

<span data-ttu-id="e3162-118">Office JavaScript API およびアドインマニフェストスキーマのリリースの前に作成されたプロジェクトでは、 **NuGet パッケージマネージャー**を使用してプロジェクトのファイルを更新し、アドインの HTML ページを更新してそれらを参照することができます。</span><span class="sxs-lookup"><span data-stu-id="e3162-118">For projects created before the release of v1.1 of the Office JavaScript API and add-in manifest schema, you can update a project's files using the **NuGet Package Manager**, and then update your add-in's HTML pages to reference them.</span></span> 

<span data-ttu-id="e3162-119">なお、この更新プロセスは _プロジェクトごと_ に適用する必要があることに注意してください。v1.1 の Office.js とアドイン マニフェスト スキーマを使用するアドイン プロジェクトごとに、この更新プロセスを繰り返します。</span><span class="sxs-lookup"><span data-stu-id="e3162-119">Note that the update process is applied on a  _per-project basis_ - you'll need to repeat the updating process for each add-in project in which you want to use v1.1 of Office.js and add-in manifest schema.</span></span>

### <a name="update-the-office-javascript-api-library-files-in-your-project-to-the-newest-release"></a><span data-ttu-id="e3162-120">プロジェクト内の Office JavaScript API ライブラリファイルを最新のリリースに更新する</span><span class="sxs-lookup"><span data-stu-id="e3162-120">Update the Office JavaScript API library files in your project to the newest release</span></span>
<span data-ttu-id="e3162-121">次の手順では、Office.js ライブラリファイルを最新バージョンに更新します。</span><span class="sxs-lookup"><span data-stu-id="e3162-121">The following steps will update your Office.js library files to the latest version.</span></span> <span data-ttu-id="e3162-122">手順では Visual Studio 2019 を使用していますが、以前のバージョンの Visual Studio に似ています。</span><span class="sxs-lookup"><span data-stu-id="e3162-122">The steps use Visual Studio 2019, but they are similar for previous versions of Visual Studio.</span></span>

1. <span data-ttu-id="e3162-123">Visual Studio 2019 で、 **Office アドイン**プロジェクトを開くか新規作成します。</span><span class="sxs-lookup"><span data-stu-id="e3162-123">In Visual Studio 2019, open or create a new **Office Add-in** project.</span></span>
2. <span data-ttu-id="e3162-124">**ツール**  >  の選択**nuget パッケージマネージャー**  >  **ソリューションの nuget パッケージを管理**します。</span><span class="sxs-lookup"><span data-stu-id="e3162-124">Choose **Tools** > **NuGet Package Manager** > **Manage Nuget Packages for Solution**.</span></span>
3. <span data-ttu-id="e3162-125">**[更新]** タブを選択します。</span><span class="sxs-lookup"><span data-stu-id="e3162-125">Choose the **Updates** tab.</span></span>
4. <span data-ttu-id="e3162-126">Microsoft.Office.js を選択します。</span><span class="sxs-lookup"><span data-stu-id="e3162-126">Select Microsoft.Office.js.</span></span> <span data-ttu-id="e3162-127">パッケージソースが**nuget.org**からのものであることを確認します。</span><span class="sxs-lookup"><span data-stu-id="e3162-127">Ensure the package source is from **nuget.org**.</span></span>
5. <span data-ttu-id="e3162-128">左側のウィンドウで、[**インストール**] を選択し、パッケージの更新プロセスを完了します。</span><span class="sxs-lookup"><span data-stu-id="e3162-128">In the left pane, choose **Install** and complete the package update process.</span></span>

<span data-ttu-id="e3162-129">更新を完了するには、さらにいくつか手順を実行する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e3162-129">You'll need to take a few additional steps to complete the update.</span></span> <span data-ttu-id="e3162-130">アドインの HTML ページの**head**タグで、既存の office.js スクリプト参照をコメントアウトまたは削除し、更新された OFFICE JavaScript API ライブラリを次のように参照します。</span><span class="sxs-lookup"><span data-stu-id="e3162-130">In the **head** tag of your add-in's HTML pages, comment out or delete any existing office.js script references, and reference the updated Office JavaScript API library as follows:</span></span>

  ```html
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
  ```

   > [!NOTE] 
   > <span data-ttu-id="e3162-131">CDN URL の `office.js` の `/1/` は、Office.js のバージョン 1 内で最新の増分リリースを使用するように指定します。</span><span class="sxs-lookup"><span data-stu-id="e3162-131">The `/1/` in the `office.js` in the CDN URL specifies to use the latest incremental release within version 1 of Office.js.</span></span>


### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a><span data-ttu-id="e3162-132">プロジェクト内のマニフェスト ファイルがバージョン 1.1 のスキーマを使用するように更新する</span><span class="sxs-lookup"><span data-stu-id="e3162-132">Update the manifest file in your project to use schema version 1.1</span></span>

<span data-ttu-id="e3162-133">アドインのマニフェスト ファイルで、**OfficeApp** 要素の **xmlns** 属性のバージョン値を `1.1` に変更して更新します (**xmlns** 以外の属性は変更しません)。</span><span class="sxs-lookup"><span data-stu-id="e3162-133">In your add-in's manifest file, update the **xmlns** attribute of the **OfficeApp** element changing the version value to `1.1` (leaving attributes other than the **xmlns** attribute unchanged).</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE]
> <span data-ttu-id="e3162-134">アドインマニフェストスキーマのバージョンを1.1 に更新した後で、**機能**要素と**機能**要素を削除し、それらを[Hosts](../reference/manifest/hosts.md)要素と[Host](../reference/manifest/host.md)要素、または[要件と要件要素](specify-office-hosts-and-api-requirements.md)のいずれかに置き換える必要があります。</span><span class="sxs-lookup"><span data-stu-id="e3162-134">After updating the version of the add-in manifest schema to 1.1, you will need to remove the **Capabilities** and **Capability** elements, and replace them with either the [Hosts](../reference/manifest/hosts.md) and [Host](../reference/manifest/host.md) elements or the [Requirements and Requirement elements](specify-office-hosts-and-api-requirements.md).</span></span>

## <a name="updating-an-office-add-in-project-created-with-a-text-editor-or-other-ide"></a><span data-ttu-id="e3162-135">テキスト エディターまたは他の IDE で作成した Office アドイン プロジェクトを更新する</span><span class="sxs-lookup"><span data-stu-id="e3162-135">Updating an Office Add-in project created with a text editor or other IDE</span></span>

<span data-ttu-id="e3162-136">Office JavaScript API およびアドインマニフェストスキーマのリリースの前に作成されたプロジェクトでは、アドインの HTML ページを更新して v2.0 ライブラリの CDN を参照し、アドインのマニフェストファイルを更新してスキーマ v1.1 を使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e3162-136">For projects created before the release of v1.1 of the Office JavaScript API and add-in manifest schema, you need to update your add-in's HTML pages to reference CDN of the v1.1 library, and update your add-in's manifest file to use schema v1.1.</span></span> 

<span data-ttu-id="e3162-137">この更新プロセスは_プロジェクトごと_に適用します。そのため、v1.1 の Office.js とアドイン マニフェスト スキーマを使用するアドイン プロジェクトごとに、この更新プロセスを繰り返す必要があります。</span><span class="sxs-lookup"><span data-stu-id="e3162-137">The update process is applied on a  _per-project basis_ - you'll need to repeat the updating process for each add-in project in which you want to use v1.1 of Office.js and add-in manifest schema.</span></span>

<span data-ttu-id="e3162-138">Office の JavaScript API ファイル (Office.js およびアプリ固有の .js ファイル) のローカルコピーは、Office アドインを開発する必要はありません (Office.js の CDN を参照すると、必要なファイルが実行時にダウンロードされます) が、ライブラリファイルのローカルコピーが必要な場合は、 [NuGet コマンドラインユーティリティ](https://docs.nuget.org/consume/installing-nuget)とコマンドを使用して `Install-Package Microsoft.Office.js` ダウンロードできます。</span><span class="sxs-lookup"><span data-stu-id="e3162-138">You don't need local copies of the Office JavaScript API files (Office.js and app-specific .js files) to develop anOffice Add-in (referencing the CDN for Office.js downloads the necessary files at runtime), but if you want a local copy of the library files you can use the [NuGet Command-Line Utility](https://docs.nuget.org/consume/installing-nuget) and the `Install-Package Microsoft.Office.js` command to download them.</span></span>

> [!NOTE]
> <span data-ttu-id="e3162-139">v1.1 アドイン マニフェストの XSD (XML スキーマ定義) のコピーの取得については、「[Office アドインのマニフェスト向けのスキーマ リファレンス (v1.1)](../develop/add-in-manifests.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="e3162-139">To get a copy of the XSD (XML Schema Definition) for the v1.1 add-in manifest, see the listing in [Schema reference for Office Add-ins manifests (v1.1)](../develop/add-in-manifests.md).</span></span>


### <a name="update-the-office-javascript-api-library-files-in-your-project-to-use-the-newest-release"></a><span data-ttu-id="e3162-140">最新のリリースを使用するようにプロジェクトの Office JavaScript API ライブラリファイルを更新する</span><span class="sxs-lookup"><span data-stu-id="e3162-140">Update the Office JavaScript API library files in your project to use the newest release</span></span>

1. <span data-ttu-id="e3162-141">テキスト エディターまたは IDE でアドインの HTML ページを開きます。</span><span class="sxs-lookup"><span data-stu-id="e3162-141">Open the HTML pages for your add-in in your text editor or IDE.</span></span>

2. <span data-ttu-id="e3162-142">アドインの HTML ページの**head**タグで、既存の office.js スクリプト参照をコメントアウトまたは削除し、更新された OFFICE JavaScript API ライブラリを次のように参照します。</span><span class="sxs-lookup"><span data-stu-id="e3162-142">In the **head** tag of your add-in's HTML pages, comment out or delete any existing office.js script references, and reference the updated Office JavaScript API library as follows:</span></span>

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
    ```

   > [!NOTE]
   > <span data-ttu-id="e3162-143">CDN URL で `office.js` の前にある `/1/` は、Office.js のバージョン 1 内で最新の増分リリースを使用するように指定します。</span><span class="sxs-lookup"><span data-stu-id="e3162-143">The `/1/` in front of `office.js` in the CDN URL specifies to use the latest incremental release within version 1 of Office.js.</span></span>

### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a><span data-ttu-id="e3162-144">プロジェクト内のマニフェスト ファイルがバージョン 1.1 のスキーマを使用するように更新する</span><span class="sxs-lookup"><span data-stu-id="e3162-144">Update the manifest file in your project to use schema version 1.1</span></span>

<span data-ttu-id="e3162-145">アドインのマニフェスト ファイルで、**OfficeApp** 要素の **xmlns** 属性のバージョン値を `1.1` に変更して更新します (**xmlns** 以外の属性は変更しません)。</span><span class="sxs-lookup"><span data-stu-id="e3162-145">In your add-in's manifest file, update the **xmlns** attribute of the **OfficeApp** element changing the version value to `1.1` (leaving attributes other than the **xmlns** attribute unchanged).</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE]
> <span data-ttu-id="e3162-146">アドインマニフェストスキーマのバージョンを1.1 に更新した後で、**機能**要素と**機能**要素を削除し、それらを[Hosts](../reference/manifest/hosts.md)要素と[Host](../reference/manifest/host.md)要素、または[要件と要件要素](specify-office-hosts-and-api-requirements.md)のいずれかに置き換える必要があります。</span><span class="sxs-lookup"><span data-stu-id="e3162-146">After updating the version of the add-in manifest schema to 1.1, you will need to remove the **Capabilities** and **Capability** elements, and replace them with either the [Hosts](../reference/manifest/hosts.md) and [Host](../reference/manifest/host.md) elements or the [Requirements and Requirement elements](specify-office-hosts-and-api-requirements.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="e3162-147">関連項目</span><span class="sxs-lookup"><span data-stu-id="e3162-147">See also</span></span>

- <span data-ttu-id="e3162-148">[Office のホストと API の要件を指定する](specify-office-hosts-and-api-requirements.md)]</span><span class="sxs-lookup"><span data-stu-id="e3162-148">[Specify Office hosts and API requirements](specify-office-hosts-and-api-requirements.md) ]</span></span>
- [<span data-ttu-id="e3162-149">Office JavaScript API について</span><span class="sxs-lookup"><span data-stu-id="e3162-149">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="e3162-150">Office の JavaScript API</span><span class="sxs-lookup"><span data-stu-id="e3162-150">Office JavaScript API</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="e3162-151">Office アドインのマニフェスト向けのスキーマ リファレンス (v1.1)</span><span class="sxs-lookup"><span data-stu-id="e3162-151">Schema reference for Office Add-ins manifests (v1.1)</span></span>](../develop/add-in-manifests.md)
