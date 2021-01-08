---
title: Office アドインの Office クライアント アプリケーションとプラットフォームの可用性
description: Excel、OneNote、Outlook、PowerPoint、Project、Word のサポートされる要件セット。
ms.date: 01/06/2021
localization_priority: Priority
ms.openlocfilehash: e09c56076a1650a4a88fb417878369c08b137e21
ms.sourcegitcommit: 7b8906fdb5cc2db93dae09d1a2c0334f87062510
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/08/2021
ms.locfileid: "49784472"
---
# <a name="office-client-application-and-platform-availability-for-office-add-ins"></a><span data-ttu-id="4b285-103">Office アドインの Office クライアント アプリケーションとプラットフォームの可用性</span><span class="sxs-lookup"><span data-stu-id="4b285-103">Office client application and platform availability for Office Add-ins</span></span>

<span data-ttu-id="4b285-p101">期待どおりの動作をするうえで、Office アドインは特定の Office アプリケーション、要件セット、API メンバー、または API のバージョンに依存することがあります。次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、および各 Office アプリケーションで現在サポートされている共通 API が含まれています。</span><span class="sxs-lookup"><span data-stu-id="4b285-p101">To work as expected, your Office Add-in might depend on a specific Office application, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

<br>

|<a href="#excel"><img src="../images/index/logo-excel.svg" alt="Excel" width="48" /><br><span data-ttu-id="4b285-106"><span>Excel</span></a></span><span class="sxs-lookup"><span data-stu-id="4b285-106"><span>Excel</span></a></span></span>|<a href="#onenote"><img src="../images/index/logo-onenote.svg" alt="OneNote" width="48" /><br><span data-ttu-id="4b285-107"><span>OneNote</span></a></span><span class="sxs-lookup"><span data-stu-id="4b285-107"><span>OneNote</span></a></span></span>|<a href="#outlook"><img src="../images/index/logo-outlook.svg" alt="Outlook" width="48" /><br><span data-ttu-id="4b285-108"><span>Outlook</span></a></span><span class="sxs-lookup"><span data-stu-id="4b285-108"><span>Outlook</span></a></span></span>|<a href="#powerpoint"><img src="../images/index/logo-powerpoint.svg" alt="PowerPoint" width="48" /><br><span data-ttu-id="4b285-109"><span>PowerPoint</span></a></span><span class="sxs-lookup"><span data-stu-id="4b285-109"><span>PowerPoint</span></a></span></span>|<a href="#project"><img src="../images/index/logo-project-server.svg" alt="Project" width="48" /><br><span data-ttu-id="4b285-110"><span>Project</span></a></span><span class="sxs-lookup"><span data-stu-id="4b285-110"><span>Project</span></a></span></span>|<a href="#word"><img src="../images/index/logo-word.svg" alt="Word" width="48" /><br><span data-ttu-id="4b285-111"><span>Word</span></a></span><span class="sxs-lookup"><span data-stu-id="4b285-111"><span>Word</span></a></span></span>|
|:---:|:---:|:---:|:---:|:---:|:---:|

> [!NOTE]
> <span data-ttu-id="4b285-112">MSI からインストールされる最初の Office 2016 リリースには、ExcelApi 1.1、WordApi 1.1、共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="4b285-112">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="4b285-113">さまざまなバージョンの Office の更新履歴の詳細については、「[関連項目](#see-also)」セクションをご確認ください。</span><span class="sxs-lookup"><span data-stu-id="4b285-113">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="4b285-114">Excel</span><span class="sxs-lookup"><span data-stu-id="4b285-114">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="4b285-115">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="4b285-115">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="4b285-116">拡張点</span><span class="sxs-lookup"><span data-stu-id="4b285-116">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="4b285-117">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="4b285-117">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="4b285-118"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="4b285-118"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-119">Office on the web</span><span class="sxs-lookup"><span data-stu-id="4b285-119">Office on the web</span></span></td>
    <td><span data-ttu-id="4b285-120">
      - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4b285-120">
      - TaskPane</span></span><br><span data-ttu-id="4b285-121">
      - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4b285-121">
      - Content</span></span><br><span data-ttu-id="4b285-122">
      - CustomFunctions</span><span class="sxs-lookup"><span data-stu-id="4b285-122">
      - CustomFunctions</span></span><br><span data-ttu-id="4b285-123">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-123">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-124">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-124">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4b285-125">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4b285-125">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4b285-126">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4b285-126">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4b285-127">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4b285-127">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4b285-128">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4b285-128">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4b285-129">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4b285-129">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4b285-130">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4b285-130">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4b285-131">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4b285-131">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4b285-132">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="4b285-132">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="4b285-133">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="4b285-133">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="4b285-134">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="4b285-134">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="4b285-135">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="4b285-135">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="4b285-136">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-136">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-137">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-137">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="4b285-138">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-138">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="4b285-139">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-139">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="4b285-140">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="4b285-140">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="4b285-141">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-141">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="4b285-142">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-142">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="4b285-143">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="4b285-143">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="4b285-144">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">設定</a></span><span class="sxs-lookup"><span data-stu-id="4b285-144">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="4b285-145">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-145">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="4b285-146">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-146">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="4b285-147">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-147">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="4b285-148">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-148">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-149">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="4b285-149">Office on Windows</span></span><br><span data-ttu-id="4b285-150">(Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="4b285-150">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="4b285-151">
      - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4b285-151">
      - TaskPane</span></span><br><span data-ttu-id="4b285-152">
      - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4b285-152">
      - Content</span></span><br><span data-ttu-id="4b285-153">
      - CustomFunctions</span><span class="sxs-lookup"><span data-stu-id="4b285-153">
      - CustomFunctions</span></span><br><span data-ttu-id="4b285-154">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-154">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-155">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-155">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4b285-156">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4b285-156">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4b285-157">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4b285-157">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4b285-158">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4b285-158">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4b285-159">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4b285-159">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4b285-160">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4b285-160">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4b285-161">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4b285-161">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4b285-162">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4b285-162">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4b285-163">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="4b285-163">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="4b285-164">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="4b285-164">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="4b285-165">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="4b285-165">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="4b285-166">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-166">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4b285-167">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-167">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="4b285-168">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-168">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-169">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-169">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="4b285-170">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-170">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="4b285-171">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-171">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="4b285-172">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="4b285-172">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="4b285-173">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-173">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="4b285-174">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-174">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="4b285-175">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="4b285-175">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="4b285-176">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">設定</a></span><span class="sxs-lookup"><span data-stu-id="4b285-176">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="4b285-177">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-177">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="4b285-178">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-178">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="4b285-179">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-179">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="4b285-180">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-180">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-181">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="4b285-181">Office 2019 on Windows</span></span><br><span data-ttu-id="4b285-182">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4b285-182">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="4b285-183">
      - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4b285-183">
      - TaskPane</span></span><br><span data-ttu-id="4b285-184">
      - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4b285-184">
      - Content</span></span><br><span data-ttu-id="4b285-185">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-185">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-186">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-186">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4b285-187">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4b285-187">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4b285-188">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4b285-188">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4b285-189">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4b285-189">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4b285-190">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4b285-190">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4b285-191">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4b285-191">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4b285-192">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4b285-192">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4b285-193">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4b285-193">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4b285-194">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-194">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4b285-195">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-195">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-196">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-196">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="4b285-197">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-197">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="4b285-198">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-198">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="4b285-199">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="4b285-199">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="4b285-200">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-200">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="4b285-201">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-201">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="4b285-202">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="4b285-202">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="4b285-203">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">設定</a></span><span class="sxs-lookup"><span data-stu-id="4b285-203">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="4b285-204">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-204">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="4b285-205">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-205">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="4b285-206">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-206">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="4b285-207">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-207">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-208">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="4b285-208">Office 2016 on Windows</span></span><br><span data-ttu-id="4b285-209">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4b285-209">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="4b285-210">
      - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4b285-210">
      - TaskPane</span></span><br><span data-ttu-id="4b285-211">
      - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4b285-211">
      - Content</span></span> </td>
    <td><span data-ttu-id="4b285-212">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-212">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4b285-213">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="4b285-213">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="4b285-214">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-214">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-215">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-215">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="4b285-216">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-216">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="4b285-217">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-217">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="4b285-218">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="4b285-218">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="4b285-219">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-219">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="4b285-220">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-220">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="4b285-221">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="4b285-221">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="4b285-222">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">設定</a></span><span class="sxs-lookup"><span data-stu-id="4b285-222">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="4b285-223">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-223">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="4b285-224">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-224">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="4b285-225">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-225">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="4b285-226">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-226">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-227">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="4b285-227">Office 2013 on Windows</span></span><br><span data-ttu-id="4b285-228">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4b285-228">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="4b285-229">
      - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4b285-229">
      - TaskPane</span></span><br><span data-ttu-id="4b285-230">
      - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4b285-230">
      - Content</span></span> </td>
    <td><span data-ttu-id="4b285-231">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="4b285-231">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="4b285-232">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-232">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-233">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-233">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="4b285-234">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-234">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="4b285-235">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="4b285-235">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="4b285-236">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-236">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="4b285-237">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-237">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="4b285-238">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="4b285-238">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="4b285-239">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">設定</a></span><span class="sxs-lookup"><span data-stu-id="4b285-239">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="4b285-240">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-240">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="4b285-241">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-241">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="4b285-242">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-242">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="4b285-243">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-243">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-244">Office on iPad</span><span class="sxs-lookup"><span data-stu-id="4b285-244">Office on iPad</span></span><br><span data-ttu-id="4b285-245">(Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="4b285-245">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="4b285-246">
      - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4b285-246">
      - TaskPane</span></span><br><span data-ttu-id="4b285-247">
      - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4b285-247">
      - Content</span></span> </td>
    <td><span data-ttu-id="4b285-248">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-248">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4b285-249">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4b285-249">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4b285-250">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4b285-250">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4b285-251">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4b285-251">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4b285-252">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4b285-252">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4b285-253">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4b285-253">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4b285-254">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4b285-254">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4b285-255">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4b285-255">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4b285-256">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="4b285-256">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="4b285-257">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="4b285-257">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="4b285-258">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="4b285-258">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="4b285-259">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-259">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4b285-260">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-260">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-261">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-261">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="4b285-262">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-262">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="4b285-263">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="4b285-263">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="4b285-264">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-264">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="4b285-265">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-265">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="4b285-266">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="4b285-266">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="4b285-267">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">設定</a></span><span class="sxs-lookup"><span data-stu-id="4b285-267">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="4b285-268">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-268">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="4b285-269">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-269">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="4b285-270">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-270">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="4b285-271">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-271">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-272">Office on Mac</span><span class="sxs-lookup"><span data-stu-id="4b285-272">Office on Mac</span></span><br><span data-ttu-id="4b285-273">(Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="4b285-273">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="4b285-274">
      - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4b285-274">
      - TaskPane</span></span><br><span data-ttu-id="4b285-275">
      - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4b285-275">
      - Content</span></span><br><span data-ttu-id="4b285-276">
      - CustomFunctions</span><span class="sxs-lookup"><span data-stu-id="4b285-276">
      - CustomFunctions</span></span><br><span data-ttu-id="4b285-277">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-277">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-278">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-278">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4b285-279">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4b285-279">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4b285-280">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4b285-280">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4b285-281">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4b285-281">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4b285-282">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4b285-282">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4b285-283">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4b285-283">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4b285-284">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4b285-284">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4b285-285">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4b285-285">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4b285-286">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="4b285-286">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="4b285-287">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="4b285-287">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="4b285-288">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="4b285-288">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="4b285-289">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-289">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4b285-290">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-290">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="4b285-291">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-291">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-292">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-292">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="4b285-293">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-293">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="4b285-294">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-294">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="4b285-295">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="4b285-295">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="4b285-296">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-296">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="4b285-297">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-297">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="4b285-298">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-298">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="4b285-299">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="4b285-299">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="4b285-300">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">設定</a></span><span class="sxs-lookup"><span data-stu-id="4b285-300">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="4b285-301">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-301">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="4b285-302">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-302">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="4b285-303">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-303">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="4b285-304">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-304">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-305">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="4b285-305">Office 2019 on Mac</span></span><br><span data-ttu-id="4b285-306">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4b285-306">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="4b285-307">
      - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4b285-307">
      - TaskPane</span></span><br><span data-ttu-id="4b285-308">
      - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4b285-308">
      - Content</span></span><br><span data-ttu-id="4b285-309">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-309">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-310">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-310">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4b285-311">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4b285-311">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="4b285-312">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4b285-312">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="4b285-313">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4b285-313">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="4b285-314">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4b285-314">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="4b285-315">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4b285-315">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="4b285-316">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4b285-316">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="4b285-317">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4b285-317">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="4b285-318">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-318">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4b285-319">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-319">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-320">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-320">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="4b285-321">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-321">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="4b285-322">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-322">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="4b285-323">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="4b285-323">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="4b285-324">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-324">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="4b285-325">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-325">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="4b285-326">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-326">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="4b285-327">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="4b285-327">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="4b285-328">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">設定</a></span><span class="sxs-lookup"><span data-stu-id="4b285-328">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="4b285-329">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-329">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="4b285-330">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-330">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="4b285-331">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-331">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="4b285-332">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-332">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-333">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="4b285-333">Office 2016 on Mac</span></span><br><span data-ttu-id="4b285-334">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4b285-334">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="4b285-335">
      - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4b285-335">
      - TaskPane</span></span><br><span data-ttu-id="4b285-336">
      - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4b285-336">
      - Content</span></span> </td>
    <td><span data-ttu-id="4b285-337">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-337">
      - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="4b285-338">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="4b285-338">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="4b285-339">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-339">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-340">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-340">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="4b285-341">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-341">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="4b285-342">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-342">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="4b285-343">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="4b285-343">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="4b285-344">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-344">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="4b285-345">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-345">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="4b285-346">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-346">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="4b285-347">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="4b285-347">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="4b285-348">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">設定</a></span><span class="sxs-lookup"><span data-stu-id="4b285-348">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="4b285-349">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-349">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="4b285-350">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-350">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="4b285-351">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-351">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="4b285-352">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-352">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
</table>

<span data-ttu-id="4b285-353">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="4b285-353">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="4b285-354">カスタム関数 (Excel のみ)</span><span class="sxs-lookup"><span data-stu-id="4b285-354">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4b285-355">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="4b285-355">Platform</span></span></th>
    <th><span data-ttu-id="4b285-356">拡張点</span><span class="sxs-lookup"><span data-stu-id="4b285-356">Extension points</span></span></th>
    <th><span data-ttu-id="4b285-357">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="4b285-357">API requirement sets</span></span></th>
    <th><span data-ttu-id="4b285-358"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="4b285-358"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-359">Office on the web</span><span class="sxs-lookup"><span data-stu-id="4b285-359">Office on the web</span></span></td>
    <td><span data-ttu-id="4b285-360">- CustomFunctions</span><span class="sxs-lookup"><span data-stu-id="4b285-360">- CustomFunctions</span></span></td>
    <td><span data-ttu-id="4b285-361">
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-361">
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span><br><span data-ttu-id="4b285-362">
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4b285-362">
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.2</a></span></span><br><span data-ttu-id="4b285-363">
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.3</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-363">
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.3</a>
    </span></span></td>
    <td></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-364">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="4b285-364">Office on Windows</span></span><br><span data-ttu-id="4b285-365">(Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="4b285-365">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="4b285-366">- CustomFunctions</span><span class="sxs-lookup"><span data-stu-id="4b285-366">- CustomFunctions</span></span></td>
    <td><span data-ttu-id="4b285-367">
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-367">
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span><br><span data-ttu-id="4b285-368">
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4b285-368">
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.2</a></span></span><br><span data-ttu-id="4b285-369">
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.3</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-369">
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.3</a>
    </span></span></td>
    <td></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-370">Office on Mac</span><span class="sxs-lookup"><span data-stu-id="4b285-370">Office on Mac</span></span><br><span data-ttu-id="4b285-371">(Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="4b285-371">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="4b285-372">- CustomFunctions</span><span class="sxs-lookup"><span data-stu-id="4b285-372">- CustomFunctions</span></span></td>
    <td><span data-ttu-id="4b285-373">
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-373">
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span><br><span data-ttu-id="4b285-374">
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4b285-374">
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.2</a></span></span><br><span data-ttu-id="4b285-375">
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.3</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-375">
      - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.3</a>
    </span></span></td>
    <td></td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="4b285-376">Outlook</span><span class="sxs-lookup"><span data-stu-id="4b285-376">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4b285-377">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="4b285-377">Platform</span></span></th>
    <th><span data-ttu-id="4b285-378">拡張点</span><span class="sxs-lookup"><span data-stu-id="4b285-378">Extension points</span></span></th>
    <th><span data-ttu-id="4b285-379">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="4b285-379">API requirement sets</span></span></th>
    <th><span data-ttu-id="4b285-380"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="4b285-380"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-381">Office on the web</span><span class="sxs-lookup"><span data-stu-id="4b285-381">Office on the web</span></span><br><span data-ttu-id="4b285-382">(モダン)</span><span class="sxs-lookup"><span data-stu-id="4b285-382">(modern)</span></span></td>
    <td><span data-ttu-id="4b285-383">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="4b285-383">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="4b285-384">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="4b285-384">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="4b285-385">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="4b285-385">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="4b285-386">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="4b285-386">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="4b285-387">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-387">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-388">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4b285-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4b285-389">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4b285-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4b285-390">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4b285-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4b285-391">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4b285-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4b285-392">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4b285-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4b285-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="4b285-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4b285-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="4b285-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4b285-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span><br><span data-ttu-id="4b285-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9">Mailbox 1.9</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9">Mailbox 1.9</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-397">使用不可</span><span class="sxs-lookup"><span data-stu-id="4b285-397">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-398">Office on the web</span><span class="sxs-lookup"><span data-stu-id="4b285-398">Office on the web</span></span><br><span data-ttu-id="4b285-399">(クラシック)</span><span class="sxs-lookup"><span data-stu-id="4b285-399">(classic)</span></span></td>
    <td><span data-ttu-id="4b285-400">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="4b285-400">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="4b285-401">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="4b285-401">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="4b285-402">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="4b285-402">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="4b285-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="4b285-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="4b285-404">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-404">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-405">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4b285-406">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4b285-406">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4b285-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4b285-407">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4b285-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4b285-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4b285-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4b285-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4b285-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-411">利用不可</span><span class="sxs-lookup"><span data-stu-id="4b285-411">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-412">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="4b285-412">Office on Windows</span></span><br><span data-ttu-id="4b285-413">(Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="4b285-413">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="4b285-414">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="4b285-414">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="4b285-415">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="4b285-415">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="4b285-416">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="4b285-416">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="4b285-417">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="4b285-417">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="4b285-418">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4b285-418">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="4b285-419">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">モジュール</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-419">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-420">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4b285-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4b285-421">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4b285-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4b285-422">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4b285-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4b285-423">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4b285-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4b285-424">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4b285-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4b285-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="4b285-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4b285-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="4b285-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="4b285-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span><br><span data-ttu-id="4b285-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9">Mailbox 1.9</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9">Mailbox 1.9</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-429">使用不可</span><span class="sxs-lookup"><span data-stu-id="4b285-429">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-430">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="4b285-430">Office 2019 on Windows</span></span><br><span data-ttu-id="4b285-431">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4b285-431">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="4b285-432">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="4b285-432">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="4b285-433">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="4b285-433">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="4b285-434">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="4b285-434">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="4b285-435">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="4b285-435">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="4b285-436">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4b285-436">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="4b285-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">モジュール</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-438">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4b285-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4b285-439">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4b285-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4b285-440">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4b285-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4b285-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4b285-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4b285-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4b285-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4b285-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="4b285-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-444">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-445">利用不可</span><span class="sxs-lookup"><span data-stu-id="4b285-445">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-446">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="4b285-446">Office 2016 on Windows</span></span><br><span data-ttu-id="4b285-447">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4b285-447">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="4b285-448">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="4b285-448">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="4b285-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="4b285-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="4b285-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="4b285-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="4b285-451">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="4b285-451">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="4b285-452">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="4b285-452">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="4b285-453">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">モジュール</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-453">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4b285-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4b285-455">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4b285-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4b285-456">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4b285-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*
    </span><span class="sxs-lookup"><span data-stu-id="4b285-457">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*
    </span></span></td>
    <td><span data-ttu-id="4b285-458">使用不可</span><span class="sxs-lookup"><span data-stu-id="4b285-458">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-459">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="4b285-459">Office 2013 on Windows</span></span><br><span data-ttu-id="4b285-460">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4b285-460">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="4b285-461">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="4b285-461">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="4b285-462">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="4b285-462">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="4b285-463">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="4b285-463">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="4b285-464">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="4b285-464">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br>
    </td>
    <td><span data-ttu-id="4b285-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-465">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4b285-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4b285-466">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4b285-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="4b285-467">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="4b285-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*
    </span><span class="sxs-lookup"><span data-stu-id="4b285-468">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*
    </span></span></td>
    <td><span data-ttu-id="4b285-469">使用不可</span><span class="sxs-lookup"><span data-stu-id="4b285-469">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-470">iOS 上の Office</span><span class="sxs-lookup"><span data-stu-id="4b285-470">Office on iOS</span></span><br><span data-ttu-id="4b285-471">(Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="4b285-471">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="4b285-472">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="4b285-472">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="4b285-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-473">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4b285-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4b285-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4b285-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4b285-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4b285-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4b285-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4b285-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-479">利用不可</span><span class="sxs-lookup"><span data-stu-id="4b285-479">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-480">Office on Mac</span><span class="sxs-lookup"><span data-stu-id="4b285-480">Office on Mac</span></span><br><span data-ttu-id="4b285-481">(Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="4b285-481">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="4b285-482">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="4b285-482">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="4b285-483">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="4b285-483">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="4b285-484">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="4b285-484">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="4b285-485">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="4b285-485">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="4b285-486">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-486">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-487">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4b285-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4b285-488">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4b285-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4b285-489">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4b285-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4b285-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4b285-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4b285-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4b285-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="4b285-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="4b285-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="4b285-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="4b285-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-495">利用不可</span><span class="sxs-lookup"><span data-stu-id="4b285-495">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-496">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="4b285-496">Office 2019 on Mac</span></span><br><span data-ttu-id="4b285-497">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4b285-497">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="4b285-498">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="4b285-498">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="4b285-499">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="4b285-499">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="4b285-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="4b285-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="4b285-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="4b285-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="4b285-502">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-502">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-503">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-503">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4b285-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4b285-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4b285-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4b285-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4b285-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4b285-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4b285-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4b285-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4b285-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-509">利用不可</span><span class="sxs-lookup"><span data-stu-id="4b285-509">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-510">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="4b285-510">Office 2016 on Mac</span></span><br><span data-ttu-id="4b285-511">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4b285-511">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="4b285-512">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="4b285-512">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="4b285-513">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="4b285-513">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="4b285-514">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="4b285-514">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="4b285-515">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="4b285-515">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="4b285-516">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-516">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4b285-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4b285-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4b285-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4b285-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4b285-520">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4b285-520">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4b285-521">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="4b285-521">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="4b285-522">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-522">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-523">利用不可</span><span class="sxs-lookup"><span data-stu-id="4b285-523">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-524">Android 上の Office</span><span class="sxs-lookup"><span data-stu-id="4b285-524">Office on Android</span></span><br><span data-ttu-id="4b285-525">(Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="4b285-525">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="4b285-526">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="4b285-526">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="4b285-527">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">予定の開催者 (作成): オンライン会議</a> (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="4b285-527">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Appointment Organizer (Compose): online meeting</a> (preview)</span></span><br><span data-ttu-id="4b285-528">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-528">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-529">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-529">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="4b285-530">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4b285-530">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="4b285-531">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4b285-531">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="4b285-532">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="4b285-532">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="4b285-533">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-533">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-534">利用不可</span><span class="sxs-lookup"><span data-stu-id="4b285-534">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="4b285-535">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="4b285-535">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="4b285-536">要件セットのクライアント サポートは、Exchange サーバー サポートの制限を受ける場合があります。</span><span class="sxs-lookup"><span data-stu-id="4b285-536">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="4b285-537">Exchange サーバーおよび Outlook クライアントによってサポートされている要件セットの範囲の詳細については、「[Outlook JavaScript APIの要件セット](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="4b285-537">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="4b285-538">Word</span><span class="sxs-lookup"><span data-stu-id="4b285-538">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4b285-539">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="4b285-539">Platform</span></span></th>
    <th><span data-ttu-id="4b285-540">拡張点</span><span class="sxs-lookup"><span data-stu-id="4b285-540">Extension points</span></span></th>
    <th><span data-ttu-id="4b285-541">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="4b285-541">API requirement sets</span></span></th>
    <th><span data-ttu-id="4b285-542"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="4b285-542"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-543">Office on the web</span><span class="sxs-lookup"><span data-stu-id="4b285-543">Office on the web</span></span></td>
    <td><span data-ttu-id="4b285-544">
      - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4b285-544">
      - TaskPane</span></span><br><span data-ttu-id="4b285-545">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-545">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-546">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-546">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="4b285-547">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4b285-547">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="4b285-548">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4b285-548">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="4b285-549">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-549">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4b285-550">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-550">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="4b285-551">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-551">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-552">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-552">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="4b285-553">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="4b285-553">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="4b285-554">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-554">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="4b285-555">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="4b285-555">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="4b285-556">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-556">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="4b285-557">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-557">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="4b285-558">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-558">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="4b285-559">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-559">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="4b285-560">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-560">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="4b285-561">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="4b285-561">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="4b285-562">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">設定</a></span><span class="sxs-lookup"><span data-stu-id="4b285-562">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="4b285-563">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-563">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="4b285-564">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-564">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="4b285-565">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-565">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="4b285-566">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-566">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="4b285-567">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-567">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-568">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="4b285-568">Office on Windows</span></span><br><span data-ttu-id="4b285-569">(Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="4b285-569">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="4b285-570">
      - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4b285-570">
      - TaskPane</span></span><br><span data-ttu-id="4b285-571">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-571">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-572">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-572">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="4b285-573">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4b285-573">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="4b285-574">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4b285-574">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="4b285-575">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-575">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4b285-576">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-576">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="4b285-577">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-577">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-578">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-578">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="4b285-579">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-579">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="4b285-580">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="4b285-580">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="4b285-581">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-581">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="4b285-582">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="4b285-582">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="4b285-583">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-583">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="4b285-584">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-584">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="4b285-585">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-585">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="4b285-586">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-586">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="4b285-587">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-587">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="4b285-588">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="4b285-588">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="4b285-589">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">設定</a></span><span class="sxs-lookup"><span data-stu-id="4b285-589">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="4b285-590">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-590">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="4b285-591">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-591">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="4b285-592">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-592">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="4b285-593">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-593">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="4b285-594">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-594">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-595">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="4b285-595">Office 2019 on Windows</span></span><br><span data-ttu-id="4b285-596">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4b285-596">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="4b285-597">
      - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4b285-597">
      - TaskPane</span></span><br><span data-ttu-id="4b285-598">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-598">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-599">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-599">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="4b285-600">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4b285-600">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="4b285-601">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4b285-601">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="4b285-602">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-602">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4b285-603">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-603">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-604">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-604">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="4b285-605">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-605">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="4b285-606">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="4b285-606">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="4b285-607">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-607">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="4b285-608">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="4b285-608">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="4b285-609">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-609">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="4b285-610">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-610">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="4b285-611">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-611">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="4b285-612">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-612">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="4b285-613">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-613">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="4b285-614">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="4b285-614">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="4b285-615">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">設定</a></span><span class="sxs-lookup"><span data-stu-id="4b285-615">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="4b285-616">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-616">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="4b285-617">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-617">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="4b285-618">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-618">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="4b285-619">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-619">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="4b285-620">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-620">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-621">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="4b285-621">Office 2016 on Windows</span></span><br><span data-ttu-id="4b285-622">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4b285-622">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="4b285-623">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4b285-623">- TaskPane</span></span></td>
    <td><span data-ttu-id="4b285-624">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-624">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="4b285-625">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="4b285-625">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="4b285-626">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-626">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-627">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-627">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="4b285-628">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-628">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="4b285-629">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="4b285-629">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="4b285-630">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-630">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="4b285-631">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="4b285-631">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="4b285-632">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-632">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="4b285-633">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-633">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="4b285-634">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-634">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="4b285-635">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-635">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="4b285-636">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-636">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="4b285-637">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="4b285-637">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="4b285-638">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">設定</a></span><span class="sxs-lookup"><span data-stu-id="4b285-638">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="4b285-639">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-639">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="4b285-640">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-640">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="4b285-641">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-641">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="4b285-642">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-642">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="4b285-643">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-643">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-644">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="4b285-644">Office 2013 on Windows</span></span><br><span data-ttu-id="4b285-645">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4b285-645">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="4b285-646">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4b285-646">- TaskPane</span></span></td>
    <td><span data-ttu-id="4b285-647">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="4b285-647">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="4b285-648">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-648">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-649">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-649">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="4b285-650">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-650">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="4b285-651">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="4b285-651">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="4b285-652">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-652">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="4b285-653">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="4b285-653">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="4b285-654">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-654">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="4b285-655">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-655">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="4b285-656">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-656">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="4b285-657">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-657">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="4b285-658">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-658">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="4b285-659">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="4b285-659">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="4b285-660">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">設定</a></span><span class="sxs-lookup"><span data-stu-id="4b285-660">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="4b285-661">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-661">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="4b285-662">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-662">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="4b285-663">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-663">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="4b285-664">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-664">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="4b285-665">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-665">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-666">Office on iPad</span><span class="sxs-lookup"><span data-stu-id="4b285-666">Office on iPad</span></span><br><span data-ttu-id="4b285-667">(Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="4b285-667">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="4b285-668">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4b285-668">- TaskPane</span></span></td>
    <td><span data-ttu-id="4b285-669">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-669">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="4b285-670">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4b285-670">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="4b285-671">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4b285-671">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="4b285-672">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-672">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4b285-673">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-673">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-674">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-674">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="4b285-675">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-675">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="4b285-676">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="4b285-676">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="4b285-677">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-677">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="4b285-678">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="4b285-678">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="4b285-679">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-679">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="4b285-680">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-680">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="4b285-681">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-681">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="4b285-682">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-682">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="4b285-683">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-683">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="4b285-684">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="4b285-684">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="4b285-685">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">設定</a></span><span class="sxs-lookup"><span data-stu-id="4b285-685">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="4b285-686">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-686">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="4b285-687">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-687">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="4b285-688">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-688">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="4b285-689">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-689">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="4b285-690">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-690">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-691">Office on Mac</span><span class="sxs-lookup"><span data-stu-id="4b285-691">Office on Mac</span></span><br><span data-ttu-id="4b285-692">(Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="4b285-692">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="4b285-693">
      - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4b285-693">
      - TaskPane</span></span><br><span data-ttu-id="4b285-694">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-694">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-695">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-695">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="4b285-696">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4b285-696">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="4b285-697">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4b285-697">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="4b285-698">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-698">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4b285-699">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-699">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="4b285-700">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-700">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-701">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-701">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="4b285-702">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-702">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="4b285-703">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="4b285-703">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="4b285-704">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-704">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="4b285-705">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="4b285-705">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="4b285-706">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-706">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="4b285-707">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-707">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="4b285-708">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-708">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="4b285-709">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-709">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="4b285-710">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-710">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="4b285-711">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="4b285-711">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="4b285-712">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">設定</a></span><span class="sxs-lookup"><span data-stu-id="4b285-712">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="4b285-713">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-713">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="4b285-714">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-714">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="4b285-715">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-715">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="4b285-716">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-716">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="4b285-717">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-717">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-718">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="4b285-718">Office 2019 on Mac</span></span><br><span data-ttu-id="4b285-719">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4b285-719">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="4b285-720">
      - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4b285-720">
      - TaskPane</span></span><br><span data-ttu-id="4b285-721">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-721">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-722">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-722">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="4b285-723">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="4b285-723">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="4b285-724">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="4b285-724">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="4b285-725">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-725">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4b285-726">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-726">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-727">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-727">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="4b285-728">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-728">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="4b285-729">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="4b285-729">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="4b285-730">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-730">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="4b285-731">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="4b285-731">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="4b285-732">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-732">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="4b285-733">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-733">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="4b285-734">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-734">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="4b285-735">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-735">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="4b285-736">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-736">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="4b285-737">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="4b285-737">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="4b285-738">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">設定</a></span><span class="sxs-lookup"><span data-stu-id="4b285-738">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="4b285-739">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-739">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="4b285-740">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-740">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="4b285-741">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-741">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="4b285-742">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-742">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="4b285-743">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-743">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-744">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="4b285-744">Office 2016 on Mac</span></span><br><span data-ttu-id="4b285-745">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4b285-745">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="4b285-746">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4b285-746">- TaskPane</span></span></td>
    <td><span data-ttu-id="4b285-747">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-747">
      - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="4b285-748">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="4b285-748">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="4b285-749">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-749">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-750">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-750">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#bindingevents">BindingEvents</a></span></span><br><span data-ttu-id="4b285-751">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-751">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="4b285-752">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span><span class="sxs-lookup"><span data-stu-id="4b285-752">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#customxmlparts">CustomXmlParts</a></span></span><br><span data-ttu-id="4b285-753">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-753">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="4b285-754">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="4b285-754">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="4b285-755">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-755">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="4b285-756">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-756">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixbindings">MatrixBindings</a></span></span><br><span data-ttu-id="4b285-757">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-757">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#matrixcoercion">MatrixCoercion</a></span></span><br><span data-ttu-id="4b285-758">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-758">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#ooxmlcoercion">OoxmlCoercion</a></span></span><br><span data-ttu-id="4b285-759">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-759">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="4b285-760">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="4b285-760">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="4b285-761">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">設定</a></span><span class="sxs-lookup"><span data-stu-id="4b285-761">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="4b285-762">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-762">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablebindings">TableBindings</a></span></span><br><span data-ttu-id="4b285-763">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-763">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#tablecoercion">TableCoercion</a></span></span><br><span data-ttu-id="4b285-764">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-764">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textbindings">TextBindings</a></span></span><br><span data-ttu-id="4b285-765">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-765">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a></span></span><br><span data-ttu-id="4b285-766">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-766">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textfile">TextFile</a>
    </span></span></td>
  </tr>
</table>

<span data-ttu-id="4b285-767">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="4b285-767">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="4b285-768">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="4b285-768">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4b285-769">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="4b285-769">Platform</span></span></th>
    <th><span data-ttu-id="4b285-770">拡張点</span><span class="sxs-lookup"><span data-stu-id="4b285-770">Extension points</span></span></th>
    <th><span data-ttu-id="4b285-771">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="4b285-771">API requirement sets</span></span></th>
    <th><span data-ttu-id="4b285-772"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="4b285-772"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-773">Office on the web</span><span class="sxs-lookup"><span data-stu-id="4b285-773">Office on the web</span></span></td>
    <td><span data-ttu-id="4b285-774">
      - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4b285-774">
      - Content</span></span><br><span data-ttu-id="4b285-775">
      - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4b285-775">
      - TaskPane</span></span><br><span data-ttu-id="4b285-776">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-776">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-777">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-777">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="4b285-778">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-778">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4b285-779">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-779">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="4b285-780">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-780">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-781">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="4b285-781">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="4b285-782">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-782">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="4b285-783">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-783">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="4b285-784">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="4b285-784">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="4b285-785">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-785">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="4b285-786">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="4b285-786">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="4b285-787">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-787">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="4b285-788">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-788">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-789">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="4b285-789">Office on Windows</span></span><br><span data-ttu-id="4b285-790">(Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="4b285-790">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="4b285-791">
      - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4b285-791">
      - Content</span></span><br><span data-ttu-id="4b285-792">
      - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4b285-792">
      - TaskPane</span></span><br><span data-ttu-id="4b285-793">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-793">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-794">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-794">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="4b285-795">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-795">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4b285-796">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-796">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="4b285-797">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-797">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-798">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="4b285-798">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="4b285-799">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-799">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="4b285-800">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-800">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="4b285-801">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="4b285-801">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="4b285-802">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-802">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="4b285-803">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="4b285-803">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="4b285-804">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-804">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="4b285-805">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-805">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-806">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="4b285-806">Office 2019 on Windows</span></span><br><span data-ttu-id="4b285-807">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4b285-807">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="4b285-808">
      - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4b285-808">
      - Content</span></span><br><span data-ttu-id="4b285-809">
      - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4b285-809">
      - TaskPane</span></span><br><span data-ttu-id="4b285-810">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-810">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-811">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-811">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4b285-812">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-812">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-813">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="4b285-813">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="4b285-814">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-814">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="4b285-815">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-815">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="4b285-816">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="4b285-816">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="4b285-817">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-817">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="4b285-818">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="4b285-818">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="4b285-819">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-819">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="4b285-820">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-820">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-821">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="4b285-821">Office 2016 on Windows</span></span><br><span data-ttu-id="4b285-822">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4b285-822">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="4b285-823">
      - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4b285-823">
      - Content</span></span><br><span data-ttu-id="4b285-824">
      - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4b285-824">
      - TaskPane</span></span> </td>
    <td><span data-ttu-id="4b285-825">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="4b285-825">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="4b285-826">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-826">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-827">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="4b285-827">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="4b285-828">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-828">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="4b285-829">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-829">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="4b285-830">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="4b285-830">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="4b285-831">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-831">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="4b285-832">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="4b285-832">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="4b285-833">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-833">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="4b285-834">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-834">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-835">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="4b285-835">Office 2013 on Windows</span></span><br><span data-ttu-id="4b285-836">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4b285-836">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="4b285-837">
      - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4b285-837">
      - Content</span></span><br><span data-ttu-id="4b285-838">
      - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4b285-838">
      - TaskPane</span></span> </td>
    <td><span data-ttu-id="4b285-839">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="4b285-839">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="4b285-840">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-840">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-841">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="4b285-841">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="4b285-842">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-842">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="4b285-843">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-843">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="4b285-844">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="4b285-844">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="4b285-845">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-845">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="4b285-846">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="4b285-846">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="4b285-847">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-847">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="4b285-848">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-848">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-849">Office on iPad</span><span class="sxs-lookup"><span data-stu-id="4b285-849">Office on iPad</span></span><br><span data-ttu-id="4b285-850">(Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="4b285-850">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="4b285-851">
      - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4b285-851">
      - Content</span></span><br><span data-ttu-id="4b285-852">
      - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4b285-852">
      - TaskPane</span></span> </td>
    <td><span data-ttu-id="4b285-853">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-853">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="4b285-854">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-854">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4b285-855">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-855">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-856">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="4b285-856">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="4b285-857">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-857">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="4b285-858">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-858">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="4b285-859">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="4b285-859">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="4b285-860">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-860">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="4b285-861">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="4b285-861">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="4b285-862">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-862">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="4b285-863">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-863">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-864">Office on Mac</span><span class="sxs-lookup"><span data-stu-id="4b285-864">Office on Mac</span></span><br><span data-ttu-id="4b285-865">(Microsoft 365 サブスクリプションに接続)</span><span class="sxs-lookup"><span data-stu-id="4b285-865">(connected to a Microsoft 365 subscription)</span></span></td>
    <td><span data-ttu-id="4b285-866">
      - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4b285-866">
      - Content</span></span><br><span data-ttu-id="4b285-867">
      - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4b285-867">
      - TaskPane</span></span><br><span data-ttu-id="4b285-868">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-868">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-869">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-869">
      - <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="4b285-870">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-870">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4b285-871">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-871">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="4b285-872">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-872">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-873">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="4b285-873">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="4b285-874">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-874">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="4b285-875">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-875">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="4b285-876">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="4b285-876">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="4b285-877">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-877">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="4b285-878">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="4b285-878">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="4b285-879">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-879">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="4b285-880">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-880">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-881">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="4b285-881">Office 2019 on Mac</span></span><br><span data-ttu-id="4b285-882">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4b285-882">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="4b285-883">
      - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4b285-883">
      - Content</span></span><br><span data-ttu-id="4b285-884">
      - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4b285-884">
      - TaskPane</span></span><br><span data-ttu-id="4b285-885">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-885">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-886">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-886">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4b285-887">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-887">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-888">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="4b285-888">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="4b285-889">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-889">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="4b285-890">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-890">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="4b285-891">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="4b285-891">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="4b285-892">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-892">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="4b285-893">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="4b285-893">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="4b285-894">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-894">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="4b285-895">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-895">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-896">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="4b285-896">Office 2016 on Mac</span></span><br><span data-ttu-id="4b285-897">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4b285-897">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="4b285-898">
      - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4b285-898">
      - Content</span></span><br><span data-ttu-id="4b285-899">
      - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4b285-899">
      - TaskPane</span></span> </td>
    <td><span data-ttu-id="4b285-900">
       - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="4b285-900">
       - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="4b285-901">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-901">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-902">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span><span class="sxs-lookup"><span data-stu-id="4b285-902">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#activeview">ActiveView</a></span></span><br><span data-ttu-id="4b285-903">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-903">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#compressedfile">CompressedFile</a></span></span><br><span data-ttu-id="4b285-904">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-904">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="4b285-905">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span><span class="sxs-lookup"><span data-stu-id="4b285-905">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#file">File</a></span></span><br><span data-ttu-id="4b285-906">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span><span class="sxs-lookup"><span data-stu-id="4b285-906">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#pdffile">PdfFile</a></span></span><br><span data-ttu-id="4b285-907">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="4b285-907">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="4b285-908">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-908">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="4b285-909">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-909">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
</table>

<span data-ttu-id="4b285-910">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="4b285-910">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="4b285-911">OneNote</span><span class="sxs-lookup"><span data-stu-id="4b285-911">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4b285-912">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="4b285-912">Platform</span></span></th>
    <th><span data-ttu-id="4b285-913">拡張点</span><span class="sxs-lookup"><span data-stu-id="4b285-913">Extension points</span></span></th>
    <th><span data-ttu-id="4b285-914">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="4b285-914">API requirement sets</span></span></th>
    <th><span data-ttu-id="4b285-915"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="4b285-915"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-916">Office on the web</span><span class="sxs-lookup"><span data-stu-id="4b285-916">Office on the web</span></span></td>
    <td><span data-ttu-id="4b285-917">
      - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="4b285-917">
      - Content</span></span><br><span data-ttu-id="4b285-918">
      - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4b285-918">
      - TaskPane</span></span><br><span data-ttu-id="4b285-919">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-919">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-920">
      - <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-920">
      - <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="4b285-921">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-921">
      - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="4b285-922">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-922">
      - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a>
    </span></span></td>
    <td><span data-ttu-id="4b285-923">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span><span class="sxs-lookup"><span data-stu-id="4b285-923">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#documentevents">DocumentEvents</a></span></span><br><span data-ttu-id="4b285-924">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span><span class="sxs-lookup"><span data-stu-id="4b285-924">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#htmlcoercion">HtmlCoercion</a></span></span><br><span data-ttu-id="4b285-925">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span><span class="sxs-lookup"><span data-stu-id="4b285-925">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#settings">Settings</a></span></span><br><span data-ttu-id="4b285-926">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-926">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="4b285-927">Project</span><span class="sxs-lookup"><span data-stu-id="4b285-927">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="4b285-928">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="4b285-928">Platform</span></span></th>
    <th><span data-ttu-id="4b285-929">拡張点</span><span class="sxs-lookup"><span data-stu-id="4b285-929">Extension points</span></span></th>
    <th><span data-ttu-id="4b285-930">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="4b285-930">API requirement sets</span></span></th>
    <th><span data-ttu-id="4b285-931"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="4b285-931"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-932">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="4b285-932">Office 2019 on Windows</span></span><br><span data-ttu-id="4b285-933">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4b285-933">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="4b285-934">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4b285-934">- TaskPane</span></span></td>
    <td><span data-ttu-id="4b285-935">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-935">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4b285-936">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="4b285-936">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="4b285-937">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-937">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-938">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="4b285-938">Office 2016 on Windows</span></span><br><span data-ttu-id="4b285-939">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4b285-939">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="4b285-940">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4b285-940">- TaskPane</span></span></td>
    <td><span data-ttu-id="4b285-941">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-941">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4b285-942">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="4b285-942">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="4b285-943">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-943">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="4b285-944">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="4b285-944">Office 2013 on Windows</span></span><br><span data-ttu-id="4b285-945">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="4b285-945">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="4b285-946">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="4b285-946">- TaskPane</span></span></td>
    <td><span data-ttu-id="4b285-947">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="4b285-947">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="4b285-948">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span><span class="sxs-lookup"><span data-stu-id="4b285-948">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#selection">Selection</a></span></span><br><span data-ttu-id="4b285-949">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span><span class="sxs-lookup"><span data-stu-id="4b285-949">
      - <a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets#textcoercion">TextCoercion</a>
    </span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="4b285-950">関連項目</span><span class="sxs-lookup"><span data-stu-id="4b285-950">See also</span></span>

- [<span data-ttu-id="4b285-951">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="4b285-951">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="4b285-952">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="4b285-952">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="4b285-953">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="4b285-953">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="4b285-954">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="4b285-954">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="4b285-955">API リファレンス ドキュメント</span><span class="sxs-lookup"><span data-stu-id="4b285-955">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="4b285-956">Microsoft 365 Apps の更新履歴</span><span class="sxs-lookup"><span data-stu-id="4b285-956">Update history for Microsoft 365 Apps</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="4b285-957">Office 2016および2019の更新履歴（クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="4b285-957">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="4b285-958">Office 2013 の更新履歴 （クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="4b285-958">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="4b285-959">Office 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="4b285-959">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="4b285-960">Outlook 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="4b285-960">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="4b285-961">Office for Mac の更新履歴</span><span class="sxs-lookup"><span data-stu-id="4b285-961">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="4b285-962">Office アドインを開発する</span><span class="sxs-lookup"><span data-stu-id="4b285-962">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
