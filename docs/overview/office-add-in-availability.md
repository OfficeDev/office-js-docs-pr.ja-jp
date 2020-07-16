---
title: Office アドインを使用できるホストおよびプラットフォーム
description: Excel、OneNote、Outlook、PowerPoint、Project、Word のサポートされる要件セット。
ms.date: 06/23/2020
localization_priority: Priority
ms.openlocfilehash: 979c873b1c5f2d1d7847414f037d5c75737aa33d
ms.sourcegitcommit: a4873c3525c7d30ef551545d27eb2c0a16b4eb50
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/25/2020
ms.locfileid: "44888160"
---
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="ca8cd-103">Office アドインを使用できるホストおよびプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="ca8cd-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="ca8cd-p101">期待どおりの動作をするうえで、Office アドインは特定の Office ホスト、要件セット、API メンバー、または API のバージョンに依存することがあります。次の表には、使用可能なプラットフォーム、拡張点、API 要件セット、および各 Office アプリケーションで現在サポートされている共通 API が含まれています。</span><span class="sxs-lookup"><span data-stu-id="ca8cd-p101">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API. The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="ca8cd-106">MSI からインストールされる最初の Office 2016 リリースには、ExcelApi 1.1、WordApi 1.1、共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="ca8cd-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="ca8cd-107">さまざまなバージョンの Office の更新履歴の詳細については、「[関連項目](#see-also)」セクションをご確認ください。</span><span class="sxs-lookup"><span data-stu-id="ca8cd-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="ca8cd-108">Excel</span><span class="sxs-lookup"><span data-stu-id="ca8cd-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="ca8cd-109">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="ca8cd-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="ca8cd-110">拡張点</span><span class="sxs-lookup"><span data-stu-id="ca8cd-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="ca8cd-111">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="ca8cd-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="ca8cd-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-113">Office on the web</span><span class="sxs-lookup"><span data-stu-id="ca8cd-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="ca8cd-114">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-114">- TaskPane</span></span><br><span data-ttu-id="ca8cd-115">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-115">
        - Content</span></span><br><span data-ttu-id="ca8cd-116">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="ca8cd-116">
        - Custom Functions</span></span><br><span data-ttu-id="ca8cd-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="ca8cd-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ca8cd-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ca8cd-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ca8cd-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ca8cd-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ca8cd-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ca8cd-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ca8cd-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ca8cd-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="ca8cd-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="ca8cd-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="ca8cd-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="ca8cd-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="ca8cd-131">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-131">
        - BindingEvents</span></span><br><span data-ttu-id="ca8cd-132">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-132">
        - CompressedFile</span></span><br><span data-ttu-id="ca8cd-133">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-133">
        - DocumentEvents</span></span><br><span data-ttu-id="ca8cd-134">
        - File</span><span class="sxs-lookup"><span data-stu-id="ca8cd-134">
        - File</span></span><br><span data-ttu-id="ca8cd-135">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-135">
        - MatrixBindings</span></span><br><span data-ttu-id="ca8cd-136">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-136">
        - MatrixCoercion</span></span><br><span data-ttu-id="ca8cd-137">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ca8cd-137">
        - Selection</span></span><br><span data-ttu-id="ca8cd-138">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-138">
        - Settings</span></span><br><span data-ttu-id="ca8cd-139">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-139">
        - TableBindings</span></span><br><span data-ttu-id="ca8cd-140">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-140">
        - TableCoercion</span></span><br><span data-ttu-id="ca8cd-141">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-141">
        - TextBindings</span></span><br><span data-ttu-id="ca8cd-142">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-142">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-143">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="ca8cd-143">Office on Windows</span></span><br><span data-ttu-id="ca8cd-144">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-144">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ca8cd-145">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-145">- TaskPane</span></span><br><span data-ttu-id="ca8cd-146">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-146">
        - Content</span></span><br><span data-ttu-id="ca8cd-147">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="ca8cd-147">
        - Custom Functions</span></span><br><span data-ttu-id="ca8cd-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="ca8cd-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="ca8cd-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ca8cd-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ca8cd-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ca8cd-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ca8cd-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ca8cd-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ca8cd-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ca8cd-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="ca8cd-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="ca8cd-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="ca8cd-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ca8cd-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="ca8cd-163">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-163">
        - BindingEvents</span></span><br><span data-ttu-id="ca8cd-164">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-164">
        - CompressedFile</span></span><br><span data-ttu-id="ca8cd-165">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-165">
        - DocumentEvents</span></span><br><span data-ttu-id="ca8cd-166">
        - File</span><span class="sxs-lookup"><span data-stu-id="ca8cd-166">
        - File</span></span><br><span data-ttu-id="ca8cd-167">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-167">
        - MatrixBindings</span></span><br><span data-ttu-id="ca8cd-168">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-168">
        - MatrixCoercion</span></span><br><span data-ttu-id="ca8cd-169">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ca8cd-169">
        - Selection</span></span><br><span data-ttu-id="ca8cd-170">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-170">
        - Settings</span></span><br><span data-ttu-id="ca8cd-171">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-171">
        - TableBindings</span></span><br><span data-ttu-id="ca8cd-172">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-172">
        - TableCoercion</span></span><br><span data-ttu-id="ca8cd-173">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-173">
        - TextBindings</span></span><br><span data-ttu-id="ca8cd-174">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-174">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-175">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="ca8cd-175">Office 2019 on Windows</span></span><br><span data-ttu-id="ca8cd-176">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-176">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ca8cd-177">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-177">- TaskPane</span></span><br><span data-ttu-id="ca8cd-178">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-178">
        - Content</span></span><br><span data-ttu-id="ca8cd-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="ca8cd-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ca8cd-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ca8cd-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ca8cd-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ca8cd-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ca8cd-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ca8cd-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ca8cd-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ca8cd-190">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-190">- BindingEvents</span></span><br><span data-ttu-id="ca8cd-191">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-191">
        - CompressedFile</span></span><br><span data-ttu-id="ca8cd-192">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-192">
        - DocumentEvents</span></span><br><span data-ttu-id="ca8cd-193">
        - File</span><span class="sxs-lookup"><span data-stu-id="ca8cd-193">
        - File</span></span><br><span data-ttu-id="ca8cd-194">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-194">
        - MatrixBindings</span></span><br><span data-ttu-id="ca8cd-195">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-195">
        - MatrixCoercion</span></span><br><span data-ttu-id="ca8cd-196">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ca8cd-196">
        - Selection</span></span><br><span data-ttu-id="ca8cd-197">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-197">
        - Settings</span></span><br><span data-ttu-id="ca8cd-198">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-198">
        - TableBindings</span></span><br><span data-ttu-id="ca8cd-199">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-199">
        - TableCoercion</span></span><br><span data-ttu-id="ca8cd-200">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-200">
        - TextBindings</span></span><br><span data-ttu-id="ca8cd-201">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-201">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-202">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="ca8cd-202">Office 2016 on Windows</span></span><br><span data-ttu-id="ca8cd-203">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-203">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ca8cd-204">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-204">- TaskPane</span></span><br><span data-ttu-id="ca8cd-205">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-205">
        - Content</span></span></td>
    <td><span data-ttu-id="ca8cd-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="ca8cd-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="ca8cd-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ca8cd-209">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-209">- BindingEvents</span></span><br><span data-ttu-id="ca8cd-210">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-210">
        - CompressedFile</span></span><br><span data-ttu-id="ca8cd-211">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-211">
        - DocumentEvents</span></span><br><span data-ttu-id="ca8cd-212">
        - File</span><span class="sxs-lookup"><span data-stu-id="ca8cd-212">
        - File</span></span><br><span data-ttu-id="ca8cd-213">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-213">
        - MatrixBindings</span></span><br><span data-ttu-id="ca8cd-214">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-214">
        - MatrixCoercion</span></span><br><span data-ttu-id="ca8cd-215">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ca8cd-215">
        - Selection</span></span><br><span data-ttu-id="ca8cd-216">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-216">
        - Settings</span></span><br><span data-ttu-id="ca8cd-217">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-217">
        - TableBindings</span></span><br><span data-ttu-id="ca8cd-218">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-218">
        - TableCoercion</span></span><br><span data-ttu-id="ca8cd-219">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-219">
        - TextBindings</span></span><br><span data-ttu-id="ca8cd-220">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-220">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-221">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="ca8cd-221">Office 2013 on Windows</span></span><br><span data-ttu-id="ca8cd-222">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-222">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ca8cd-223">
        - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-223">
        - TaskPane</span></span><br><span data-ttu-id="ca8cd-224">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-224">
        - Content</span></span></td>
    <td>  <span data-ttu-id="ca8cd-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="ca8cd-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="ca8cd-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ca8cd-227">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-227">
        - BindingEvents</span></span><br><span data-ttu-id="ca8cd-228">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-228">
        - DocumentEvents</span></span><br><span data-ttu-id="ca8cd-229">
        - File</span><span class="sxs-lookup"><span data-stu-id="ca8cd-229">
        - File</span></span><br><span data-ttu-id="ca8cd-230">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-230">
        - MatrixBindings</span></span><br><span data-ttu-id="ca8cd-231">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-231">
        - MatrixCoercion</span></span><br><span data-ttu-id="ca8cd-232">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ca8cd-232">
        - Selection</span></span><br><span data-ttu-id="ca8cd-233">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-233">
        - Settings</span></span><br><span data-ttu-id="ca8cd-234">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-234">
        - TableBindings</span></span><br><span data-ttu-id="ca8cd-235">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-235">
        - TableCoercion</span></span><br><span data-ttu-id="ca8cd-236">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-236">
        - TextBindings</span></span><br><span data-ttu-id="ca8cd-237">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-237">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-238">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="ca8cd-238">Office on iPad</span></span><br><span data-ttu-id="ca8cd-239">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-239">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="ca8cd-240">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-240">- TaskPane</span></span><br><span data-ttu-id="ca8cd-241">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-241">
        - Content</span></span></td>
    <td><span data-ttu-id="ca8cd-242">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-242">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ca8cd-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ca8cd-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ca8cd-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ca8cd-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ca8cd-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ca8cd-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ca8cd-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="ca8cd-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="ca8cd-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="ca8cd-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ca8cd-255">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-255">- BindingEvents</span></span><br><span data-ttu-id="ca8cd-256">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-256">
        - DocumentEvents</span></span><br><span data-ttu-id="ca8cd-257">
        - File</span><span class="sxs-lookup"><span data-stu-id="ca8cd-257">
        - File</span></span><br><span data-ttu-id="ca8cd-258">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-258">
        - MatrixBindings</span></span><br><span data-ttu-id="ca8cd-259">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-259">
        - MatrixCoercion</span></span><br><span data-ttu-id="ca8cd-260">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ca8cd-260">
        - Selection</span></span><br><span data-ttu-id="ca8cd-261">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-261">
        - Settings</span></span><br><span data-ttu-id="ca8cd-262">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-262">
        - TableBindings</span></span><br><span data-ttu-id="ca8cd-263">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-263">
        - TableCoercion</span></span><br><span data-ttu-id="ca8cd-264">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-264">
        - TextBindings</span></span><br><span data-ttu-id="ca8cd-265">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-265">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-266">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="ca8cd-266">Office on Mac</span></span><br><span data-ttu-id="ca8cd-267">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-267">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="ca8cd-268">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-268">- TaskPane</span></span><br><span data-ttu-id="ca8cd-269">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-269">
        - Content</span></span><br><span data-ttu-id="ca8cd-270">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="ca8cd-270">
        - Custom Functions</span></span><br><span data-ttu-id="ca8cd-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="ca8cd-272">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-272">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ca8cd-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ca8cd-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ca8cd-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ca8cd-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ca8cd-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ca8cd-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ca8cd-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="ca8cd-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="ca8cd-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="ca8cd-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ca8cd-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="ca8cd-286">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-286">- BindingEvents</span></span><br><span data-ttu-id="ca8cd-287">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-287">
        - CompressedFile</span></span><br><span data-ttu-id="ca8cd-288">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-288">
        - DocumentEvents</span></span><br><span data-ttu-id="ca8cd-289">
        - File</span><span class="sxs-lookup"><span data-stu-id="ca8cd-289">
        - File</span></span><br><span data-ttu-id="ca8cd-290">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-290">
        - MatrixBindings</span></span><br><span data-ttu-id="ca8cd-291">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-291">
        - MatrixCoercion</span></span><br><span data-ttu-id="ca8cd-292">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-292">
        - PdfFile</span></span><br><span data-ttu-id="ca8cd-293">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ca8cd-293">
        - Selection</span></span><br><span data-ttu-id="ca8cd-294">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-294">
        - Settings</span></span><br><span data-ttu-id="ca8cd-295">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-295">
        - TableBindings</span></span><br><span data-ttu-id="ca8cd-296">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-296">
        - TableCoercion</span></span><br><span data-ttu-id="ca8cd-297">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-297">
        - TextBindings</span></span><br><span data-ttu-id="ca8cd-298">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-298">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-299">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="ca8cd-299">Office 2019 on Mac</span></span><br><span data-ttu-id="ca8cd-300">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-300">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ca8cd-301">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-301">- TaskPane</span></span><br><span data-ttu-id="ca8cd-302">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-302">
        - Content</span></span><br><span data-ttu-id="ca8cd-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="ca8cd-304">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-304">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="ca8cd-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="ca8cd-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="ca8cd-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="ca8cd-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="ca8cd-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="ca8cd-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="ca8cd-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ca8cd-314">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-314">- BindingEvents</span></span><br><span data-ttu-id="ca8cd-315">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-315">
        - CompressedFile</span></span><br><span data-ttu-id="ca8cd-316">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-316">
        - DocumentEvents</span></span><br><span data-ttu-id="ca8cd-317">
        - File</span><span class="sxs-lookup"><span data-stu-id="ca8cd-317">
        - File</span></span><br><span data-ttu-id="ca8cd-318">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-318">
        - MatrixBindings</span></span><br><span data-ttu-id="ca8cd-319">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-319">
        - MatrixCoercion</span></span><br><span data-ttu-id="ca8cd-320">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-320">
        - PdfFile</span></span><br><span data-ttu-id="ca8cd-321">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ca8cd-321">
        - Selection</span></span><br><span data-ttu-id="ca8cd-322">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-322">
        - Settings</span></span><br><span data-ttu-id="ca8cd-323">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-323">
        - TableBindings</span></span><br><span data-ttu-id="ca8cd-324">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-324">
        - TableCoercion</span></span><br><span data-ttu-id="ca8cd-325">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-325">
        - TextBindings</span></span><br><span data-ttu-id="ca8cd-326">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-326">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-327">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="ca8cd-327">Office 2016 on Mac</span></span><br><span data-ttu-id="ca8cd-328">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-328">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="ca8cd-329">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-329">- TaskPane</span></span><br><span data-ttu-id="ca8cd-330">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-330">
        - Content</span></span></td>
    <td><span data-ttu-id="ca8cd-331">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-331">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-332">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="ca8cd-332">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="ca8cd-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="ca8cd-334">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-334">- BindingEvents</span></span><br><span data-ttu-id="ca8cd-335">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-335">
        - CompressedFile</span></span><br><span data-ttu-id="ca8cd-336">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-336">
        - DocumentEvents</span></span><br><span data-ttu-id="ca8cd-337">
        - File</span><span class="sxs-lookup"><span data-stu-id="ca8cd-337">
        - File</span></span><br><span data-ttu-id="ca8cd-338">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-338">
        - MatrixBindings</span></span><br><span data-ttu-id="ca8cd-339">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-339">
        - MatrixCoercion</span></span><br><span data-ttu-id="ca8cd-340">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-340">
        - PdfFile</span></span><br><span data-ttu-id="ca8cd-341">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="ca8cd-341">
        - Selection</span></span><br><span data-ttu-id="ca8cd-342">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-342">
        - Settings</span></span><br><span data-ttu-id="ca8cd-343">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-343">
        - TableBindings</span></span><br><span data-ttu-id="ca8cd-344">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-344">
        - TableCoercion</span></span><br><span data-ttu-id="ca8cd-345">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-345">
        - TextBindings</span></span><br><span data-ttu-id="ca8cd-346">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-346">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="ca8cd-347">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="ca8cd-347">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="ca8cd-348">カスタム関数 (Excel のみ)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-348">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="ca8cd-349">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="ca8cd-349">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="ca8cd-350">拡張点</span><span class="sxs-lookup"><span data-stu-id="ca8cd-350">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="ca8cd-351">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="ca8cd-351">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="ca8cd-352"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-352"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-353">Office on the web</span><span class="sxs-lookup"><span data-stu-id="ca8cd-353">Office on the web</span></span></td>
    <td><span data-ttu-id="ca8cd-354">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="ca8cd-354">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="ca8cd-355">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-355">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-356">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="ca8cd-356">Office on Windows</span></span><br><span data-ttu-id="ca8cd-357">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-357">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="ca8cd-358">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="ca8cd-358">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="ca8cd-359">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-359">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-360">Office on Mac</span><span class="sxs-lookup"><span data-stu-id="ca8cd-360">Office on Mac</span></span><br><span data-ttu-id="ca8cd-361">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-361">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="ca8cd-362">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="ca8cd-362">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="ca8cd-363">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-363">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="ca8cd-364">Outlook</span><span class="sxs-lookup"><span data-stu-id="ca8cd-364">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ca8cd-365">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="ca8cd-365">Platform</span></span></th>
    <th><span data-ttu-id="ca8cd-366">拡張点</span><span class="sxs-lookup"><span data-stu-id="ca8cd-366">Extension points</span></span></th>
    <th><span data-ttu-id="ca8cd-367">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="ca8cd-367">API requirement sets</span></span></th>
    <th><span data-ttu-id="ca8cd-368"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-368"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-369">Office on the web</span><span class="sxs-lookup"><span data-stu-id="ca8cd-369">Office on the web</span></span><br><span data-ttu-id="ca8cd-370">(モダン)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-370">(modern)</span></span></td>
    <td> <span data-ttu-id="ca8cd-371">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-371">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ca8cd-372">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-372">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="ca8cd-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="ca8cd-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="ca8cd-375">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-375">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-376">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-376">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ca8cd-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ca8cd-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ca8cd-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ca8cd-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ca8cd-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="ca8cd-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="ca8cd-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="ca8cd-384">利用不可</span><span class="sxs-lookup"><span data-stu-id="ca8cd-384">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-385">Office on the web</span><span class="sxs-lookup"><span data-stu-id="ca8cd-385">Office on the web</span></span><br><span data-ttu-id="ca8cd-386">(クラシック)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-386">(classic)</span></span></td>
    <td> <span data-ttu-id="ca8cd-387">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-387">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ca8cd-388">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-388">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="ca8cd-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="ca8cd-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="ca8cd-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ca8cd-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ca8cd-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ca8cd-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ca8cd-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ca8cd-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="ca8cd-398">使用不可</span><span class="sxs-lookup"><span data-stu-id="ca8cd-398">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-399">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="ca8cd-399">Office on Windows</span></span><br><span data-ttu-id="ca8cd-400">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-400">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ca8cd-401">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-401">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ca8cd-402">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-402">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="ca8cd-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="ca8cd-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="ca8cd-405">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-405">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="ca8cd-406">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">モジュール</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-406">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-407">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-407">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ca8cd-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ca8cd-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ca8cd-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ca8cd-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ca8cd-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="ca8cd-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="ca8cd-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="ca8cd-415">利用不可</span><span class="sxs-lookup"><span data-stu-id="ca8cd-415">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-416">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="ca8cd-416">Office 2019 on Windows</span></span><br><span data-ttu-id="ca8cd-417">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-417">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ca8cd-418">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-418">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ca8cd-419">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-419">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="ca8cd-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="ca8cd-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="ca8cd-422">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-422">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="ca8cd-423">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">モジュール</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-423">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-424">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-424">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ca8cd-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ca8cd-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ca8cd-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ca8cd-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ca8cd-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="ca8cd-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="ca8cd-431">使用不可</span><span class="sxs-lookup"><span data-stu-id="ca8cd-431">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-432">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="ca8cd-432">Office 2016 on Windows</span></span><br><span data-ttu-id="ca8cd-433">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-433">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ca8cd-434">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-434">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ca8cd-435">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-435">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="ca8cd-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="ca8cd-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="ca8cd-438">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-438">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="ca8cd-439">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">モジュール</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-439">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-440">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-440">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ca8cd-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ca8cd-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ca8cd-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="ca8cd-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="ca8cd-444">使用不可</span><span class="sxs-lookup"><span data-stu-id="ca8cd-444">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-445">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="ca8cd-445">Office 2013 on Windows</span></span><br><span data-ttu-id="ca8cd-446">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-446">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ca8cd-447">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-447">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ca8cd-448">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-448">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="ca8cd-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="ca8cd-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br>
    <td> <span data-ttu-id="ca8cd-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ca8cd-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ca8cd-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="ca8cd-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="ca8cd-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="ca8cd-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="ca8cd-455">使用不可</span><span class="sxs-lookup"><span data-stu-id="ca8cd-455">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-456">iOS 上の Office</span><span class="sxs-lookup"><span data-stu-id="ca8cd-456">Office on iOS</span></span><br><span data-ttu-id="ca8cd-457">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-457">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ca8cd-458">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-458">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ca8cd-459">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-459">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-460">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-460">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ca8cd-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ca8cd-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ca8cd-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ca8cd-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="ca8cd-465">使用不可</span><span class="sxs-lookup"><span data-stu-id="ca8cd-465">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-466">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="ca8cd-466">Office on Mac</span></span><br><span data-ttu-id="ca8cd-467">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-467">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ca8cd-468">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-468">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ca8cd-469">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-469">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="ca8cd-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="ca8cd-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="ca8cd-472">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-472">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-473">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-473">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ca8cd-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ca8cd-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ca8cd-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ca8cd-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ca8cd-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="ca8cd-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="ca8cd-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="ca8cd-481">利用不可</span><span class="sxs-lookup"><span data-stu-id="ca8cd-481">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-482">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="ca8cd-482">Office 2019 on Mac</span></span><br><span data-ttu-id="ca8cd-483">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-483">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ca8cd-484">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-484">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ca8cd-485">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-485">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="ca8cd-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="ca8cd-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="ca8cd-488">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-488">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-489">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-489">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ca8cd-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ca8cd-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ca8cd-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ca8cd-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ca8cd-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="ca8cd-495">使用不可</span><span class="sxs-lookup"><span data-stu-id="ca8cd-495">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-496">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="ca8cd-496">Office 2016 on Mac</span></span><br><span data-ttu-id="ca8cd-497">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-497">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ca8cd-498">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-498">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ca8cd-499">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-499">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="ca8cd-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="ca8cd-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="ca8cd-502">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-502">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-503">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-503">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ca8cd-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ca8cd-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ca8cd-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ca8cd-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="ca8cd-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="ca8cd-509">使用不可</span><span class="sxs-lookup"><span data-stu-id="ca8cd-509">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-510">Android 上の Office</span><span class="sxs-lookup"><span data-stu-id="ca8cd-510">Office on Android</span></span><br><span data-ttu-id="ca8cd-511">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-511">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ca8cd-512">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-512">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="ca8cd-513">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">予定の開催者 (作成): オンライン会議</a> (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-513">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Appointment Organizer (Compose): online meeting</a> (preview)</span></span><br><span data-ttu-id="ca8cd-514">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-514">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-515">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-515">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="ca8cd-516">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-516">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="ca8cd-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="ca8cd-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="ca8cd-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="ca8cd-520">利用不可</span><span class="sxs-lookup"><span data-stu-id="ca8cd-520">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="ca8cd-521">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="ca8cd-521">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ca8cd-522">要件セットのクライアント サポートは、Exchange サーバー サポートの制限を受ける場合があります。</span><span class="sxs-lookup"><span data-stu-id="ca8cd-522">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="ca8cd-523">Exchange サーバーおよび Outlook クライアントによってサポートされている要件セットの範囲の詳細については、「[Outlook JavaScript APIの要件セット](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ca8cd-523">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="ca8cd-524">Word</span><span class="sxs-lookup"><span data-stu-id="ca8cd-524">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ca8cd-525">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="ca8cd-525">Platform</span></span></th>
    <th><span data-ttu-id="ca8cd-526">拡張点</span><span class="sxs-lookup"><span data-stu-id="ca8cd-526">Extension points</span></span></th>
    <th><span data-ttu-id="ca8cd-527">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="ca8cd-527">API requirement sets</span></span></th>
    <th><span data-ttu-id="ca8cd-528"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-528"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-529">Office on the web</span><span class="sxs-lookup"><span data-stu-id="ca8cd-529">Office on the web</span></span></td>
    <td> <span data-ttu-id="ca8cd-530">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-530">- TaskPane</span></span><br><span data-ttu-id="ca8cd-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-532">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-532">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ca8cd-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ca8cd-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ca8cd-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-538">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-538">- BindingEvents</span></span><br><span data-ttu-id="ca8cd-539">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ca8cd-539">
         - CustomXmlParts</span></span><br><span data-ttu-id="ca8cd-540">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-540">
         - DocumentEvents</span></span><br><span data-ttu-id="ca8cd-541">
         - File</span><span class="sxs-lookup"><span data-stu-id="ca8cd-541">
         - File</span></span><br><span data-ttu-id="ca8cd-542">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-542">
         - HtmlCoercion</span></span><br><span data-ttu-id="ca8cd-543">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-543">
         - MatrixBindings</span></span><br><span data-ttu-id="ca8cd-544">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-544">
         - MatrixCoercion</span></span><br><span data-ttu-id="ca8cd-545">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-545">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ca8cd-546">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-546">
         - PdfFile</span></span><br><span data-ttu-id="ca8cd-547">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ca8cd-547">
         - Selection</span></span><br><span data-ttu-id="ca8cd-548">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-548">
         - Settings</span></span><br><span data-ttu-id="ca8cd-549">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-549">
         - TableBindings</span></span><br><span data-ttu-id="ca8cd-550">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-550">
         - TableCoercion</span></span><br><span data-ttu-id="ca8cd-551">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-551">
         - TextBindings</span></span><br><span data-ttu-id="ca8cd-552">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-552">
         - TextCoercion</span></span><br><span data-ttu-id="ca8cd-553">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-553">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-554">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="ca8cd-554">Office on Windows</span></span><br><span data-ttu-id="ca8cd-555">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-555">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ca8cd-556">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-556">- TaskPane</span></span><br><span data-ttu-id="ca8cd-557">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-557">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-558">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-558">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ca8cd-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ca8cd-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-562">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-562">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ca8cd-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-564">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-564">- BindingEvents</span></span><br><span data-ttu-id="ca8cd-565">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-565">
         - CompressedFile</span></span><br><span data-ttu-id="ca8cd-566">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ca8cd-566">
         - CustomXmlParts</span></span><br><span data-ttu-id="ca8cd-567">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-567">
         - DocumentEvents</span></span><br><span data-ttu-id="ca8cd-568">
         - File</span><span class="sxs-lookup"><span data-stu-id="ca8cd-568">
         - File</span></span><br><span data-ttu-id="ca8cd-569">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-569">
         - HtmlCoercion</span></span><br><span data-ttu-id="ca8cd-570">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-570">
         - MatrixBindings</span></span><br><span data-ttu-id="ca8cd-571">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-571">
         - MatrixCoercion</span></span><br><span data-ttu-id="ca8cd-572">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-572">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ca8cd-573">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-573">
         - PdfFile</span></span><br><span data-ttu-id="ca8cd-574">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ca8cd-574">
         - Selection</span></span><br><span data-ttu-id="ca8cd-575">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-575">
         - Settings</span></span><br><span data-ttu-id="ca8cd-576">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-576">
         - TableBindings</span></span><br><span data-ttu-id="ca8cd-577">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-577">
         - TableCoercion</span></span><br><span data-ttu-id="ca8cd-578">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-578">
         - TextBindings</span></span><br><span data-ttu-id="ca8cd-579">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-579">
         - TextCoercion</span></span><br><span data-ttu-id="ca8cd-580">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-580">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-581">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="ca8cd-581">Office 2019 on Windows</span></span><br><span data-ttu-id="ca8cd-582">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-582">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ca8cd-583">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-583">- TaskPane</span></span><br><span data-ttu-id="ca8cd-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-585">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-585">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-586">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-586">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ca8cd-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ca8cd-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-589">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-589">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-590">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-590">- BindingEvents</span></span><br><span data-ttu-id="ca8cd-591">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-591">
         - CompressedFile</span></span><br><span data-ttu-id="ca8cd-592">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ca8cd-592">
         - CustomXmlParts</span></span><br><span data-ttu-id="ca8cd-593">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-593">
         - DocumentEvents</span></span><br><span data-ttu-id="ca8cd-594">
         - File</span><span class="sxs-lookup"><span data-stu-id="ca8cd-594">
         - File</span></span><br><span data-ttu-id="ca8cd-595">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-595">
         - HtmlCoercion</span></span><br><span data-ttu-id="ca8cd-596">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-596">
         - MatrixBindings</span></span><br><span data-ttu-id="ca8cd-597">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-597">
         - MatrixCoercion</span></span><br><span data-ttu-id="ca8cd-598">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-598">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ca8cd-599">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-599">
         - PdfFile</span></span><br><span data-ttu-id="ca8cd-600">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ca8cd-600">
         - Selection</span></span><br><span data-ttu-id="ca8cd-601">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-601">
         - Settings</span></span><br><span data-ttu-id="ca8cd-602">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-602">
         - TableBindings</span></span><br><span data-ttu-id="ca8cd-603">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-603">
         - TableCoercion</span></span><br><span data-ttu-id="ca8cd-604">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-604">
         - TextBindings</span></span><br><span data-ttu-id="ca8cd-605">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-605">
         - TextCoercion</span></span><br><span data-ttu-id="ca8cd-606">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-606">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-607">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="ca8cd-607">Office 2016 on Windows</span></span><br><span data-ttu-id="ca8cd-608">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-608">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ca8cd-609">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-609">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ca8cd-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="ca8cd-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="ca8cd-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-613">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-613">- BindingEvents</span></span><br><span data-ttu-id="ca8cd-614">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-614">
         - CompressedFile</span></span><br><span data-ttu-id="ca8cd-615">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ca8cd-615">
         - CustomXmlParts</span></span><br><span data-ttu-id="ca8cd-616">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-616">
         - DocumentEvents</span></span><br><span data-ttu-id="ca8cd-617">
         - File</span><span class="sxs-lookup"><span data-stu-id="ca8cd-617">
         - File</span></span><br><span data-ttu-id="ca8cd-618">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-618">
         - HtmlCoercion</span></span><br><span data-ttu-id="ca8cd-619">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-619">
         - MatrixBindings</span></span><br><span data-ttu-id="ca8cd-620">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-620">
         - MatrixCoercion</span></span><br><span data-ttu-id="ca8cd-621">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-621">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ca8cd-622">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-622">
         - PdfFile</span></span><br><span data-ttu-id="ca8cd-623">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ca8cd-623">
         - Selection</span></span><br><span data-ttu-id="ca8cd-624">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-624">
         - Settings</span></span><br><span data-ttu-id="ca8cd-625">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-625">
         - TableBindings</span></span><br><span data-ttu-id="ca8cd-626">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-626">
         - TableCoercion</span></span><br><span data-ttu-id="ca8cd-627">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-627">
         - TextBindings</span></span><br><span data-ttu-id="ca8cd-628">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-628">
         - TextCoercion</span></span><br><span data-ttu-id="ca8cd-629">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-629">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-630">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="ca8cd-630">Office 2013 on Windows</span></span><br><span data-ttu-id="ca8cd-631">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-631">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ca8cd-632">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-632">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ca8cd-633">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="ca8cd-633">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="ca8cd-634">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-634">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-635">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-635">- BindingEvents</span></span><br><span data-ttu-id="ca8cd-636">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-636">
         - CompressedFile</span></span><br><span data-ttu-id="ca8cd-637">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ca8cd-637">
         - CustomXmlParts</span></span><br><span data-ttu-id="ca8cd-638">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-638">
         - DocumentEvents</span></span><br><span data-ttu-id="ca8cd-639">
         - File</span><span class="sxs-lookup"><span data-stu-id="ca8cd-639">
         - File</span></span><br><span data-ttu-id="ca8cd-640">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-640">
         - HtmlCoercion</span></span><br><span data-ttu-id="ca8cd-641">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-641">
         - MatrixBindings</span></span><br><span data-ttu-id="ca8cd-642">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-642">
         - MatrixCoercion</span></span><br><span data-ttu-id="ca8cd-643">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-643">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ca8cd-644">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-644">
         - PdfFile</span></span><br><span data-ttu-id="ca8cd-645">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ca8cd-645">
         - Selection</span></span><br><span data-ttu-id="ca8cd-646">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-646">
         - Settings</span></span><br><span data-ttu-id="ca8cd-647">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-647">
         - TableBindings</span></span><br><span data-ttu-id="ca8cd-648">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-648">
         - TableCoercion</span></span><br><span data-ttu-id="ca8cd-649">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-649">
         - TextBindings</span></span><br><span data-ttu-id="ca8cd-650">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-650">
         - TextCoercion</span></span><br><span data-ttu-id="ca8cd-651">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-651">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-652">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="ca8cd-652">Office on iPad</span></span><br><span data-ttu-id="ca8cd-653">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-653">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ca8cd-654">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-654">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ca8cd-655">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-655">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-656">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-656">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ca8cd-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ca8cd-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="ca8cd-660">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-660">- BindingEvents</span></span><br><span data-ttu-id="ca8cd-661">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-661">
         - CompressedFile</span></span><br><span data-ttu-id="ca8cd-662">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ca8cd-662">
         - CustomXmlParts</span></span><br><span data-ttu-id="ca8cd-663">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-663">
         - DocumentEvents</span></span><br><span data-ttu-id="ca8cd-664">
         - File</span><span class="sxs-lookup"><span data-stu-id="ca8cd-664">
         - File</span></span><br><span data-ttu-id="ca8cd-665">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-665">
         - HtmlCoercion</span></span><br><span data-ttu-id="ca8cd-666">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-666">
         - MatrixBindings</span></span><br><span data-ttu-id="ca8cd-667">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-667">
         - MatrixCoercion</span></span><br><span data-ttu-id="ca8cd-668">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-668">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ca8cd-669">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-669">
         - PdfFile</span></span><br><span data-ttu-id="ca8cd-670">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ca8cd-670">
         - Selection</span></span><br><span data-ttu-id="ca8cd-671">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-671">
         - Settings</span></span><br><span data-ttu-id="ca8cd-672">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-672">
         - TableBindings</span></span><br><span data-ttu-id="ca8cd-673">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-673">
         - TableCoercion</span></span><br><span data-ttu-id="ca8cd-674">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-674">
         - TextBindings</span></span><br><span data-ttu-id="ca8cd-675">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-675">
         - TextCoercion</span></span><br><span data-ttu-id="ca8cd-676">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-676">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-677">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="ca8cd-677">Office on Mac</span></span><br><span data-ttu-id="ca8cd-678">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-678">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ca8cd-679">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-679">- TaskPane</span></span><br><span data-ttu-id="ca8cd-680">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-680">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-681">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-681">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ca8cd-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ca8cd-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ca8cd-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="ca8cd-687">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-687">- BindingEvents</span></span><br><span data-ttu-id="ca8cd-688">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-688">
         - CompressedFile</span></span><br><span data-ttu-id="ca8cd-689">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ca8cd-689">
         - CustomXmlParts</span></span><br><span data-ttu-id="ca8cd-690">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-690">
         - DocumentEvents</span></span><br><span data-ttu-id="ca8cd-691">
         - File</span><span class="sxs-lookup"><span data-stu-id="ca8cd-691">
         - File</span></span><br><span data-ttu-id="ca8cd-692">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-692">
         - HtmlCoercion</span></span><br><span data-ttu-id="ca8cd-693">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-693">
         - MatrixBindings</span></span><br><span data-ttu-id="ca8cd-694">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-694">
         - MatrixCoercion</span></span><br><span data-ttu-id="ca8cd-695">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-695">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ca8cd-696">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-696">
         - PdfFile</span></span><br><span data-ttu-id="ca8cd-697">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ca8cd-697">
         - Selection</span></span><br><span data-ttu-id="ca8cd-698">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-698">
         - Settings</span></span><br><span data-ttu-id="ca8cd-699">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-699">
         - TableBindings</span></span><br><span data-ttu-id="ca8cd-700">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-700">
         - TableCoercion</span></span><br><span data-ttu-id="ca8cd-701">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-701">
         - TextBindings</span></span><br><span data-ttu-id="ca8cd-702">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-702">
         - TextCoercion</span></span><br><span data-ttu-id="ca8cd-703">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-703">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-704">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="ca8cd-704">Office 2019 on Mac</span></span><br><span data-ttu-id="ca8cd-705">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-705">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ca8cd-706">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-706">- TaskPane</span></span><br><span data-ttu-id="ca8cd-707">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-707">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-708">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-708">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-709">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-709">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="ca8cd-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="ca8cd-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="ca8cd-713">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-713">- BindingEvents</span></span><br><span data-ttu-id="ca8cd-714">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-714">
         - CompressedFile</span></span><br><span data-ttu-id="ca8cd-715">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ca8cd-715">
         - CustomXmlParts</span></span><br><span data-ttu-id="ca8cd-716">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-716">
         - DocumentEvents</span></span><br><span data-ttu-id="ca8cd-717">
         - File</span><span class="sxs-lookup"><span data-stu-id="ca8cd-717">
         - File</span></span><br><span data-ttu-id="ca8cd-718">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-718">
         - HtmlCoercion</span></span><br><span data-ttu-id="ca8cd-719">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-719">
         - MatrixBindings</span></span><br><span data-ttu-id="ca8cd-720">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-720">
         - MatrixCoercion</span></span><br><span data-ttu-id="ca8cd-721">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-721">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ca8cd-722">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-722">
         - PdfFile</span></span><br><span data-ttu-id="ca8cd-723">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ca8cd-723">
         - Selection</span></span><br><span data-ttu-id="ca8cd-724">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-724">
         - Settings</span></span><br><span data-ttu-id="ca8cd-725">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-725">
         - TableBindings</span></span><br><span data-ttu-id="ca8cd-726">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-726">
         - TableCoercion</span></span><br><span data-ttu-id="ca8cd-727">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-727">
         - TextBindings</span></span><br><span data-ttu-id="ca8cd-728">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-728">
         - TextCoercion</span></span><br><span data-ttu-id="ca8cd-729">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-729">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-730">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="ca8cd-730">Office 2016 on Mac</span></span><br><span data-ttu-id="ca8cd-731">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-731">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ca8cd-732">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-732">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ca8cd-733">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-733">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-734">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="ca8cd-734">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="ca8cd-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-736">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-736">- BindingEvents</span></span><br><span data-ttu-id="ca8cd-737">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-737">
         - CompressedFile</span></span><br><span data-ttu-id="ca8cd-738">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="ca8cd-738">
         - CustomXmlParts</span></span><br><span data-ttu-id="ca8cd-739">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-739">
         - DocumentEvents</span></span><br><span data-ttu-id="ca8cd-740">
         - File</span><span class="sxs-lookup"><span data-stu-id="ca8cd-740">
         - File</span></span><br><span data-ttu-id="ca8cd-741">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-741">
         - HtmlCoercion</span></span><br><span data-ttu-id="ca8cd-742">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-742">
         - MatrixBindings</span></span><br><span data-ttu-id="ca8cd-743">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-743">
         - MatrixCoercion</span></span><br><span data-ttu-id="ca8cd-744">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-744">
         - OoxmlCoercion</span></span><br><span data-ttu-id="ca8cd-745">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-745">
         - PdfFile</span></span><br><span data-ttu-id="ca8cd-746">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ca8cd-746">
         - Selection</span></span><br><span data-ttu-id="ca8cd-747">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-747">
         - Settings</span></span><br><span data-ttu-id="ca8cd-748">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-748">
         - TableBindings</span></span><br><span data-ttu-id="ca8cd-749">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-749">
         - TableCoercion</span></span><br><span data-ttu-id="ca8cd-750">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-750">
         - TextBindings</span></span><br><span data-ttu-id="ca8cd-751">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-751">
         - TextCoercion</span></span><br><span data-ttu-id="ca8cd-752">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-752">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="ca8cd-753">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="ca8cd-753">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="ca8cd-754">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="ca8cd-754">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ca8cd-755">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="ca8cd-755">Platform</span></span></th>
    <th><span data-ttu-id="ca8cd-756">拡張点</span><span class="sxs-lookup"><span data-stu-id="ca8cd-756">Extension points</span></span></th>
    <th><span data-ttu-id="ca8cd-757">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="ca8cd-757">API requirement sets</span></span></th>
    <th><span data-ttu-id="ca8cd-758"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-758"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-759">Office on the web</span><span class="sxs-lookup"><span data-stu-id="ca8cd-759">Office on the web</span></span></td>
    <td> <span data-ttu-id="ca8cd-760">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-760">- Content</span></span><br><span data-ttu-id="ca8cd-761">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-761">
         - TaskPane</span></span><br><span data-ttu-id="ca8cd-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-763">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-763">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-764">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-764">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ca8cd-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-767">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ca8cd-767">- ActiveView</span></span><br><span data-ttu-id="ca8cd-768">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-768">
         - CompressedFile</span></span><br><span data-ttu-id="ca8cd-769">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-769">
         - DocumentEvents</span></span><br><span data-ttu-id="ca8cd-770">
         - File</span><span class="sxs-lookup"><span data-stu-id="ca8cd-770">
         - File</span></span><br><span data-ttu-id="ca8cd-771">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-771">
         - PdfFile</span></span><br><span data-ttu-id="ca8cd-772">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ca8cd-772">
         - Selection</span></span><br><span data-ttu-id="ca8cd-773">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-773">
         - Settings</span></span><br><span data-ttu-id="ca8cd-774">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-774">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-775">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="ca8cd-775">Office on Windows</span></span><br><span data-ttu-id="ca8cd-776">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-776">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ca8cd-777">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-777">- Content</span></span><br><span data-ttu-id="ca8cd-778">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-778">
         - TaskPane</span></span><br><span data-ttu-id="ca8cd-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-780">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-780">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ca8cd-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-784">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ca8cd-784">- ActiveView</span></span><br><span data-ttu-id="ca8cd-785">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-785">
         - CompressedFile</span></span><br><span data-ttu-id="ca8cd-786">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-786">
         - DocumentEvents</span></span><br><span data-ttu-id="ca8cd-787">
         - File</span><span class="sxs-lookup"><span data-stu-id="ca8cd-787">
         - File</span></span><br><span data-ttu-id="ca8cd-788">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-788">
         - PdfFile</span></span><br><span data-ttu-id="ca8cd-789">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ca8cd-789">
         - Selection</span></span><br><span data-ttu-id="ca8cd-790">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-790">
         - Settings</span></span><br><span data-ttu-id="ca8cd-791">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-791">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-792">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="ca8cd-792">Office 2019 on Windows</span></span><br><span data-ttu-id="ca8cd-793">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-793">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ca8cd-794">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-794">- Content</span></span><br><span data-ttu-id="ca8cd-795">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-795">
         - TaskPane</span></span><br><span data-ttu-id="ca8cd-796">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-796">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-797">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-797">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-798">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-798">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-799">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ca8cd-799">- ActiveView</span></span><br><span data-ttu-id="ca8cd-800">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-800">
         - CompressedFile</span></span><br><span data-ttu-id="ca8cd-801">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-801">
         - DocumentEvents</span></span><br><span data-ttu-id="ca8cd-802">
         - File</span><span class="sxs-lookup"><span data-stu-id="ca8cd-802">
         - File</span></span><br><span data-ttu-id="ca8cd-803">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-803">
         - PdfFile</span></span><br><span data-ttu-id="ca8cd-804">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ca8cd-804">
         - Selection</span></span><br><span data-ttu-id="ca8cd-805">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-805">
         - Settings</span></span><br><span data-ttu-id="ca8cd-806">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-806">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-807">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="ca8cd-807">Office 2016 on Windows</span></span><br><span data-ttu-id="ca8cd-808">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-808">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ca8cd-809">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-809">- Content</span></span><br><span data-ttu-id="ca8cd-810">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-810">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="ca8cd-811">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="ca8cd-811">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="ca8cd-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-813">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ca8cd-813">- ActiveView</span></span><br><span data-ttu-id="ca8cd-814">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-814">
         - CompressedFile</span></span><br><span data-ttu-id="ca8cd-815">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-815">
         - DocumentEvents</span></span><br><span data-ttu-id="ca8cd-816">
         - File</span><span class="sxs-lookup"><span data-stu-id="ca8cd-816">
         - File</span></span><br><span data-ttu-id="ca8cd-817">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-817">
         - PdfFile</span></span><br><span data-ttu-id="ca8cd-818">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ca8cd-818">
         - Selection</span></span><br><span data-ttu-id="ca8cd-819">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-819">
         - Settings</span></span><br><span data-ttu-id="ca8cd-820">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-820">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-821">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="ca8cd-821">Office 2013 on Windows</span></span><br><span data-ttu-id="ca8cd-822">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-822">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ca8cd-823">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-823">- Content</span></span><br><span data-ttu-id="ca8cd-824">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-824">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="ca8cd-825">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="ca8cd-825">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="ca8cd-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-827">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ca8cd-827">- ActiveView</span></span><br><span data-ttu-id="ca8cd-828">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-828">
         - CompressedFile</span></span><br><span data-ttu-id="ca8cd-829">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-829">
         - DocumentEvents</span></span><br><span data-ttu-id="ca8cd-830">
         - File</span><span class="sxs-lookup"><span data-stu-id="ca8cd-830">
         - File</span></span><br><span data-ttu-id="ca8cd-831">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-831">
         - PdfFile</span></span><br><span data-ttu-id="ca8cd-832">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ca8cd-832">
         - Selection</span></span><br><span data-ttu-id="ca8cd-833">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-833">
         - Settings</span></span><br><span data-ttu-id="ca8cd-834">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-834">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-835">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="ca8cd-835">Office on iPad</span></span><br><span data-ttu-id="ca8cd-836">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-836">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ca8cd-837">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-837">- Content</span></span><br><span data-ttu-id="ca8cd-838">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-838">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="ca8cd-839">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-839">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-842">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ca8cd-842">- ActiveView</span></span><br><span data-ttu-id="ca8cd-843">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-843">
         - CompressedFile</span></span><br><span data-ttu-id="ca8cd-844">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-844">
         - DocumentEvents</span></span><br><span data-ttu-id="ca8cd-845">
         - File</span><span class="sxs-lookup"><span data-stu-id="ca8cd-845">
         - File</span></span><br><span data-ttu-id="ca8cd-846">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-846">
         - PdfFile</span></span><br><span data-ttu-id="ca8cd-847">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ca8cd-847">
         - Selection</span></span><br><span data-ttu-id="ca8cd-848">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-848">
         - Settings</span></span><br><span data-ttu-id="ca8cd-849">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-849">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-850">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="ca8cd-850">Office on Mac</span></span><br><span data-ttu-id="ca8cd-851">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-851">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="ca8cd-852">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-852">- Content</span></span><br><span data-ttu-id="ca8cd-853">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-853">
         - TaskPane</span></span><br><span data-ttu-id="ca8cd-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-855">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-855">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-856">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-856">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="ca8cd-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-859">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ca8cd-859">- ActiveView</span></span><br><span data-ttu-id="ca8cd-860">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-860">
         - CompressedFile</span></span><br><span data-ttu-id="ca8cd-861">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-861">
         - DocumentEvents</span></span><br><span data-ttu-id="ca8cd-862">
         - File</span><span class="sxs-lookup"><span data-stu-id="ca8cd-862">
         - File</span></span><br><span data-ttu-id="ca8cd-863">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-863">
         - PdfFile</span></span><br><span data-ttu-id="ca8cd-864">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ca8cd-864">
         - Selection</span></span><br><span data-ttu-id="ca8cd-865">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-865">
         - Settings</span></span><br><span data-ttu-id="ca8cd-866">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-866">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-867">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="ca8cd-867">Office 2019 on Mac</span></span><br><span data-ttu-id="ca8cd-868">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-868">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ca8cd-869">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-869">- Content</span></span><br><span data-ttu-id="ca8cd-870">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-870">
         - TaskPane</span></span><br><span data-ttu-id="ca8cd-871">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-871">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-872">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-872">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-873">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-873">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-874">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ca8cd-874">- ActiveView</span></span><br><span data-ttu-id="ca8cd-875">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-875">
         - CompressedFile</span></span><br><span data-ttu-id="ca8cd-876">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-876">
         - DocumentEvents</span></span><br><span data-ttu-id="ca8cd-877">
         - File</span><span class="sxs-lookup"><span data-stu-id="ca8cd-877">
         - File</span></span><br><span data-ttu-id="ca8cd-878">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-878">
         - PdfFile</span></span><br><span data-ttu-id="ca8cd-879">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ca8cd-879">
         - Selection</span></span><br><span data-ttu-id="ca8cd-880">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-880">
         - Settings</span></span><br><span data-ttu-id="ca8cd-881">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-881">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-882">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="ca8cd-882">Office 2016 on Mac</span></span><br><span data-ttu-id="ca8cd-883">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-883">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ca8cd-884">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-884">- Content</span></span><br><span data-ttu-id="ca8cd-885">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-885">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="ca8cd-886">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="ca8cd-886">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="ca8cd-887">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-887">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-888">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="ca8cd-888">- ActiveView</span></span><br><span data-ttu-id="ca8cd-889">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-889">
         - CompressedFile</span></span><br><span data-ttu-id="ca8cd-890">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-890">
         - DocumentEvents</span></span><br><span data-ttu-id="ca8cd-891">
         - File</span><span class="sxs-lookup"><span data-stu-id="ca8cd-891">
         - File</span></span><br><span data-ttu-id="ca8cd-892">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="ca8cd-892">
         - PdfFile</span></span><br><span data-ttu-id="ca8cd-893">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="ca8cd-893">
         - Selection</span></span><br><span data-ttu-id="ca8cd-894">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-894">
         - Settings</span></span><br><span data-ttu-id="ca8cd-895">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-895">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="ca8cd-896">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="ca8cd-896">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="ca8cd-897">OneNote</span><span class="sxs-lookup"><span data-stu-id="ca8cd-897">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ca8cd-898">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="ca8cd-898">Platform</span></span></th>
    <th><span data-ttu-id="ca8cd-899">拡張点</span><span class="sxs-lookup"><span data-stu-id="ca8cd-899">Extension points</span></span></th>
    <th><span data-ttu-id="ca8cd-900">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="ca8cd-900">API requirement sets</span></span></th>
    <th><span data-ttu-id="ca8cd-901"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-901"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-902">Office on the web</span><span class="sxs-lookup"><span data-stu-id="ca8cd-902">Office on the web</span></span></td>
    <td> <span data-ttu-id="ca8cd-903">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-903">- Content</span></span><br><span data-ttu-id="ca8cd-904">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-904">
         - TaskPane</span></span><br><span data-ttu-id="ca8cd-905">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-905">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-906">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-906">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-907">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-907">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="ca8cd-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-909">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="ca8cd-909">- DocumentEvents</span></span><br><span data-ttu-id="ca8cd-910">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-910">
         - HtmlCoercion</span></span><br><span data-ttu-id="ca8cd-911">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="ca8cd-911">
         - Settings</span></span><br><span data-ttu-id="ca8cd-912">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-912">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="ca8cd-913">Project</span><span class="sxs-lookup"><span data-stu-id="ca8cd-913">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="ca8cd-914">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="ca8cd-914">Platform</span></span></th>
    <th><span data-ttu-id="ca8cd-915">拡張点</span><span class="sxs-lookup"><span data-stu-id="ca8cd-915">Extension points</span></span></th>
    <th><span data-ttu-id="ca8cd-916">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="ca8cd-916">API requirement sets</span></span></th>
    <th><span data-ttu-id="ca8cd-917"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-917"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-918">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="ca8cd-918">Office 2019 on Windows</span></span><br><span data-ttu-id="ca8cd-919">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-919">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ca8cd-920">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-920">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ca8cd-921">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-921">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-922">- Selection</span><span class="sxs-lookup"><span data-stu-id="ca8cd-922">- Selection</span></span><br><span data-ttu-id="ca8cd-923">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-923">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-924">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="ca8cd-924">Office 2016 on Windows</span></span><br><span data-ttu-id="ca8cd-925">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-925">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ca8cd-926">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-926">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ca8cd-927">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-927">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-928">- Selection</span><span class="sxs-lookup"><span data-stu-id="ca8cd-928">- Selection</span></span><br><span data-ttu-id="ca8cd-929">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-929">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="ca8cd-930">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="ca8cd-930">Office 2013 on Windows</span></span><br><span data-ttu-id="ca8cd-931">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="ca8cd-931">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="ca8cd-932">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="ca8cd-932">- TaskPane</span></span></td>
    <td> <span data-ttu-id="ca8cd-933">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="ca8cd-933">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="ca8cd-934">- Selection</span><span class="sxs-lookup"><span data-stu-id="ca8cd-934">- Selection</span></span><br><span data-ttu-id="ca8cd-935">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="ca8cd-935">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="ca8cd-936">関連項目</span><span class="sxs-lookup"><span data-stu-id="ca8cd-936">See also</span></span>

- [<span data-ttu-id="ca8cd-937">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="ca8cd-937">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="ca8cd-938">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="ca8cd-938">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="ca8cd-939">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="ca8cd-939">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="ca8cd-940">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="ca8cd-940">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="ca8cd-941">API リファレンス ドキュメント</span><span class="sxs-lookup"><span data-stu-id="ca8cd-941">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="ca8cd-942">Office 365 ProPlus の更新履歴</span><span class="sxs-lookup"><span data-stu-id="ca8cd-942">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="ca8cd-943">Office 2016および2019の更新履歴（クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="ca8cd-943">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="ca8cd-944">Office 2013 の更新履歴 （クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="ca8cd-944">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="ca8cd-945">Office 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="ca8cd-945">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="ca8cd-946">Outlook 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="ca8cd-946">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="ca8cd-947">Office for Mac の更新履歴</span><span class="sxs-lookup"><span data-stu-id="ca8cd-947">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="ca8cd-948">Office アドインを構築する</span><span class="sxs-lookup"><span data-stu-id="ca8cd-948">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)