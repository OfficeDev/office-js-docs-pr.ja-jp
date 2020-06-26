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
# <a name="office-add-in-host-and-platform-availability"></a><span data-ttu-id="f4b4e-103">Office アドインを使用できるホストおよびプラットフォーム</span><span class="sxs-lookup"><span data-stu-id="f4b4e-103">Office Add-in host and platform availability</span></span>

<span data-ttu-id="f4b4e-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span><span class="sxs-lookup"><span data-stu-id="f4b4e-104">To work as expected, your Office Add-in might depend on a specific Office host, a requirement set, an API member, or a version of the API.</span></span> <span data-ttu-id="f4b4e-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span><span class="sxs-lookup"><span data-stu-id="f4b4e-105">The following tables contain the available platforms, extension points, API requirement sets, and Common APIs that are currently supported for each Office application.</span></span>

> [!NOTE]
> <span data-ttu-id="f4b4e-106">MSI からインストールされる最初の Office 2016 リリースには、ExcelApi 1.1、WordApi 1.1、共通 API の要件セットのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="f4b4e-106">The initial Office 2016 release installed via MSI only contains the ExcelApi 1.1, WordApi 1.1, and Common API requirement sets.</span></span> <span data-ttu-id="f4b4e-107">さまざまなバージョンの Office の更新履歴の詳細については、「[関連項目](#see-also)」セクションをご確認ください。</span><span class="sxs-lookup"><span data-stu-id="f4b4e-107">For more information about the update history of the various Office versions, check out the [See also](#see-also) section.</span></span>

## <a name="excel"></a><span data-ttu-id="f4b4e-108">Excel</span><span class="sxs-lookup"><span data-stu-id="f4b4e-108">Excel</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="f4b4e-109">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="f4b4e-109">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="f4b4e-110">拡張点</span><span class="sxs-lookup"><span data-stu-id="f4b4e-110">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="f4b4e-111">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="f4b4e-111">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="f4b4e-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-112"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-113">Office on the web</span><span class="sxs-lookup"><span data-stu-id="f4b4e-113">Office on the web</span></span></td>
    <td> <span data-ttu-id="f4b4e-114">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-114">- TaskPane</span></span><br><span data-ttu-id="f4b4e-115">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-115">
        - Content</span></span><br><span data-ttu-id="f4b4e-116">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="f4b4e-116">
        - Custom Functions</span></span><br><span data-ttu-id="f4b4e-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="f4b4e-117">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="f4b4e-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-118">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-119">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="f4b4e-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-120">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="f4b4e-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-121">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="f4b4e-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-122">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="f4b4e-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-123">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="f4b4e-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-124">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="f4b4e-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-125">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="f4b4e-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-126">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="f4b4e-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-127">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="f4b4e-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-128">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="f4b4e-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-129">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-online-requirement-set">ExcelApiOnline</a></span></span><br><span data-ttu-id="f4b4e-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-130">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td><span data-ttu-id="f4b4e-131">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-131">
        - BindingEvents</span></span><br><span data-ttu-id="f4b4e-132">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-132">
        - CompressedFile</span></span><br><span data-ttu-id="f4b4e-133">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-133">
        - DocumentEvents</span></span><br><span data-ttu-id="f4b4e-134">
        - File</span><span class="sxs-lookup"><span data-stu-id="f4b4e-134">
        - File</span></span><br><span data-ttu-id="f4b4e-135">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-135">
        - MatrixBindings</span></span><br><span data-ttu-id="f4b4e-136">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-136">
        - MatrixCoercion</span></span><br><span data-ttu-id="f4b4e-137">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="f4b4e-137">
        - Selection</span></span><br><span data-ttu-id="f4b4e-138">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-138">
        - Settings</span></span><br><span data-ttu-id="f4b4e-139">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-139">
        - TableBindings</span></span><br><span data-ttu-id="f4b4e-140">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-140">
        - TableCoercion</span></span><br><span data-ttu-id="f4b4e-141">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-141">
        - TextBindings</span></span><br><span data-ttu-id="f4b4e-142">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-142">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-143">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="f4b4e-143">Office on Windows</span></span><br><span data-ttu-id="f4b4e-144">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-144">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="f4b4e-145">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-145">- TaskPane</span></span><br><span data-ttu-id="f4b4e-146">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-146">
        - Content</span></span><br><span data-ttu-id="f4b4e-147">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="f4b4e-147">
        - Custom Functions</span></span><br><span data-ttu-id="f4b4e-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a>
    </span><span class="sxs-lookup"><span data-stu-id="f4b4e-148">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a>
    </span></span></td>
    <td><span data-ttu-id="f4b4e-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-149">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-150">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="f4b4e-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-151">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="f4b4e-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-152">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="f4b4e-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-153">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="f4b4e-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-154">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="f4b4e-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-155">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="f4b4e-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-156">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="f4b4e-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-157">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="f4b4e-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-158">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="f4b4e-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-159">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="f4b4e-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-160">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-161">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="f4b4e-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-162">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="f4b4e-163">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-163">
        - BindingEvents</span></span><br><span data-ttu-id="f4b4e-164">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-164">
        - CompressedFile</span></span><br><span data-ttu-id="f4b4e-165">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-165">
        - DocumentEvents</span></span><br><span data-ttu-id="f4b4e-166">
        - File</span><span class="sxs-lookup"><span data-stu-id="f4b4e-166">
        - File</span></span><br><span data-ttu-id="f4b4e-167">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-167">
        - MatrixBindings</span></span><br><span data-ttu-id="f4b4e-168">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-168">
        - MatrixCoercion</span></span><br><span data-ttu-id="f4b4e-169">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="f4b4e-169">
        - Selection</span></span><br><span data-ttu-id="f4b4e-170">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-170">
        - Settings</span></span><br><span data-ttu-id="f4b4e-171">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-171">
        - TableBindings</span></span><br><span data-ttu-id="f4b4e-172">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-172">
        - TableCoercion</span></span><br><span data-ttu-id="f4b4e-173">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-173">
        - TextBindings</span></span><br><span data-ttu-id="f4b4e-174">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-174">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-175">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="f4b4e-175">Office 2019 on Windows</span></span><br><span data-ttu-id="f4b4e-176">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-176">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="f4b4e-177">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-177">- TaskPane</span></span><br><span data-ttu-id="f4b4e-178">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-178">
        - Content</span></span><br><span data-ttu-id="f4b4e-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-179">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="f4b4e-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-180">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-181">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="f4b4e-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-182">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="f4b4e-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-183">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="f4b4e-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-184">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="f4b4e-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-185">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="f4b4e-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-186">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="f4b4e-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-187">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="f4b4e-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-188">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-189">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="f4b4e-190">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-190">- BindingEvents</span></span><br><span data-ttu-id="f4b4e-191">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-191">
        - CompressedFile</span></span><br><span data-ttu-id="f4b4e-192">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-192">
        - DocumentEvents</span></span><br><span data-ttu-id="f4b4e-193">
        - File</span><span class="sxs-lookup"><span data-stu-id="f4b4e-193">
        - File</span></span><br><span data-ttu-id="f4b4e-194">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-194">
        - MatrixBindings</span></span><br><span data-ttu-id="f4b4e-195">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-195">
        - MatrixCoercion</span></span><br><span data-ttu-id="f4b4e-196">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="f4b4e-196">
        - Selection</span></span><br><span data-ttu-id="f4b4e-197">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-197">
        - Settings</span></span><br><span data-ttu-id="f4b4e-198">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-198">
        - TableBindings</span></span><br><span data-ttu-id="f4b4e-199">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-199">
        - TableCoercion</span></span><br><span data-ttu-id="f4b4e-200">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-200">
        - TextBindings</span></span><br><span data-ttu-id="f4b4e-201">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-201">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-202">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="f4b4e-202">Office 2016 on Windows</span></span><br><span data-ttu-id="f4b4e-203">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-203">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="f4b4e-204">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-204">- TaskPane</span></span><br><span data-ttu-id="f4b4e-205">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-205">
        - Content</span></span></td>
    <td><span data-ttu-id="f4b4e-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-206">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="f4b4e-207">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="f4b4e-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-208">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="f4b4e-209">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-209">- BindingEvents</span></span><br><span data-ttu-id="f4b4e-210">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-210">
        - CompressedFile</span></span><br><span data-ttu-id="f4b4e-211">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-211">
        - DocumentEvents</span></span><br><span data-ttu-id="f4b4e-212">
        - File</span><span class="sxs-lookup"><span data-stu-id="f4b4e-212">
        - File</span></span><br><span data-ttu-id="f4b4e-213">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-213">
        - MatrixBindings</span></span><br><span data-ttu-id="f4b4e-214">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-214">
        - MatrixCoercion</span></span><br><span data-ttu-id="f4b4e-215">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="f4b4e-215">
        - Selection</span></span><br><span data-ttu-id="f4b4e-216">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-216">
        - Settings</span></span><br><span data-ttu-id="f4b4e-217">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-217">
        - TableBindings</span></span><br><span data-ttu-id="f4b4e-218">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-218">
        - TableCoercion</span></span><br><span data-ttu-id="f4b4e-219">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-219">
        - TextBindings</span></span><br><span data-ttu-id="f4b4e-220">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-220">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-221">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="f4b4e-221">Office 2013 on Windows</span></span><br><span data-ttu-id="f4b4e-222">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-222">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="f4b4e-223">
        - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-223">
        - TaskPane</span></span><br><span data-ttu-id="f4b4e-224">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-224">
        - Content</span></span></td>
    <td>  <span data-ttu-id="f4b4e-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="f4b4e-225">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="f4b4e-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-226">
          - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="f4b4e-227">
        - BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-227">
        - BindingEvents</span></span><br><span data-ttu-id="f4b4e-228">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-228">
        - DocumentEvents</span></span><br><span data-ttu-id="f4b4e-229">
        - File</span><span class="sxs-lookup"><span data-stu-id="f4b4e-229">
        - File</span></span><br><span data-ttu-id="f4b4e-230">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-230">
        - MatrixBindings</span></span><br><span data-ttu-id="f4b4e-231">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-231">
        - MatrixCoercion</span></span><br><span data-ttu-id="f4b4e-232">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="f4b4e-232">
        - Selection</span></span><br><span data-ttu-id="f4b4e-233">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-233">
        - Settings</span></span><br><span data-ttu-id="f4b4e-234">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-234">
        - TableBindings</span></span><br><span data-ttu-id="f4b4e-235">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-235">
        - TableCoercion</span></span><br><span data-ttu-id="f4b4e-236">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-236">
        - TextBindings</span></span><br><span data-ttu-id="f4b4e-237">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-237">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-238">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="f4b4e-238">Office on iPad</span></span><br><span data-ttu-id="f4b4e-239">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-239">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="f4b4e-240">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-240">- TaskPane</span></span><br><span data-ttu-id="f4b4e-241">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-241">
        - Content</span></span></td>
    <td><span data-ttu-id="f4b4e-242">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-242">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-243">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="f4b4e-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-244">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="f4b4e-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-245">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="f4b4e-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-246">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="f4b4e-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-247">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="f4b4e-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-248">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="f4b4e-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-249">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="f4b4e-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-250">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="f4b4e-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-251">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="f4b4e-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-252">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="f4b4e-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-253">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-254">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="f4b4e-255">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-255">- BindingEvents</span></span><br><span data-ttu-id="f4b4e-256">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-256">
        - DocumentEvents</span></span><br><span data-ttu-id="f4b4e-257">
        - File</span><span class="sxs-lookup"><span data-stu-id="f4b4e-257">
        - File</span></span><br><span data-ttu-id="f4b4e-258">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-258">
        - MatrixBindings</span></span><br><span data-ttu-id="f4b4e-259">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-259">
        - MatrixCoercion</span></span><br><span data-ttu-id="f4b4e-260">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="f4b4e-260">
        - Selection</span></span><br><span data-ttu-id="f4b4e-261">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-261">
        - Settings</span></span><br><span data-ttu-id="f4b4e-262">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-262">
        - TableBindings</span></span><br><span data-ttu-id="f4b4e-263">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-263">
        - TableCoercion</span></span><br><span data-ttu-id="f4b4e-264">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-264">
        - TextBindings</span></span><br><span data-ttu-id="f4b4e-265">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-265">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-266">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="f4b4e-266">Office on Mac</span></span><br><span data-ttu-id="f4b4e-267">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-267">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="f4b4e-268">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-268">- TaskPane</span></span><br><span data-ttu-id="f4b4e-269">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-269">
        - Content</span></span><br><span data-ttu-id="f4b4e-270">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="f4b4e-270">
        - Custom Functions</span></span><br><span data-ttu-id="f4b4e-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-271">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="f4b4e-272">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-272">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-273">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="f4b4e-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-274">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="f4b4e-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-275">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="f4b4e-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-276">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="f4b4e-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-277">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="f4b4e-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-278">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="f4b4e-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-279">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="f4b4e-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-280">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-9-requirement-set">ExcelApi 1.9</a></span></span><br><span data-ttu-id="f4b4e-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-281">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-10-requirement-set">ExcelApi 1.10</a></span></span><br><span data-ttu-id="f4b4e-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-282">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-11-requirement-set">ExcelApi 1.11</a></span></span><br><span data-ttu-id="f4b4e-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-283">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-284">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="f4b4e-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-285">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td><span data-ttu-id="f4b4e-286">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-286">- BindingEvents</span></span><br><span data-ttu-id="f4b4e-287">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-287">
        - CompressedFile</span></span><br><span data-ttu-id="f4b4e-288">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-288">
        - DocumentEvents</span></span><br><span data-ttu-id="f4b4e-289">
        - File</span><span class="sxs-lookup"><span data-stu-id="f4b4e-289">
        - File</span></span><br><span data-ttu-id="f4b4e-290">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-290">
        - MatrixBindings</span></span><br><span data-ttu-id="f4b4e-291">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-291">
        - MatrixCoercion</span></span><br><span data-ttu-id="f4b4e-292">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-292">
        - PdfFile</span></span><br><span data-ttu-id="f4b4e-293">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="f4b4e-293">
        - Selection</span></span><br><span data-ttu-id="f4b4e-294">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-294">
        - Settings</span></span><br><span data-ttu-id="f4b4e-295">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-295">
        - TableBindings</span></span><br><span data-ttu-id="f4b4e-296">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-296">
        - TableCoercion</span></span><br><span data-ttu-id="f4b4e-297">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-297">
        - TextBindings</span></span><br><span data-ttu-id="f4b4e-298">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-298">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-299">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="f4b4e-299">Office 2019 on Mac</span></span><br><span data-ttu-id="f4b4e-300">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-300">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="f4b4e-301">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-301">- TaskPane</span></span><br><span data-ttu-id="f4b4e-302">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-302">
        - Content</span></span><br><span data-ttu-id="f4b4e-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-303">
        - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td><span data-ttu-id="f4b4e-304">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-304">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-305">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-2-requirement-set">ExcelApi 1.2</a></span></span><br><span data-ttu-id="f4b4e-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-306">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-3-requirement-set">ExcelApi 1.3</a></span></span><br><span data-ttu-id="f4b4e-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-307">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-4-requirement-set">ExcelApi 1.4</a></span></span><br><span data-ttu-id="f4b4e-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-308">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-5-requirement-set">ExcelApi 1.5</a></span></span><br><span data-ttu-id="f4b4e-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-309">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-6-requirement-set">ExcelApi 1.6</a></span></span><br><span data-ttu-id="f4b4e-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-310">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-7-requirement-set">ExcelApi 1.7</a></span></span><br><span data-ttu-id="f4b4e-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-311">
        - <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-8-requirement-set">ExcelApi 1.8</a></span></span><br><span data-ttu-id="f4b4e-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-312">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-313">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="f4b4e-314">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-314">- BindingEvents</span></span><br><span data-ttu-id="f4b4e-315">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-315">
        - CompressedFile</span></span><br><span data-ttu-id="f4b4e-316">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-316">
        - DocumentEvents</span></span><br><span data-ttu-id="f4b4e-317">
        - File</span><span class="sxs-lookup"><span data-stu-id="f4b4e-317">
        - File</span></span><br><span data-ttu-id="f4b4e-318">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-318">
        - MatrixBindings</span></span><br><span data-ttu-id="f4b4e-319">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-319">
        - MatrixCoercion</span></span><br><span data-ttu-id="f4b4e-320">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-320">
        - PdfFile</span></span><br><span data-ttu-id="f4b4e-321">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="f4b4e-321">
        - Selection</span></span><br><span data-ttu-id="f4b4e-322">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-322">
        - Settings</span></span><br><span data-ttu-id="f4b4e-323">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-323">
        - TableBindings</span></span><br><span data-ttu-id="f4b4e-324">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-324">
        - TableCoercion</span></span><br><span data-ttu-id="f4b4e-325">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-325">
        - TextBindings</span></span><br><span data-ttu-id="f4b4e-326">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-326">
        - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-327">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="f4b4e-327">Office 2016 on Mac</span></span><br><span data-ttu-id="f4b4e-328">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-328">(one-time purchase)</span></span></td>
    <td><span data-ttu-id="f4b4e-329">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-329">- TaskPane</span></span><br><span data-ttu-id="f4b4e-330">
        - コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-330">
        - Content</span></span></td>
    <td><span data-ttu-id="f4b4e-331">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-331">- <a href="/office/dev/add-ins/reference/requirement-sets/excel-api-1-1-requirement-set">ExcelApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-332">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="f4b4e-332">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="f4b4e-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-333">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td><span data-ttu-id="f4b4e-334">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-334">- BindingEvents</span></span><br><span data-ttu-id="f4b4e-335">
        - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-335">
        - CompressedFile</span></span><br><span data-ttu-id="f4b4e-336">
        - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-336">
        - DocumentEvents</span></span><br><span data-ttu-id="f4b4e-337">
        - File</span><span class="sxs-lookup"><span data-stu-id="f4b4e-337">
        - File</span></span><br><span data-ttu-id="f4b4e-338">
        - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-338">
        - MatrixBindings</span></span><br><span data-ttu-id="f4b4e-339">
        - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-339">
        - MatrixCoercion</span></span><br><span data-ttu-id="f4b4e-340">
        - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-340">
        - PdfFile</span></span><br><span data-ttu-id="f4b4e-341">
        - Selection</span><span class="sxs-lookup"><span data-stu-id="f4b4e-341">
        - Selection</span></span><br><span data-ttu-id="f4b4e-342">
        - Settings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-342">
        - Settings</span></span><br><span data-ttu-id="f4b4e-343">
        - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-343">
        - TableBindings</span></span><br><span data-ttu-id="f4b4e-344">
        - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-344">
        - TableCoercion</span></span><br><span data-ttu-id="f4b4e-345">
        - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-345">
        - TextBindings</span></span><br><span data-ttu-id="f4b4e-346">
        - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-346">
        - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="f4b4e-347">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="f4b4e-347">*&ast; - Added with post-release updates.*</span></span>

## <a name="custom-functions-excel-only"></a><span data-ttu-id="f4b4e-348">カスタム関数 (Excel のみ)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-348">Custom Functions (Excel only)</span></span>

<table style="width:80%">
  <tr>
    <th style="width:10%"><span data-ttu-id="f4b4e-349">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="f4b4e-349">Platform</span></span></th>
    <th style="width:10%"><span data-ttu-id="f4b4e-350">拡張点</span><span class="sxs-lookup"><span data-stu-id="f4b4e-350">Extension points</span></span></th>
    <th style="width:20%"><span data-ttu-id="f4b4e-351">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="f4b4e-351">API requirement sets</span></span></th>
    <th style="width:40%"><span data-ttu-id="f4b4e-352"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-352"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-353">Office on the web</span><span class="sxs-lookup"><span data-stu-id="f4b4e-353">Office on the web</span></span></td>
    <td><span data-ttu-id="f4b4e-354">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="f4b4e-354">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="f4b4e-355">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-355">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-356">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="f4b4e-356">Office on Windows</span></span><br><span data-ttu-id="f4b4e-357">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-357">(connected to Office 365 subscription)</span></span></td>
    <td><span data-ttu-id="f4b4e-358">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="f4b4e-358">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="f4b4e-359">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-359">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-360">Office on Mac</span><span class="sxs-lookup"><span data-stu-id="f4b4e-360">Office on Mac</span></span><br><span data-ttu-id="f4b4e-361">(Office 365 に接続された)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-361">(connected to Office 365)</span></span></td>
    <td><span data-ttu-id="f4b4e-362">
        - カスタム関数</span><span class="sxs-lookup"><span data-stu-id="f4b4e-362">
        - Custom Functions</span></span></td>
    <td><span data-ttu-id="f4b4e-363">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-363">
        - <a href="/office/dev/add-ins/excel/custom-functions-requirement-sets">CustomFunctionsRuntime 1.1</a></span></span></td>
    <td>
    </td>
  </tr>
</table>

## <a name="outlook"></a><span data-ttu-id="f4b4e-364">Outlook</span><span class="sxs-lookup"><span data-stu-id="f4b4e-364">Outlook</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="f4b4e-365">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="f4b4e-365">Platform</span></span></th>
    <th><span data-ttu-id="f4b4e-366">拡張点</span><span class="sxs-lookup"><span data-stu-id="f4b4e-366">Extension points</span></span></th>
    <th><span data-ttu-id="f4b4e-367">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="f4b4e-367">API requirement sets</span></span></th>
    <th><span data-ttu-id="f4b4e-368"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-368"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-369">Office on the web</span><span class="sxs-lookup"><span data-stu-id="f4b4e-369">Office on the web</span></span><br><span data-ttu-id="f4b4e-370">(モダン)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-370">(modern)</span></span></td>
    <td> <span data-ttu-id="f4b4e-371">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-371">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="f4b4e-372">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-372">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="f4b4e-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-373">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="f4b4e-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-374">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="f4b4e-375">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-375">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-376">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-376">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f4b4e-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-377">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f4b4e-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-378">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f4b4e-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-379">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f4b4e-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-380">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="f4b4e-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-381">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="f4b4e-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-382">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="f4b4e-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-383">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="f4b4e-384">利用不可</span><span class="sxs-lookup"><span data-stu-id="f4b4e-384">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-385">Office on the web</span><span class="sxs-lookup"><span data-stu-id="f4b4e-385">Office on the web</span></span><br><span data-ttu-id="f4b4e-386">(クラシック)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-386">(classic)</span></span></td>
    <td> <span data-ttu-id="f4b4e-387">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-387">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="f4b4e-388">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-388">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="f4b4e-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-389">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="f4b4e-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-390">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="f4b4e-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-391">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-392">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f4b4e-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-393">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f4b4e-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-394">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f4b4e-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-395">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f4b4e-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-396">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="f4b4e-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-397">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="f4b4e-398">使用不可</span><span class="sxs-lookup"><span data-stu-id="f4b4e-398">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-399">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="f4b4e-399">Office on Windows</span></span><br><span data-ttu-id="f4b4e-400">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-400">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="f4b4e-401">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-401">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="f4b4e-402">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-402">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="f4b4e-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-403">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="f4b4e-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-404">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="f4b4e-405">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-405">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="f4b4e-406">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">モジュール</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-406">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-407">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-407">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f4b4e-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-408">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f4b4e-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-409">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f4b4e-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-410">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f4b4e-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-411">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="f4b4e-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-412">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="f4b4e-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-413">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="f4b4e-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-414">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="f4b4e-415">利用不可</span><span class="sxs-lookup"><span data-stu-id="f4b4e-415">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-416">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="f4b4e-416">Office 2019 on Windows</span></span><br><span data-ttu-id="f4b4e-417">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-417">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f4b4e-418">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-418">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="f4b4e-419">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-419">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="f4b4e-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-420">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="f4b4e-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-421">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="f4b4e-422">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-422">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="f4b4e-423">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">モジュール</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-423">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-424">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-424">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f4b4e-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-425">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f4b4e-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-426">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f4b4e-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-427">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f4b4e-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-428">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="f4b4e-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-429">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="f4b4e-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-430">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span></td>
    <td><span data-ttu-id="f4b4e-431">使用不可</span><span class="sxs-lookup"><span data-stu-id="f4b4e-431">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-432">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="f4b4e-432">Office 2016 on Windows</span></span><br><span data-ttu-id="f4b4e-433">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-433">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f4b4e-434">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-434">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="f4b4e-435">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-435">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="f4b4e-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-436">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="f4b4e-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-437">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="f4b4e-438">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-438">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span><br><span data-ttu-id="f4b4e-439">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">モジュール</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-439">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#module">Modules</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-440">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-440">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f4b4e-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-441">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f4b4e-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-442">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f4b4e-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="f4b4e-443">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="f4b4e-444">使用不可</span><span class="sxs-lookup"><span data-stu-id="f4b4e-444">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-445">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="f4b4e-445">Office 2013 on Windows</span></span><br><span data-ttu-id="f4b4e-446">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-446">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f4b4e-447">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-447">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="f4b4e-448">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-448">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="f4b4e-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-449">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="f4b4e-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-450">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br>
    <td> <span data-ttu-id="f4b4e-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-451">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f4b4e-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-452">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f4b4e-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span><span class="sxs-lookup"><span data-stu-id="f4b4e-453">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a>*</span></span><br><span data-ttu-id="f4b4e-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span><span class="sxs-lookup"><span data-stu-id="f4b4e-454">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a>*</span></span></td>
    <td><span data-ttu-id="f4b4e-455">使用不可</span><span class="sxs-lookup"><span data-stu-id="f4b4e-455">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-456">iOS 上の Office</span><span class="sxs-lookup"><span data-stu-id="f4b4e-456">Office on iOS</span></span><br><span data-ttu-id="f4b4e-457">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-457">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="f4b4e-458">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-458">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="f4b4e-459">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-459">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-460">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-460">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f4b4e-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-461">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f4b4e-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-462">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f4b4e-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-463">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f4b4e-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-464">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="f4b4e-465">使用不可</span><span class="sxs-lookup"><span data-stu-id="f4b4e-465">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-466">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="f4b4e-466">Office on Mac</span></span><br><span data-ttu-id="f4b4e-467">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-467">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="f4b4e-468">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-468">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="f4b4e-469">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-469">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="f4b4e-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-470">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="f4b4e-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-471">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="f4b4e-472">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-472">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-473">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-473">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f4b4e-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-474">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f4b4e-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-475">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f4b4e-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-476">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f4b4e-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-477">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="f4b4e-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-478">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span><br><span data-ttu-id="f4b4e-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-479">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7">Mailbox 1.7</a></span></span><br><span data-ttu-id="f4b4e-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-480">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8">Mailbox 1.8</a></span></span></td>
    <td><span data-ttu-id="f4b4e-481">利用不可</span><span class="sxs-lookup"><span data-stu-id="f4b4e-481">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-482">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="f4b4e-482">Office 2019 on Mac</span></span><br><span data-ttu-id="f4b4e-483">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-483">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f4b4e-484">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-484">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="f4b4e-485">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-485">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="f4b4e-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-486">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="f4b4e-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-487">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="f4b4e-488">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-488">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-489">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-489">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f4b4e-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-490">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f4b4e-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-491">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f4b4e-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-492">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f4b4e-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-493">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="f4b4e-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-494">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="f4b4e-495">使用不可</span><span class="sxs-lookup"><span data-stu-id="f4b4e-495">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-496">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="f4b4e-496">Office 2016 on Mac</span></span><br><span data-ttu-id="f4b4e-497">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-497">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f4b4e-498">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-498">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="f4b4e-499">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">メッセージ作成</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-499">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#messagecomposecommandsurface">Message Compose</a></span></span><br><span data-ttu-id="f4b4e-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">予定の出席者 (読み取り)</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-500">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentattendeecommandsurface">Appointment Attendee (Read)</a></span></span><br><span data-ttu-id="f4b4e-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">予定の開催者 (作成)</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-501">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#appointmentorganizercommandsurface">Appointment Organizer (Compose)</a></span></span><br><span data-ttu-id="f4b4e-502">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-502">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-503">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-503">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f4b4e-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-504">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f4b4e-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-505">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f4b4e-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-506">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f4b4e-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-507">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span><br><span data-ttu-id="f4b4e-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-508">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6">Mailbox 1.6</a></span></span></td>
    <td><span data-ttu-id="f4b4e-509">使用不可</span><span class="sxs-lookup"><span data-stu-id="f4b4e-509">Not available</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-510">Android 上の Office</span><span class="sxs-lookup"><span data-stu-id="f4b4e-510">Office on Android</span></span><br><span data-ttu-id="f4b4e-511">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-511">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="f4b4e-512">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">既読メッセージ</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-512">- <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobilemessagereadcommandsurface">Message Read</a></span></span><br><span data-ttu-id="f4b4e-513">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">予定の開催者 (作成): オンライン会議</a> (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-513">
      - <a href="/office/dev/add-ins/reference/manifest/extensionpoint#mobileonlinemeetingcommandsurface-preview">Appointment Organizer (Compose): online meeting</a> (preview)</span></span><br><span data-ttu-id="f4b4e-514">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-514">
      - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-515">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-515">- <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1">Mailbox 1.1</a></span></span><br><span data-ttu-id="f4b4e-516">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-516">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2">Mailbox 1.2</a></span></span><br><span data-ttu-id="f4b4e-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-517">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3">Mailbox 1.3</a></span></span><br><span data-ttu-id="f4b4e-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-518">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4">Mailbox 1.4</a></span></span><br><span data-ttu-id="f4b4e-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-519">
      - <a href="/office/dev/add-ins/reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5">Mailbox 1.5</a></span></span></td>
    <td><span data-ttu-id="f4b4e-520">利用不可</span><span class="sxs-lookup"><span data-stu-id="f4b4e-520">Not available</span></span></td>
  </tr>
</table>

<span data-ttu-id="f4b4e-521">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="f4b4e-521">*&ast; - Added with post-release updates.*</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f4b4e-522">要件セットのクライアント サポートは、Exchange サーバー サポートの制限を受ける場合があります。</span><span class="sxs-lookup"><span data-stu-id="f4b4e-522">Client support for a requirement set may be restricted by Exchange server support.</span></span> <span data-ttu-id="f4b4e-523">Exchange サーバーおよび Outlook クライアントによってサポートされている要件セットの範囲の詳細については、「[Outlook JavaScript APIの要件セット](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f4b4e-523">See [Outlook JavaScript API requirement sets](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) for details about the range of requirement sets supported by Exchange server and Outlook clients.</span></span>

<br/>

## <a name="word"></a><span data-ttu-id="f4b4e-524">Word</span><span class="sxs-lookup"><span data-stu-id="f4b4e-524">Word</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="f4b4e-525">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="f4b4e-525">Platform</span></span></th>
    <th><span data-ttu-id="f4b4e-526">拡張点</span><span class="sxs-lookup"><span data-stu-id="f4b4e-526">Extension points</span></span></th>
    <th><span data-ttu-id="f4b4e-527">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="f4b4e-527">API requirement sets</span></span></th>
    <th><span data-ttu-id="f4b4e-528"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-528"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-529">Office on the web</span><span class="sxs-lookup"><span data-stu-id="f4b4e-529">Office on the web</span></span></td>
    <td> <span data-ttu-id="f4b4e-530">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-530">- TaskPane</span></span><br><span data-ttu-id="f4b4e-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-531">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-532">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-532">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-533">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="f4b4e-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-534">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="f4b4e-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-535">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-536">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="f4b4e-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-537">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-538">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-538">- BindingEvents</span></span><br><span data-ttu-id="f4b4e-539">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f4b4e-539">
         - CustomXmlParts</span></span><br><span data-ttu-id="f4b4e-540">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-540">
         - DocumentEvents</span></span><br><span data-ttu-id="f4b4e-541">
         - File</span><span class="sxs-lookup"><span data-stu-id="f4b4e-541">
         - File</span></span><br><span data-ttu-id="f4b4e-542">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-542">
         - HtmlCoercion</span></span><br><span data-ttu-id="f4b4e-543">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-543">
         - MatrixBindings</span></span><br><span data-ttu-id="f4b4e-544">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-544">
         - MatrixCoercion</span></span><br><span data-ttu-id="f4b4e-545">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-545">
         - OoxmlCoercion</span></span><br><span data-ttu-id="f4b4e-546">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-546">
         - PdfFile</span></span><br><span data-ttu-id="f4b4e-547">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f4b4e-547">
         - Selection</span></span><br><span data-ttu-id="f4b4e-548">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-548">
         - Settings</span></span><br><span data-ttu-id="f4b4e-549">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-549">
         - TableBindings</span></span><br><span data-ttu-id="f4b4e-550">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-550">
         - TableCoercion</span></span><br><span data-ttu-id="f4b4e-551">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-551">
         - TextBindings</span></span><br><span data-ttu-id="f4b4e-552">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-552">
         - TextCoercion</span></span><br><span data-ttu-id="f4b4e-553">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-553">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-554">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="f4b4e-554">Office on Windows</span></span><br><span data-ttu-id="f4b4e-555">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-555">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="f4b4e-556">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-556">- TaskPane</span></span><br><span data-ttu-id="f4b4e-557">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-557">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-558">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-558">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-559">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="f4b4e-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-560">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="f4b4e-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-561">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-562">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-562">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="f4b4e-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-563">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-564">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-564">- BindingEvents</span></span><br><span data-ttu-id="f4b4e-565">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-565">
         - CompressedFile</span></span><br><span data-ttu-id="f4b4e-566">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f4b4e-566">
         - CustomXmlParts</span></span><br><span data-ttu-id="f4b4e-567">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-567">
         - DocumentEvents</span></span><br><span data-ttu-id="f4b4e-568">
         - File</span><span class="sxs-lookup"><span data-stu-id="f4b4e-568">
         - File</span></span><br><span data-ttu-id="f4b4e-569">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-569">
         - HtmlCoercion</span></span><br><span data-ttu-id="f4b4e-570">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-570">
         - MatrixBindings</span></span><br><span data-ttu-id="f4b4e-571">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-571">
         - MatrixCoercion</span></span><br><span data-ttu-id="f4b4e-572">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-572">
         - OoxmlCoercion</span></span><br><span data-ttu-id="f4b4e-573">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-573">
         - PdfFile</span></span><br><span data-ttu-id="f4b4e-574">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f4b4e-574">
         - Selection</span></span><br><span data-ttu-id="f4b4e-575">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-575">
         - Settings</span></span><br><span data-ttu-id="f4b4e-576">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-576">
         - TableBindings</span></span><br><span data-ttu-id="f4b4e-577">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-577">
         - TableCoercion</span></span><br><span data-ttu-id="f4b4e-578">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-578">
         - TextBindings</span></span><br><span data-ttu-id="f4b4e-579">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-579">
         - TextCoercion</span></span><br><span data-ttu-id="f4b4e-580">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-580">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-581">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="f4b4e-581">Office 2019 on Windows</span></span><br><span data-ttu-id="f4b4e-582">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-582">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f4b4e-583">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-583">- TaskPane</span></span><br><span data-ttu-id="f4b4e-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-584">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-585">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-585">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-586">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-586">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="f4b4e-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-587">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="f4b4e-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-588">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-589">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-589">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-590">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-590">- BindingEvents</span></span><br><span data-ttu-id="f4b4e-591">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-591">
         - CompressedFile</span></span><br><span data-ttu-id="f4b4e-592">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f4b4e-592">
         - CustomXmlParts</span></span><br><span data-ttu-id="f4b4e-593">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-593">
         - DocumentEvents</span></span><br><span data-ttu-id="f4b4e-594">
         - File</span><span class="sxs-lookup"><span data-stu-id="f4b4e-594">
         - File</span></span><br><span data-ttu-id="f4b4e-595">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-595">
         - HtmlCoercion</span></span><br><span data-ttu-id="f4b4e-596">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-596">
         - MatrixBindings</span></span><br><span data-ttu-id="f4b4e-597">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-597">
         - MatrixCoercion</span></span><br><span data-ttu-id="f4b4e-598">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-598">
         - OoxmlCoercion</span></span><br><span data-ttu-id="f4b4e-599">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-599">
         - PdfFile</span></span><br><span data-ttu-id="f4b4e-600">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f4b4e-600">
         - Selection</span></span><br><span data-ttu-id="f4b4e-601">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-601">
         - Settings</span></span><br><span data-ttu-id="f4b4e-602">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-602">
         - TableBindings</span></span><br><span data-ttu-id="f4b4e-603">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-603">
         - TableCoercion</span></span><br><span data-ttu-id="f4b4e-604">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-604">
         - TextBindings</span></span><br><span data-ttu-id="f4b4e-605">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-605">
         - TextCoercion</span></span><br><span data-ttu-id="f4b4e-606">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-606">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-607">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="f4b4e-607">Office 2016 on Windows</span></span><br><span data-ttu-id="f4b4e-608">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-608">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f4b4e-609">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-609">- TaskPane</span></span></td>
    <td> <span data-ttu-id="f4b4e-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-610">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="f4b4e-611">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="f4b4e-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-612">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-613">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-613">- BindingEvents</span></span><br><span data-ttu-id="f4b4e-614">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-614">
         - CompressedFile</span></span><br><span data-ttu-id="f4b4e-615">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f4b4e-615">
         - CustomXmlParts</span></span><br><span data-ttu-id="f4b4e-616">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-616">
         - DocumentEvents</span></span><br><span data-ttu-id="f4b4e-617">
         - File</span><span class="sxs-lookup"><span data-stu-id="f4b4e-617">
         - File</span></span><br><span data-ttu-id="f4b4e-618">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-618">
         - HtmlCoercion</span></span><br><span data-ttu-id="f4b4e-619">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-619">
         - MatrixBindings</span></span><br><span data-ttu-id="f4b4e-620">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-620">
         - MatrixCoercion</span></span><br><span data-ttu-id="f4b4e-621">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-621">
         - OoxmlCoercion</span></span><br><span data-ttu-id="f4b4e-622">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-622">
         - PdfFile</span></span><br><span data-ttu-id="f4b4e-623">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f4b4e-623">
         - Selection</span></span><br><span data-ttu-id="f4b4e-624">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-624">
         - Settings</span></span><br><span data-ttu-id="f4b4e-625">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-625">
         - TableBindings</span></span><br><span data-ttu-id="f4b4e-626">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-626">
         - TableCoercion</span></span><br><span data-ttu-id="f4b4e-627">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-627">
         - TextBindings</span></span><br><span data-ttu-id="f4b4e-628">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-628">
         - TextCoercion</span></span><br><span data-ttu-id="f4b4e-629">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-629">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-630">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="f4b4e-630">Office 2013 on Windows</span></span><br><span data-ttu-id="f4b4e-631">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-631">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f4b4e-632">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-632">- TaskPane</span></span></td>
    <td> <span data-ttu-id="f4b4e-633">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="f4b4e-633">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="f4b4e-634">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-634">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-635">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-635">- BindingEvents</span></span><br><span data-ttu-id="f4b4e-636">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-636">
         - CompressedFile</span></span><br><span data-ttu-id="f4b4e-637">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f4b4e-637">
         - CustomXmlParts</span></span><br><span data-ttu-id="f4b4e-638">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-638">
         - DocumentEvents</span></span><br><span data-ttu-id="f4b4e-639">
         - File</span><span class="sxs-lookup"><span data-stu-id="f4b4e-639">
         - File</span></span><br><span data-ttu-id="f4b4e-640">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-640">
         - HtmlCoercion</span></span><br><span data-ttu-id="f4b4e-641">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-641">
         - MatrixBindings</span></span><br><span data-ttu-id="f4b4e-642">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-642">
         - MatrixCoercion</span></span><br><span data-ttu-id="f4b4e-643">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-643">
         - OoxmlCoercion</span></span><br><span data-ttu-id="f4b4e-644">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-644">
         - PdfFile</span></span><br><span data-ttu-id="f4b4e-645">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f4b4e-645">
         - Selection</span></span><br><span data-ttu-id="f4b4e-646">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-646">
         - Settings</span></span><br><span data-ttu-id="f4b4e-647">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-647">
         - TableBindings</span></span><br><span data-ttu-id="f4b4e-648">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-648">
         - TableCoercion</span></span><br><span data-ttu-id="f4b4e-649">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-649">
         - TextBindings</span></span><br><span data-ttu-id="f4b4e-650">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-650">
         - TextCoercion</span></span><br><span data-ttu-id="f4b4e-651">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-651">
         - TextFile</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-652">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="f4b4e-652">Office on iPad</span></span><br><span data-ttu-id="f4b4e-653">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-653">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="f4b4e-654">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-654">- TaskPane</span></span></td>
    <td> <span data-ttu-id="f4b4e-655">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-655">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-656">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-656">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="f4b4e-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-657">
         - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="f4b4e-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-658">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-659">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="f4b4e-660">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-660">- BindingEvents</span></span><br><span data-ttu-id="f4b4e-661">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-661">
         - CompressedFile</span></span><br><span data-ttu-id="f4b4e-662">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f4b4e-662">
         - CustomXmlParts</span></span><br><span data-ttu-id="f4b4e-663">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-663">
         - DocumentEvents</span></span><br><span data-ttu-id="f4b4e-664">
         - File</span><span class="sxs-lookup"><span data-stu-id="f4b4e-664">
         - File</span></span><br><span data-ttu-id="f4b4e-665">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-665">
         - HtmlCoercion</span></span><br><span data-ttu-id="f4b4e-666">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-666">
         - MatrixBindings</span></span><br><span data-ttu-id="f4b4e-667">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-667">
         - MatrixCoercion</span></span><br><span data-ttu-id="f4b4e-668">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-668">
         - OoxmlCoercion</span></span><br><span data-ttu-id="f4b4e-669">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-669">
         - PdfFile</span></span><br><span data-ttu-id="f4b4e-670">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f4b4e-670">
         - Selection</span></span><br><span data-ttu-id="f4b4e-671">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-671">
         - Settings</span></span><br><span data-ttu-id="f4b4e-672">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-672">
         - TableBindings</span></span><br><span data-ttu-id="f4b4e-673">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-673">
         - TableCoercion</span></span><br><span data-ttu-id="f4b4e-674">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-674">
         - TextBindings</span></span><br><span data-ttu-id="f4b4e-675">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-675">
         - TextCoercion</span></span><br><span data-ttu-id="f4b4e-676">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-676">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-677">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="f4b4e-677">Office on Mac</span></span><br><span data-ttu-id="f4b4e-678">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-678">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="f4b4e-679">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-679">- TaskPane</span></span><br><span data-ttu-id="f4b4e-680">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-680">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-681">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-681">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-682">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="f4b4e-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-683">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="f4b4e-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-684">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-685">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="f4b4e-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-686">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
</td>
    <td> <span data-ttu-id="f4b4e-687">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-687">- BindingEvents</span></span><br><span data-ttu-id="f4b4e-688">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-688">
         - CompressedFile</span></span><br><span data-ttu-id="f4b4e-689">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f4b4e-689">
         - CustomXmlParts</span></span><br><span data-ttu-id="f4b4e-690">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-690">
         - DocumentEvents</span></span><br><span data-ttu-id="f4b4e-691">
         - File</span><span class="sxs-lookup"><span data-stu-id="f4b4e-691">
         - File</span></span><br><span data-ttu-id="f4b4e-692">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-692">
         - HtmlCoercion</span></span><br><span data-ttu-id="f4b4e-693">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-693">
         - MatrixBindings</span></span><br><span data-ttu-id="f4b4e-694">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-694">
         - MatrixCoercion</span></span><br><span data-ttu-id="f4b4e-695">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-695">
         - OoxmlCoercion</span></span><br><span data-ttu-id="f4b4e-696">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-696">
         - PdfFile</span></span><br><span data-ttu-id="f4b4e-697">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f4b4e-697">
         - Selection</span></span><br><span data-ttu-id="f4b4e-698">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-698">
         - Settings</span></span><br><span data-ttu-id="f4b4e-699">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-699">
         - TableBindings</span></span><br><span data-ttu-id="f4b4e-700">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-700">
         - TableCoercion</span></span><br><span data-ttu-id="f4b4e-701">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-701">
         - TextBindings</span></span><br><span data-ttu-id="f4b4e-702">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-702">
         - TextCoercion</span></span><br><span data-ttu-id="f4b4e-703">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-703">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-704">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="f4b4e-704">Office 2019 on Mac</span></span><br><span data-ttu-id="f4b4e-705">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-705">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f4b4e-706">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-706">- TaskPane</span></span><br><span data-ttu-id="f4b4e-707">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-707">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-708">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-708">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-709">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-709">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-2-requirement-set">WordApi 1.2</a></span></span><br><span data-ttu-id="f4b4e-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-710">
        - <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-3-requirement-set">WordApi 1.3</a></span></span><br><span data-ttu-id="f4b4e-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-711">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-712">
        - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
</td>
    <td> <span data-ttu-id="f4b4e-713">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-713">- BindingEvents</span></span><br><span data-ttu-id="f4b4e-714">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-714">
         - CompressedFile</span></span><br><span data-ttu-id="f4b4e-715">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f4b4e-715">
         - CustomXmlParts</span></span><br><span data-ttu-id="f4b4e-716">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-716">
         - DocumentEvents</span></span><br><span data-ttu-id="f4b4e-717">
         - File</span><span class="sxs-lookup"><span data-stu-id="f4b4e-717">
         - File</span></span><br><span data-ttu-id="f4b4e-718">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-718">
         - HtmlCoercion</span></span><br><span data-ttu-id="f4b4e-719">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-719">
         - MatrixBindings</span></span><br><span data-ttu-id="f4b4e-720">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-720">
         - MatrixCoercion</span></span><br><span data-ttu-id="f4b4e-721">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-721">
         - OoxmlCoercion</span></span><br><span data-ttu-id="f4b4e-722">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-722">
         - PdfFile</span></span><br><span data-ttu-id="f4b4e-723">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f4b4e-723">
         - Selection</span></span><br><span data-ttu-id="f4b4e-724">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-724">
         - Settings</span></span><br><span data-ttu-id="f4b4e-725">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-725">
         - TableBindings</span></span><br><span data-ttu-id="f4b4e-726">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-726">
         - TableCoercion</span></span><br><span data-ttu-id="f4b4e-727">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-727">
         - TextBindings</span></span><br><span data-ttu-id="f4b4e-728">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-728">
         - TextCoercion</span></span><br><span data-ttu-id="f4b4e-729">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-729">
         - TextFile</span></span> </td>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-730">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="f4b4e-730">Office 2016 on Mac</span></span><br><span data-ttu-id="f4b4e-731">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-731">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f4b4e-732">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-732">- TaskPane</span></span></td>
    <td> <span data-ttu-id="f4b4e-733">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-733">- <a href="/office/dev/add-ins/reference/requirement-sets/word-api-1-1-requirement-set">WordApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-734">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span><span class="sxs-lookup"><span data-stu-id="f4b4e-734">
        - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>*</span></span><br><span data-ttu-id="f4b4e-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-735">
       - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-736">- BindingEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-736">- BindingEvents</span></span><br><span data-ttu-id="f4b4e-737">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-737">
         - CompressedFile</span></span><br><span data-ttu-id="f4b4e-738">
         - CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="f4b4e-738">
         - CustomXmlParts</span></span><br><span data-ttu-id="f4b4e-739">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-739">
         - DocumentEvents</span></span><br><span data-ttu-id="f4b4e-740">
         - File</span><span class="sxs-lookup"><span data-stu-id="f4b4e-740">
         - File</span></span><br><span data-ttu-id="f4b4e-741">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-741">
         - HtmlCoercion</span></span><br><span data-ttu-id="f4b4e-742">
         - MatrixBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-742">
         - MatrixBindings</span></span><br><span data-ttu-id="f4b4e-743">
         - MatrixCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-743">
         - MatrixCoercion</span></span><br><span data-ttu-id="f4b4e-744">
         - OoxmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-744">
         - OoxmlCoercion</span></span><br><span data-ttu-id="f4b4e-745">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-745">
         - PdfFile</span></span><br><span data-ttu-id="f4b4e-746">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f4b4e-746">
         - Selection</span></span><br><span data-ttu-id="f4b4e-747">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-747">
         - Settings</span></span><br><span data-ttu-id="f4b4e-748">
         - TableBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-748">
         - TableBindings</span></span><br><span data-ttu-id="f4b4e-749">
         - TableCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-749">
         - TableCoercion</span></span><br><span data-ttu-id="f4b4e-750">
         - TextBindings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-750">
         - TextBindings</span></span><br><span data-ttu-id="f4b4e-751">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-751">
         - TextCoercion</span></span><br><span data-ttu-id="f4b4e-752">
         - TextFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-752">
         - TextFile</span></span> </td>
  </tr>
</table>

<span data-ttu-id="f4b4e-753">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="f4b4e-753">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="powerpoint"></a><span data-ttu-id="f4b4e-754">PowerPoint</span><span class="sxs-lookup"><span data-stu-id="f4b4e-754">PowerPoint</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="f4b4e-755">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="f4b4e-755">Platform</span></span></th>
    <th><span data-ttu-id="f4b4e-756">拡張点</span><span class="sxs-lookup"><span data-stu-id="f4b4e-756">Extension points</span></span></th>
    <th><span data-ttu-id="f4b4e-757">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="f4b4e-757">API requirement sets</span></span></th>
    <th><span data-ttu-id="f4b4e-758"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-758"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-759">Office on the web</span><span class="sxs-lookup"><span data-stu-id="f4b4e-759">Office on the web</span></span></td>
    <td> <span data-ttu-id="f4b4e-760">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-760">- Content</span></span><br><span data-ttu-id="f4b4e-761">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-761">
         - TaskPane</span></span><br><span data-ttu-id="f4b4e-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-762">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-763">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-763">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-764">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-764">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-765">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="f4b4e-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-766">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-767">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f4b4e-767">- ActiveView</span></span><br><span data-ttu-id="f4b4e-768">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-768">
         - CompressedFile</span></span><br><span data-ttu-id="f4b4e-769">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-769">
         - DocumentEvents</span></span><br><span data-ttu-id="f4b4e-770">
         - File</span><span class="sxs-lookup"><span data-stu-id="f4b4e-770">
         - File</span></span><br><span data-ttu-id="f4b4e-771">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-771">
         - PdfFile</span></span><br><span data-ttu-id="f4b4e-772">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f4b4e-772">
         - Selection</span></span><br><span data-ttu-id="f4b4e-773">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-773">
         - Settings</span></span><br><span data-ttu-id="f4b4e-774">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-774">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-775">Windows での Office</span><span class="sxs-lookup"><span data-stu-id="f4b4e-775">Office on Windows</span></span><br><span data-ttu-id="f4b4e-776">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-776">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="f4b4e-777">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-777">- Content</span></span><br><span data-ttu-id="f4b4e-778">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-778">
         - TaskPane</span></span><br><span data-ttu-id="f4b4e-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-779">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-780">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-780">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-781">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-782">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="f4b4e-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-783">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-784">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f4b4e-784">- ActiveView</span></span><br><span data-ttu-id="f4b4e-785">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-785">
         - CompressedFile</span></span><br><span data-ttu-id="f4b4e-786">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-786">
         - DocumentEvents</span></span><br><span data-ttu-id="f4b4e-787">
         - File</span><span class="sxs-lookup"><span data-stu-id="f4b4e-787">
         - File</span></span><br><span data-ttu-id="f4b4e-788">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-788">
         - PdfFile</span></span><br><span data-ttu-id="f4b4e-789">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f4b4e-789">
         - Selection</span></span><br><span data-ttu-id="f4b4e-790">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-790">
         - Settings</span></span><br><span data-ttu-id="f4b4e-791">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-791">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-792">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="f4b4e-792">Office 2019 on Windows</span></span><br><span data-ttu-id="f4b4e-793">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-793">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f4b4e-794">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-794">- Content</span></span><br><span data-ttu-id="f4b4e-795">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-795">
         - TaskPane</span></span><br><span data-ttu-id="f4b4e-796">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-796">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-797">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-797">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-798">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-798">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-799">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f4b4e-799">- ActiveView</span></span><br><span data-ttu-id="f4b4e-800">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-800">
         - CompressedFile</span></span><br><span data-ttu-id="f4b4e-801">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-801">
         - DocumentEvents</span></span><br><span data-ttu-id="f4b4e-802">
         - File</span><span class="sxs-lookup"><span data-stu-id="f4b4e-802">
         - File</span></span><br><span data-ttu-id="f4b4e-803">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-803">
         - PdfFile</span></span><br><span data-ttu-id="f4b4e-804">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f4b4e-804">
         - Selection</span></span><br><span data-ttu-id="f4b4e-805">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-805">
         - Settings</span></span><br><span data-ttu-id="f4b4e-806">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-806">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-807">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="f4b4e-807">Office 2016 on Windows</span></span><br><span data-ttu-id="f4b4e-808">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-808">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f4b4e-809">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-809">- Content</span></span><br><span data-ttu-id="f4b4e-810">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-810">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="f4b4e-811">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="f4b4e-811">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="f4b4e-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-812">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-813">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f4b4e-813">- ActiveView</span></span><br><span data-ttu-id="f4b4e-814">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-814">
         - CompressedFile</span></span><br><span data-ttu-id="f4b4e-815">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-815">
         - DocumentEvents</span></span><br><span data-ttu-id="f4b4e-816">
         - File</span><span class="sxs-lookup"><span data-stu-id="f4b4e-816">
         - File</span></span><br><span data-ttu-id="f4b4e-817">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-817">
         - PdfFile</span></span><br><span data-ttu-id="f4b4e-818">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f4b4e-818">
         - Selection</span></span><br><span data-ttu-id="f4b4e-819">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-819">
         - Settings</span></span><br><span data-ttu-id="f4b4e-820">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-820">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-821">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="f4b4e-821">Office 2013 on Windows</span></span><br><span data-ttu-id="f4b4e-822">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-822">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f4b4e-823">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-823">- Content</span></span><br><span data-ttu-id="f4b4e-824">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-824">
         - TaskPane</span></span><br>
    </td>
    <td> <span data-ttu-id="f4b4e-825">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="f4b4e-825">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="f4b4e-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-826">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-827">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f4b4e-827">- ActiveView</span></span><br><span data-ttu-id="f4b4e-828">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-828">
         - CompressedFile</span></span><br><span data-ttu-id="f4b4e-829">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-829">
         - DocumentEvents</span></span><br><span data-ttu-id="f4b4e-830">
         - File</span><span class="sxs-lookup"><span data-stu-id="f4b4e-830">
         - File</span></span><br><span data-ttu-id="f4b4e-831">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-831">
         - PdfFile</span></span><br><span data-ttu-id="f4b4e-832">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f4b4e-832">
         - Selection</span></span><br><span data-ttu-id="f4b4e-833">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-833">
         - Settings</span></span><br><span data-ttu-id="f4b4e-834">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-834">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-835">iPad 上の Office</span><span class="sxs-lookup"><span data-stu-id="f4b4e-835">Office on iPad</span></span><br><span data-ttu-id="f4b4e-836">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-836">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="f4b4e-837">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-837">- Content</span></span><br><span data-ttu-id="f4b4e-838">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-838">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="f4b4e-839">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-839">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-840">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-841">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-842">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f4b4e-842">- ActiveView</span></span><br><span data-ttu-id="f4b4e-843">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-843">
         - CompressedFile</span></span><br><span data-ttu-id="f4b4e-844">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-844">
         - DocumentEvents</span></span><br><span data-ttu-id="f4b4e-845">
         - File</span><span class="sxs-lookup"><span data-stu-id="f4b4e-845">
         - File</span></span><br><span data-ttu-id="f4b4e-846">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-846">
         - PdfFile</span></span><br><span data-ttu-id="f4b4e-847">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f4b4e-847">
         - Selection</span></span><br><span data-ttu-id="f4b4e-848">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-848">
         - Settings</span></span><br><span data-ttu-id="f4b4e-849">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-849">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-850">Mac 上の Office</span><span class="sxs-lookup"><span data-stu-id="f4b4e-850">Office on Mac</span></span><br><span data-ttu-id="f4b4e-851">(Office 365 サブスクリプションに接続済み)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-851">(connected to Office 365 subscription)</span></span></td>
    <td> <span data-ttu-id="f4b4e-852">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-852">- Content</span></span><br><span data-ttu-id="f4b4e-853">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-853">
         - TaskPane</span></span><br><span data-ttu-id="f4b4e-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-854">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-855">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-855">- <a href="/office/dev/add-ins/reference/requirement-sets/powerpoint-api-requirement-sets">PowerPointApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-856">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-856">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-857">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span><br><span data-ttu-id="f4b4e-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-858">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-12">ImageCoercion 1.2</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-859">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f4b4e-859">- ActiveView</span></span><br><span data-ttu-id="f4b4e-860">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-860">
         - CompressedFile</span></span><br><span data-ttu-id="f4b4e-861">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-861">
         - DocumentEvents</span></span><br><span data-ttu-id="f4b4e-862">
         - File</span><span class="sxs-lookup"><span data-stu-id="f4b4e-862">
         - File</span></span><br><span data-ttu-id="f4b4e-863">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-863">
         - PdfFile</span></span><br><span data-ttu-id="f4b4e-864">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f4b4e-864">
         - Selection</span></span><br><span data-ttu-id="f4b4e-865">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-865">
         - Settings</span></span><br><span data-ttu-id="f4b4e-866">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-866">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-867">Mac 上の Office 2019</span><span class="sxs-lookup"><span data-stu-id="f4b4e-867">Office 2019 on Mac</span></span><br><span data-ttu-id="f4b4e-868">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-868">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f4b4e-869">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-869">- Content</span></span><br><span data-ttu-id="f4b4e-870">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-870">
         - TaskPane</span></span><br><span data-ttu-id="f4b4e-871">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-871">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-872">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-872">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-873">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-873">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-874">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f4b4e-874">- ActiveView</span></span><br><span data-ttu-id="f4b4e-875">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-875">
         - CompressedFile</span></span><br><span data-ttu-id="f4b4e-876">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-876">
         - DocumentEvents</span></span><br><span data-ttu-id="f4b4e-877">
         - File</span><span class="sxs-lookup"><span data-stu-id="f4b4e-877">
         - File</span></span><br><span data-ttu-id="f4b4e-878">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-878">
         - PdfFile</span></span><br><span data-ttu-id="f4b4e-879">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f4b4e-879">
         - Selection</span></span><br><span data-ttu-id="f4b4e-880">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-880">
         - Settings</span></span><br><span data-ttu-id="f4b4e-881">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-881">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-882">Mac 上の Office 2016</span><span class="sxs-lookup"><span data-stu-id="f4b4e-882">Office 2016 on Mac</span></span><br><span data-ttu-id="f4b4e-883">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-883">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f4b4e-884">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-884">- Content</span></span><br><span data-ttu-id="f4b4e-885">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-885">
         - TaskPane</span></span></td>
    <td> <span data-ttu-id="f4b4e-886">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span><span class="sxs-lookup"><span data-stu-id="f4b4e-886">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>\*</span></span><br><span data-ttu-id="f4b4e-887">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-887">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-888">- ActiveView</span><span class="sxs-lookup"><span data-stu-id="f4b4e-888">- ActiveView</span></span><br><span data-ttu-id="f4b4e-889">
         - CompressedFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-889">
         - CompressedFile</span></span><br><span data-ttu-id="f4b4e-890">
         - DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-890">
         - DocumentEvents</span></span><br><span data-ttu-id="f4b4e-891">
         - File</span><span class="sxs-lookup"><span data-stu-id="f4b4e-891">
         - File</span></span><br><span data-ttu-id="f4b4e-892">
         - PdfFile</span><span class="sxs-lookup"><span data-stu-id="f4b4e-892">
         - PdfFile</span></span><br><span data-ttu-id="f4b4e-893">
         - Selection</span><span class="sxs-lookup"><span data-stu-id="f4b4e-893">
         - Selection</span></span><br><span data-ttu-id="f4b4e-894">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-894">
         - Settings</span></span><br><span data-ttu-id="f4b4e-895">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-895">
         - TextCoercion</span></span></td>
  </tr>
</table>

<span data-ttu-id="f4b4e-896">*&ast; - リリース後の更新プログラムで追加されました。*</span><span class="sxs-lookup"><span data-stu-id="f4b4e-896">*&ast; - Added with post-release updates.*</span></span>

<br/>

## <a name="onenote"></a><span data-ttu-id="f4b4e-897">OneNote</span><span class="sxs-lookup"><span data-stu-id="f4b4e-897">OneNote</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="f4b4e-898">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="f4b4e-898">Platform</span></span></th>
    <th><span data-ttu-id="f4b4e-899">拡張点</span><span class="sxs-lookup"><span data-stu-id="f4b4e-899">Extension points</span></span></th>
    <th><span data-ttu-id="f4b4e-900">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="f4b4e-900">API requirement sets</span></span></th>
    <th><span data-ttu-id="f4b4e-901"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-901"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-902">Office on the web</span><span class="sxs-lookup"><span data-stu-id="f4b4e-902">Office on the web</span></span></td>
    <td> <span data-ttu-id="f4b4e-903">- コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-903">- Content</span></span><br><span data-ttu-id="f4b4e-904">
         - 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-904">
         - TaskPane</span></span><br><span data-ttu-id="f4b4e-905">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">アドイン コマンド</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-905">
         - <a href="/office/dev/add-ins/reference/requirement-sets/add-in-commands-requirement-sets">Add-in Commands</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-906">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-906">- <a href="/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-907">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-907">
         - <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span><br><span data-ttu-id="f4b4e-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-908">
         - <a href="/office/dev/add-ins/reference/requirement-sets/image-coercion-requirement-sets#imagecoercion-11">ImageCoercion 1.1</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-909">- DocumentEvents</span><span class="sxs-lookup"><span data-stu-id="f4b4e-909">- DocumentEvents</span></span><br><span data-ttu-id="f4b4e-910">
         - HtmlCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-910">
         - HtmlCoercion</span></span><br><span data-ttu-id="f4b4e-911">
         - Settings</span><span class="sxs-lookup"><span data-stu-id="f4b4e-911">
         - Settings</span></span><br><span data-ttu-id="f4b4e-912">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-912">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="project"></a><span data-ttu-id="f4b4e-913">Project</span><span class="sxs-lookup"><span data-stu-id="f4b4e-913">Project</span></span>

<table style="width:80%">
  <tr>
    <th><span data-ttu-id="f4b4e-914">プラットフォーム</span><span class="sxs-lookup"><span data-stu-id="f4b4e-914">Platform</span></span></th>
    <th><span data-ttu-id="f4b4e-915">拡張点</span><span class="sxs-lookup"><span data-stu-id="f4b4e-915">Extension points</span></span></th>
    <th><span data-ttu-id="f4b4e-916">API 要件セット</span><span class="sxs-lookup"><span data-stu-id="f4b4e-916">API requirement sets</span></span></th>
    <th><span data-ttu-id="f4b4e-917"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>共通 API</b></a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-917"><a href="/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets"><b>Common APIs</b></a></span></span></th>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-918">Windows 版 Office 2019</span><span class="sxs-lookup"><span data-stu-id="f4b4e-918">Office 2019 on Windows</span></span><br><span data-ttu-id="f4b4e-919">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-919">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f4b4e-920">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-920">- TaskPane</span></span></td>
    <td> <span data-ttu-id="f4b4e-921">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-921">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-922">- Selection</span><span class="sxs-lookup"><span data-stu-id="f4b4e-922">- Selection</span></span><br><span data-ttu-id="f4b4e-923">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-923">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-924">Windows 版 Office 2016</span><span class="sxs-lookup"><span data-stu-id="f4b4e-924">Office 2016 on Windows</span></span><br><span data-ttu-id="f4b4e-925">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-925">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f4b4e-926">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-926">- TaskPane</span></span></td>
    <td> <span data-ttu-id="f4b4e-927">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-927">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-928">- Selection</span><span class="sxs-lookup"><span data-stu-id="f4b4e-928">- Selection</span></span><br><span data-ttu-id="f4b4e-929">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-929">
         - TextCoercion</span></span></td>
  </tr>
  <tr>
    <td><span data-ttu-id="f4b4e-930">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="f4b4e-930">Office 2013 on Windows</span></span><br><span data-ttu-id="f4b4e-931">(1 回限りの購入)</span><span class="sxs-lookup"><span data-stu-id="f4b4e-931">(one-time purchase)</span></span></td>
    <td> <span data-ttu-id="f4b4e-932">- 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="f4b4e-932">- TaskPane</span></span></td>
    <td> <span data-ttu-id="f4b4e-933">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span><span class="sxs-lookup"><span data-stu-id="f4b4e-933">- <a href="/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></span></span></td>
    <td> <span data-ttu-id="f4b4e-934">- Selection</span><span class="sxs-lookup"><span data-stu-id="f4b4e-934">- Selection</span></span><br><span data-ttu-id="f4b4e-935">
         - TextCoercion</span><span class="sxs-lookup"><span data-stu-id="f4b4e-935">
         - TextCoercion</span></span></td>
  </tr>
</table>

<br/>

## <a name="see-also"></a><span data-ttu-id="f4b4e-936">関連項目</span><span class="sxs-lookup"><span data-stu-id="f4b4e-936">See also</span></span>

- [<span data-ttu-id="f4b4e-937">Office アドイン プラットフォームの概要</span><span class="sxs-lookup"><span data-stu-id="f4b4e-937">Office Add-ins platform overview</span></span>](office-add-ins.md)
- [<span data-ttu-id="f4b4e-938">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="f4b4e-938">Office versions and requirement sets</span></span>](../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="f4b4e-939">共通 API の要件セット</span><span class="sxs-lookup"><span data-stu-id="f4b4e-939">Common API requirement sets</span></span>](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [<span data-ttu-id="f4b4e-940">アドイン コマンドの要件セット</span><span class="sxs-lookup"><span data-stu-id="f4b4e-940">Add-in Commands requirement sets</span></span>](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [<span data-ttu-id="f4b4e-941">API リファレンス ドキュメント</span><span class="sxs-lookup"><span data-stu-id="f4b4e-941">API Reference documentation</span></span>](../reference/javascript-api-for-office.md)
- [<span data-ttu-id="f4b4e-942">Office 365 ProPlus の更新履歴</span><span class="sxs-lookup"><span data-stu-id="f4b4e-942">Update history for Office 365 ProPlus</span></span>](/officeupdates/update-history-office365-proplus-by-date)
- [<span data-ttu-id="f4b4e-943">Office 2016および2019の更新履歴（クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="f4b4e-943">Office 2016 and 2019 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2019)
- [<span data-ttu-id="f4b4e-944">Office 2013 の更新履歴 （クリックして実行）</span><span class="sxs-lookup"><span data-stu-id="f4b4e-944">Office 2013 update history (Click-To-Run)</span></span>](/officeupdates/update-history-office-2013)
- [<span data-ttu-id="f4b4e-945">Office 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="f4b4e-945">Office 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/office-updates-msi)
- [<span data-ttu-id="f4b4e-946">Outlook 2010、2013、および2016の更新履歴（MSI）</span><span class="sxs-lookup"><span data-stu-id="f4b4e-946">Outlook 2010, 2013, and 2016 update history (MSI)</span></span>](/officeupdates/outlook-updates-msi)
- [<span data-ttu-id="f4b4e-947">Office for Mac の更新履歴</span><span class="sxs-lookup"><span data-stu-id="f4b4e-947">Update history for Office for Mac</span></span>](/officeupdates/update-history-office-for-mac)
- [<span data-ttu-id="f4b4e-948">Office アドインを構築する</span><span class="sxs-lookup"><span data-stu-id="f4b4e-948">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)