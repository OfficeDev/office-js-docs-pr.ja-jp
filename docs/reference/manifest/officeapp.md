---
title: マニフェスト ファイルの OfficeApp 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 42b6fe2e1c33322b90016d5e7ceec7b1bfe5b72d
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433167"
---
# <a name="officeapp-element"></a><span data-ttu-id="93bbf-102">OfficeApp 要素</span><span class="sxs-lookup"><span data-stu-id="93bbf-102">OfficeApp element</span></span>

<span data-ttu-id="93bbf-103">Office アドインのマニフェストのルート要素。</span><span class="sxs-lookup"><span data-stu-id="93bbf-103">The root element in the manifest of an Office Add-in.</span></span>

<span data-ttu-id="93bbf-104">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="93bbf-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="93bbf-105">構文</span><span class="sxs-lookup"><span data-stu-id="93bbf-105">Syntax</span></span>

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="93bbf-106">次に含まれる</span><span class="sxs-lookup"><span data-stu-id="93bbf-106">Contained in</span></span>

 <span data-ttu-id="93bbf-107">_none_</span><span class="sxs-lookup"><span data-stu-id="93bbf-107">_none_</span></span>

## <a name="must-contain"></a><span data-ttu-id="93bbf-108">含める必要があるもの</span><span class="sxs-lookup"><span data-stu-id="93bbf-108">Must contain</span></span>

|<span data-ttu-id="93bbf-109">**要素**</span><span class="sxs-lookup"><span data-stu-id="93bbf-109">**Element**</span></span>|<span data-ttu-id="93bbf-110">**Content**</span><span class="sxs-lookup"><span data-stu-id="93bbf-110">**Content**</span></span>|<span data-ttu-id="93bbf-111">**Mail**</span><span class="sxs-lookup"><span data-stu-id="93bbf-111">**Mail**</span></span>|<span data-ttu-id="93bbf-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="93bbf-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="93bbf-113">Id</span><span class="sxs-lookup"><span data-stu-id="93bbf-113">Id</span></span>](id.md)|<span data-ttu-id="93bbf-114">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-114">x</span></span>|<span data-ttu-id="93bbf-115">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-115">x</span></span>|<span data-ttu-id="93bbf-116">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-116">x</span></span>|
|[<span data-ttu-id="93bbf-117">Version</span><span class="sxs-lookup"><span data-stu-id="93bbf-117">Version</span></span>](version.md)|<span data-ttu-id="93bbf-118">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-118">x</span></span>|<span data-ttu-id="93bbf-119">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-119">x</span></span>|<span data-ttu-id="93bbf-120">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-120">x</span></span>|
|[<span data-ttu-id="93bbf-121">ProviderName</span><span class="sxs-lookup"><span data-stu-id="93bbf-121">ProviderName</span></span>](providername.md)|<span data-ttu-id="93bbf-122">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-122">x</span></span>|<span data-ttu-id="93bbf-123">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-123">x</span></span>|<span data-ttu-id="93bbf-124">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-124">x</span></span>|
|[<span data-ttu-id="93bbf-125">DefaultLocale</span><span class="sxs-lookup"><span data-stu-id="93bbf-125">DefaultLocale</span></span>](defaultlocale.md)|<span data-ttu-id="93bbf-126">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-126">x</span></span>|<span data-ttu-id="93bbf-127">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-127">x</span></span>|<span data-ttu-id="93bbf-128">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-128">x</span></span>|
|[<span data-ttu-id="93bbf-129">DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="93bbf-129">DefaultSettings</span></span>](defaultsettings.md)|<span data-ttu-id="93bbf-130">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-130">x</span></span>||<span data-ttu-id="93bbf-131">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-131">x</span></span>|
|[<span data-ttu-id="93bbf-132">DisplayName</span><span class="sxs-lookup"><span data-stu-id="93bbf-132">DisplayName</span></span>](displayname.md)|<span data-ttu-id="93bbf-133">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-133">x</span></span>|<span data-ttu-id="93bbf-134">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-134">x</span></span>|<span data-ttu-id="93bbf-135">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-135">x</span></span>|
|[<span data-ttu-id="93bbf-136">Description</span><span class="sxs-lookup"><span data-stu-id="93bbf-136">Description</span></span>](description.md)|<span data-ttu-id="93bbf-137">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-137">x</span></span>|<span data-ttu-id="93bbf-138">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-138">x</span></span>|<span data-ttu-id="93bbf-139">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-139">x</span></span>|
|[<span data-ttu-id="93bbf-140">FormSettings</span><span class="sxs-lookup"><span data-stu-id="93bbf-140">FormSettings</span></span>](formsettings.md)||<span data-ttu-id="93bbf-141">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-141">x</span></span>||
|[<span data-ttu-id="93bbf-142">Permissions</span><span class="sxs-lookup"><span data-stu-id="93bbf-142">Permissions</span></span>](permissions.md)|<span data-ttu-id="93bbf-143">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-143">x</span></span>||<span data-ttu-id="93bbf-144">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-144">x</span></span>|
|[<span data-ttu-id="93bbf-145">Rule</span><span class="sxs-lookup"><span data-stu-id="93bbf-145">Rule</span></span>](rule.md)||<span data-ttu-id="93bbf-146">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-146">x</span></span>||

## <a name="can-contain"></a><span data-ttu-id="93bbf-147">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="93bbf-147">Can contain</span></span>

|<span data-ttu-id="93bbf-148">**要素**</span><span class="sxs-lookup"><span data-stu-id="93bbf-148">**Element**</span></span>|<span data-ttu-id="93bbf-149">**Content**</span><span class="sxs-lookup"><span data-stu-id="93bbf-149">**Content**</span></span>|<span data-ttu-id="93bbf-150">**Mail**</span><span class="sxs-lookup"><span data-stu-id="93bbf-150">**Mail**</span></span>|<span data-ttu-id="93bbf-151">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="93bbf-151">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="93bbf-152">AlternateId</span><span class="sxs-lookup"><span data-stu-id="93bbf-152">AlternateId</span></span>](alternateid.md)|<span data-ttu-id="93bbf-153">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-153">x</span></span>|<span data-ttu-id="93bbf-154">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-154">x</span></span>|<span data-ttu-id="93bbf-155">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-155">x</span></span>|
|[<span data-ttu-id="93bbf-156">IconUrl</span><span class="sxs-lookup"><span data-stu-id="93bbf-156">IconUrl</span></span>](iconurl.md)|<span data-ttu-id="93bbf-157">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-157">x</span></span>|<span data-ttu-id="93bbf-158">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-158">x</span></span>|<span data-ttu-id="93bbf-159">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-159">x</span></span>|
|[<span data-ttu-id="93bbf-160">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="93bbf-160">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|<span data-ttu-id="93bbf-161">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-161">x</span></span>|<span data-ttu-id="93bbf-162">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-162">x</span></span>|<span data-ttu-id="93bbf-163">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-163">x</span></span>|
|[<span data-ttu-id="93bbf-164">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="93bbf-164">SupportUrl</span></span>](supporturl.md)|<span data-ttu-id="93bbf-165">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-165">x</span></span>|<span data-ttu-id="93bbf-166">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-166">x</span></span>|<span data-ttu-id="93bbf-167">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-167">x</span></span>|
|[<span data-ttu-id="93bbf-168">AppDomains</span><span class="sxs-lookup"><span data-stu-id="93bbf-168">AppDomains</span></span>](appdomains.md)|<span data-ttu-id="93bbf-169">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-169">x</span></span>|<span data-ttu-id="93bbf-170">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-170">x</span></span>|<span data-ttu-id="93bbf-171">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-171">x</span></span>|
|[<span data-ttu-id="93bbf-172">Hosts</span><span class="sxs-lookup"><span data-stu-id="93bbf-172">Hosts</span></span>](hosts.md)|<span data-ttu-id="93bbf-173">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-173">x</span></span>|<span data-ttu-id="93bbf-174">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-174">x</span></span>|<span data-ttu-id="93bbf-175">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-175">x</span></span>|
|[<span data-ttu-id="93bbf-176">Requirements</span><span class="sxs-lookup"><span data-stu-id="93bbf-176">Requirements</span></span>](requirements.md)|<span data-ttu-id="93bbf-177">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-177">x</span></span>|<span data-ttu-id="93bbf-178">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-178">x</span></span>|<span data-ttu-id="93bbf-179">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-179">x</span></span>|
|[<span data-ttu-id="93bbf-180">AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="93bbf-180">AllowSnapshot</span></span>](allowsnapshot.md)|<span data-ttu-id="93bbf-181">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-181">x</span></span>|||
|[<span data-ttu-id="93bbf-182">Permissions</span><span class="sxs-lookup"><span data-stu-id="93bbf-182">Permissions</span></span>](permissions.md)||<span data-ttu-id="93bbf-183">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-183">x</span></span>||
|[<span data-ttu-id="93bbf-184">DisableEntityHighlighting</span><span class="sxs-lookup"><span data-stu-id="93bbf-184">DisableEntityHighlighting</span></span>](disableentityhighlighting.md)||<span data-ttu-id="93bbf-185">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-185">x</span></span>||
|[<span data-ttu-id="93bbf-186">Dictionary</span><span class="sxs-lookup"><span data-stu-id="93bbf-186">Dictionary</span></span>](dictionary.md)|||<span data-ttu-id="93bbf-187">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-187">x</span></span>|
|[<span data-ttu-id="93bbf-188">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="93bbf-188">VersionOverrides</span></span>](versionoverrides.md)||<span data-ttu-id="93bbf-189">x</span><span class="sxs-lookup"><span data-stu-id="93bbf-189">x</span></span>||

## <a name="attributes"></a><span data-ttu-id="93bbf-190">属性</span><span class="sxs-lookup"><span data-stu-id="93bbf-190">Attributes</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="93bbf-191">xmlns</span><span class="sxs-lookup"><span data-stu-id="93bbf-191">xmlns</span></span>|<span data-ttu-id="93bbf-p101">Office アドイン マニフェストの名前空間とスキーマ バージョンを定義します。この属性は常に `"http://schemas.microsoft.com/office/appforoffice/1.1"` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="93bbf-p101">Defines the Office Add-in manifest namespace and schema version. This attribute should always be set to  `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span></span>|
|<span data-ttu-id="93bbf-194">xmlns:xsi</span><span class="sxs-lookup"><span data-stu-id="93bbf-194">xmlns:xsi</span></span>|<span data-ttu-id="93bbf-p102">XMLSchema インスタンスを定義します。この属性は常に `"http://www.w3.org/2001/XMLSchema-instance"` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="93bbf-p102">Defines the XMLSchema instance. This attribute should always be set to  `"http://www.w3.org/2001/XMLSchema-instance"`</span></span>|
|<span data-ttu-id="93bbf-197">xsi:type</span><span class="sxs-lookup"><span data-stu-id="93bbf-197">xsi:type</span></span>|<span data-ttu-id="93bbf-p103">Office アドインの種類を定義します。この属性は、`"ContentApp"`、`"MailApp"`、または `"TaskPaneApp"` のいずれかに設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="93bbf-p103">Defines the kind of Office Add-in. This attribute should be set to one of:  `"ContentApp"`,  `"MailApp"`, or  `"TaskPaneApp"`</span></span>|
