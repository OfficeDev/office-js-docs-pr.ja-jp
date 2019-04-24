---
title: マニフェスト ファイルの OfficeApp 要素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 86f38ab77e98bb01370e40c8ada38bae171e0c2d
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450458"
---
# <a name="officeapp-element"></a><span data-ttu-id="76721-102">OfficeApp 要素</span><span class="sxs-lookup"><span data-stu-id="76721-102">OfficeApp element</span></span>

<span data-ttu-id="76721-103">Office アドインのマニフェストのルート要素。</span><span class="sxs-lookup"><span data-stu-id="76721-103">The root element in the manifest of an Office Add-in.</span></span>

<span data-ttu-id="76721-104">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="76721-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="76721-105">構文</span><span class="sxs-lookup"><span data-stu-id="76721-105">Syntax</span></span>

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="76721-106">次に含まれる</span><span class="sxs-lookup"><span data-stu-id="76721-106">Contained in</span></span>

 <span data-ttu-id="76721-107">_none_</span><span class="sxs-lookup"><span data-stu-id="76721-107">_none_</span></span>

## <a name="must-contain"></a><span data-ttu-id="76721-108">含める必要があるもの</span><span class="sxs-lookup"><span data-stu-id="76721-108">Must contain</span></span>

|<span data-ttu-id="76721-109">**要素**</span><span class="sxs-lookup"><span data-stu-id="76721-109">**Element**</span></span>|<span data-ttu-id="76721-110">**コンテンツ**</span><span class="sxs-lookup"><span data-stu-id="76721-110">**Content**</span></span>|<span data-ttu-id="76721-111">**メール**</span><span class="sxs-lookup"><span data-stu-id="76721-111">**Mail**</span></span>|<span data-ttu-id="76721-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="76721-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="76721-113">Id</span><span class="sxs-lookup"><span data-stu-id="76721-113">Id</span></span>](id.md)|<span data-ttu-id="76721-114">x</span><span class="sxs-lookup"><span data-stu-id="76721-114">x</span></span>|<span data-ttu-id="76721-115">x</span><span class="sxs-lookup"><span data-stu-id="76721-115">x</span></span>|<span data-ttu-id="76721-116">x</span><span class="sxs-lookup"><span data-stu-id="76721-116">x</span></span>|
|[<span data-ttu-id="76721-117">バージョン</span><span class="sxs-lookup"><span data-stu-id="76721-117">Version</span></span>](version.md)|<span data-ttu-id="76721-118">x</span><span class="sxs-lookup"><span data-stu-id="76721-118">x</span></span>|<span data-ttu-id="76721-119">x</span><span class="sxs-lookup"><span data-stu-id="76721-119">x</span></span>|<span data-ttu-id="76721-120">x</span><span class="sxs-lookup"><span data-stu-id="76721-120">x</span></span>|
|[<span data-ttu-id="76721-121">ProviderName</span><span class="sxs-lookup"><span data-stu-id="76721-121">ProviderName</span></span>](providername.md)|<span data-ttu-id="76721-122">x</span><span class="sxs-lookup"><span data-stu-id="76721-122">x</span></span>|<span data-ttu-id="76721-123">x</span><span class="sxs-lookup"><span data-stu-id="76721-123">x</span></span>|<span data-ttu-id="76721-124">x</span><span class="sxs-lookup"><span data-stu-id="76721-124">x</span></span>|
|[<span data-ttu-id="76721-125">DefaultLocale</span><span class="sxs-lookup"><span data-stu-id="76721-125">DefaultLocale</span></span>](defaultlocale.md)|<span data-ttu-id="76721-126">x</span><span class="sxs-lookup"><span data-stu-id="76721-126">x</span></span>|<span data-ttu-id="76721-127">x</span><span class="sxs-lookup"><span data-stu-id="76721-127">x</span></span>|<span data-ttu-id="76721-128">x</span><span class="sxs-lookup"><span data-stu-id="76721-128">x</span></span>|
|[<span data-ttu-id="76721-129">DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="76721-129">DefaultSettings</span></span>](defaultsettings.md)|<span data-ttu-id="76721-130">x</span><span class="sxs-lookup"><span data-stu-id="76721-130">x</span></span>||<span data-ttu-id="76721-131">x</span><span class="sxs-lookup"><span data-stu-id="76721-131">x</span></span>|
|[<span data-ttu-id="76721-132">DisplayName</span><span class="sxs-lookup"><span data-stu-id="76721-132">DisplayName</span></span>](displayname.md)|<span data-ttu-id="76721-133">x</span><span class="sxs-lookup"><span data-stu-id="76721-133">x</span></span>|<span data-ttu-id="76721-134">x</span><span class="sxs-lookup"><span data-stu-id="76721-134">x</span></span>|<span data-ttu-id="76721-135">x</span><span class="sxs-lookup"><span data-stu-id="76721-135">x</span></span>|
|[<span data-ttu-id="76721-136">説明</span><span class="sxs-lookup"><span data-stu-id="76721-136">Description</span></span>](description.md)|<span data-ttu-id="76721-137">x</span><span class="sxs-lookup"><span data-stu-id="76721-137">x</span></span>|<span data-ttu-id="76721-138">x</span><span class="sxs-lookup"><span data-stu-id="76721-138">x</span></span>|<span data-ttu-id="76721-139">x</span><span class="sxs-lookup"><span data-stu-id="76721-139">x</span></span>|
|[<span data-ttu-id="76721-140">FormSettings</span><span class="sxs-lookup"><span data-stu-id="76721-140">FormSettings</span></span>](formsettings.md)||<span data-ttu-id="76721-141">x</span><span class="sxs-lookup"><span data-stu-id="76721-141">x</span></span>||
|[<span data-ttu-id="76721-142">アクセス許可</span><span class="sxs-lookup"><span data-stu-id="76721-142">Permissions</span></span>](permissions.md)|<span data-ttu-id="76721-143">x</span><span class="sxs-lookup"><span data-stu-id="76721-143">x</span></span>||<span data-ttu-id="76721-144">x</span><span class="sxs-lookup"><span data-stu-id="76721-144">x</span></span>|
|[<span data-ttu-id="76721-145">Rule</span><span class="sxs-lookup"><span data-stu-id="76721-145">Rule</span></span>](rule.md)||<span data-ttu-id="76721-146">x</span><span class="sxs-lookup"><span data-stu-id="76721-146">x</span></span>||

## <a name="can-contain"></a><span data-ttu-id="76721-147">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="76721-147">Can contain</span></span>

|<span data-ttu-id="76721-148">**Element**</span><span class="sxs-lookup"><span data-stu-id="76721-148">**Element**</span></span>|<span data-ttu-id="76721-149">**コンテンツ**</span><span class="sxs-lookup"><span data-stu-id="76721-149">**Content**</span></span>|<span data-ttu-id="76721-150">**メール**</span><span class="sxs-lookup"><span data-stu-id="76721-150">**Mail**</span></span>|<span data-ttu-id="76721-151">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="76721-151">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="76721-152">AlternateId</span><span class="sxs-lookup"><span data-stu-id="76721-152">AlternateId</span></span>](alternateid.md)|<span data-ttu-id="76721-153">x</span><span class="sxs-lookup"><span data-stu-id="76721-153">x</span></span>|<span data-ttu-id="76721-154">x</span><span class="sxs-lookup"><span data-stu-id="76721-154">x</span></span>|<span data-ttu-id="76721-155">x</span><span class="sxs-lookup"><span data-stu-id="76721-155">x</span></span>|
|[<span data-ttu-id="76721-156">IconUrl</span><span class="sxs-lookup"><span data-stu-id="76721-156">IconUrl</span></span>](iconurl.md)|<span data-ttu-id="76721-157">x</span><span class="sxs-lookup"><span data-stu-id="76721-157">x</span></span>|<span data-ttu-id="76721-158">x</span><span class="sxs-lookup"><span data-stu-id="76721-158">x</span></span>|<span data-ttu-id="76721-159">x</span><span class="sxs-lookup"><span data-stu-id="76721-159">x</span></span>|
|[<span data-ttu-id="76721-160">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="76721-160">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|<span data-ttu-id="76721-161">x</span><span class="sxs-lookup"><span data-stu-id="76721-161">x</span></span>|<span data-ttu-id="76721-162">x</span><span class="sxs-lookup"><span data-stu-id="76721-162">x</span></span>|<span data-ttu-id="76721-163">x</span><span class="sxs-lookup"><span data-stu-id="76721-163">x</span></span>|
|[<span data-ttu-id="76721-164">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="76721-164">SupportUrl</span></span>](supporturl.md)|<span data-ttu-id="76721-165">x</span><span class="sxs-lookup"><span data-stu-id="76721-165">x</span></span>|<span data-ttu-id="76721-166">x</span><span class="sxs-lookup"><span data-stu-id="76721-166">x</span></span>|<span data-ttu-id="76721-167">x</span><span class="sxs-lookup"><span data-stu-id="76721-167">x</span></span>|
|[<span data-ttu-id="76721-168">AppDomains</span><span class="sxs-lookup"><span data-stu-id="76721-168">AppDomains</span></span>](appdomains.md)|<span data-ttu-id="76721-169">x</span><span class="sxs-lookup"><span data-stu-id="76721-169">x</span></span>|<span data-ttu-id="76721-170">x</span><span class="sxs-lookup"><span data-stu-id="76721-170">x</span></span>|<span data-ttu-id="76721-171">x</span><span class="sxs-lookup"><span data-stu-id="76721-171">x</span></span>|
|[<span data-ttu-id="76721-172">Hosts</span><span class="sxs-lookup"><span data-stu-id="76721-172">Hosts</span></span>](hosts.md)|<span data-ttu-id="76721-173">x</span><span class="sxs-lookup"><span data-stu-id="76721-173">x</span></span>|<span data-ttu-id="76721-174">x</span><span class="sxs-lookup"><span data-stu-id="76721-174">x</span></span>|<span data-ttu-id="76721-175">x</span><span class="sxs-lookup"><span data-stu-id="76721-175">x</span></span>|
|[<span data-ttu-id="76721-176">Requirements</span><span class="sxs-lookup"><span data-stu-id="76721-176">Requirements</span></span>](requirements.md)|<span data-ttu-id="76721-177">x</span><span class="sxs-lookup"><span data-stu-id="76721-177">x</span></span>|<span data-ttu-id="76721-178">x</span><span class="sxs-lookup"><span data-stu-id="76721-178">x</span></span>|<span data-ttu-id="76721-179">x</span><span class="sxs-lookup"><span data-stu-id="76721-179">x</span></span>|
|[<span data-ttu-id="76721-180">AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="76721-180">AllowSnapshot</span></span>](allowsnapshot.md)|<span data-ttu-id="76721-181">x</span><span class="sxs-lookup"><span data-stu-id="76721-181">x</span></span>|||
|[<span data-ttu-id="76721-182">アクセス許可</span><span class="sxs-lookup"><span data-stu-id="76721-182">Permissions</span></span>](permissions.md)||<span data-ttu-id="76721-183">x</span><span class="sxs-lookup"><span data-stu-id="76721-183">x</span></span>||
|[<span data-ttu-id="76721-184">DisableEntityHighlighting</span><span class="sxs-lookup"><span data-stu-id="76721-184">DisableEntityHighlighting</span></span>](disableentityhighlighting.md)||<span data-ttu-id="76721-185">x</span><span class="sxs-lookup"><span data-stu-id="76721-185">x</span></span>||
|[<span data-ttu-id="76721-186">Dictionary</span><span class="sxs-lookup"><span data-stu-id="76721-186">Dictionary</span></span>](dictionary.md)|||<span data-ttu-id="76721-187">x</span><span class="sxs-lookup"><span data-stu-id="76721-187">x</span></span>|
|[<span data-ttu-id="76721-188">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="76721-188">VersionOverrides</span></span>](versionoverrides.md)||<span data-ttu-id="76721-189">x</span><span class="sxs-lookup"><span data-stu-id="76721-189">x</span></span>||

## <a name="attributes"></a><span data-ttu-id="76721-190">属性</span><span class="sxs-lookup"><span data-stu-id="76721-190">Attributes</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="76721-191">xmlns</span><span class="sxs-lookup"><span data-stu-id="76721-191">xmlns</span></span>|<span data-ttu-id="76721-p101">Office アドイン マニフェストの名前空間とスキーマ バージョンを定義します。この属性は常に `"http://schemas.microsoft.com/office/appforoffice/1.1"` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="76721-p101">Defines the Office Add-in manifest namespace and schema version. This attribute should always be set to  `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span></span>|
|<span data-ttu-id="76721-194">xmlns:xsi</span><span class="sxs-lookup"><span data-stu-id="76721-194">xmlns:xsi</span></span>|<span data-ttu-id="76721-p102">XMLSchema インスタンスを定義します。この属性は常に `"http://www.w3.org/2001/XMLSchema-instance"` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="76721-p102">Defines the XMLSchema instance. This attribute should always be set to  `"http://www.w3.org/2001/XMLSchema-instance"`</span></span>|
|<span data-ttu-id="76721-197">xsi:type</span><span class="sxs-lookup"><span data-stu-id="76721-197">xsi:type</span></span>|<span data-ttu-id="76721-p103">Office アドインの種類を定義します。この属性は、`"ContentApp"`、`"MailApp"`、または `"TaskPaneApp"` のいずれかに設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="76721-p103">Defines the kind of Office Add-in. This attribute should be set to one of:  `"ContentApp"`,  `"MailApp"`, or  `"TaskPaneApp"`</span></span>|
