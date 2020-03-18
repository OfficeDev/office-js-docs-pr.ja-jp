---
title: マニフェスト ファイルの OfficeApp 要素
description: OfficeApp 要素は、Office アドインマニフェストのルート要素です。
ms.date: 02/04/2020
localization_priority: Normal
ms.openlocfilehash: 038933f2d06ee5f485dbdb7dd7abdbd95fb97c7d
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720597"
---
# <a name="officeapp-element"></a><span data-ttu-id="35cd4-103">OfficeApp 要素</span><span class="sxs-lookup"><span data-stu-id="35cd4-103">OfficeApp element</span></span>

<span data-ttu-id="35cd4-104">Office アドインのマニフェストのルート要素。</span><span class="sxs-lookup"><span data-stu-id="35cd4-104">The root element in the manifest of an Office Add-in.</span></span>

<span data-ttu-id="35cd4-105">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="35cd4-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="35cd4-106">構文</span><span class="sxs-lookup"><span data-stu-id="35cd4-106">Syntax</span></span>

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="35cd4-107">次に含まれる</span><span class="sxs-lookup"><span data-stu-id="35cd4-107">Contained in</span></span>

 <span data-ttu-id="35cd4-108">_none_</span><span class="sxs-lookup"><span data-stu-id="35cd4-108">_none_</span></span>

## <a name="must-contain"></a><span data-ttu-id="35cd4-109">含める必要があるもの</span><span class="sxs-lookup"><span data-stu-id="35cd4-109">Must contain</span></span>

|<span data-ttu-id="35cd4-110">**要素**</span><span class="sxs-lookup"><span data-stu-id="35cd4-110">**Element**</span></span>|<span data-ttu-id="35cd4-111">**コンテンツ**</span><span class="sxs-lookup"><span data-stu-id="35cd4-111">**Content**</span></span>|<span data-ttu-id="35cd4-112">**メール**</span><span class="sxs-lookup"><span data-stu-id="35cd4-112">**Mail**</span></span>|<span data-ttu-id="35cd4-113">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="35cd4-113">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="35cd4-114">Id</span><span class="sxs-lookup"><span data-stu-id="35cd4-114">Id</span></span>](id.md)|<span data-ttu-id="35cd4-115">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-115">x</span></span>|<span data-ttu-id="35cd4-116">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-116">x</span></span>|<span data-ttu-id="35cd4-117">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-117">x</span></span>|
|[<span data-ttu-id="35cd4-118">バージョン</span><span class="sxs-lookup"><span data-stu-id="35cd4-118">Version</span></span>](version.md)|<span data-ttu-id="35cd4-119">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-119">x</span></span>|<span data-ttu-id="35cd4-120">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-120">x</span></span>|<span data-ttu-id="35cd4-121">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-121">x</span></span>|
|[<span data-ttu-id="35cd4-122">ProviderName</span><span class="sxs-lookup"><span data-stu-id="35cd4-122">ProviderName</span></span>](providername.md)|<span data-ttu-id="35cd4-123">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-123">x</span></span>|<span data-ttu-id="35cd4-124">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-124">x</span></span>|<span data-ttu-id="35cd4-125">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-125">x</span></span>|
|[<span data-ttu-id="35cd4-126">DefaultLocale</span><span class="sxs-lookup"><span data-stu-id="35cd4-126">DefaultLocale</span></span>](defaultlocale.md)|<span data-ttu-id="35cd4-127">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-127">x</span></span>|<span data-ttu-id="35cd4-128">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-128">x</span></span>|<span data-ttu-id="35cd4-129">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-129">x</span></span>|
|[<span data-ttu-id="35cd4-130">DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="35cd4-130">DefaultSettings</span></span>](defaultsettings.md)|<span data-ttu-id="35cd4-131">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-131">x</span></span>||<span data-ttu-id="35cd4-132">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-132">x</span></span>|
|[<span data-ttu-id="35cd4-133">DisplayName</span><span class="sxs-lookup"><span data-stu-id="35cd4-133">DisplayName</span></span>](displayname.md)|<span data-ttu-id="35cd4-134">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-134">x</span></span>|<span data-ttu-id="35cd4-135">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-135">x</span></span>|<span data-ttu-id="35cd4-136">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-136">x</span></span>|
|[<span data-ttu-id="35cd4-137">説明</span><span class="sxs-lookup"><span data-stu-id="35cd4-137">Description</span></span>](description.md)|<span data-ttu-id="35cd4-138">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-138">x</span></span>|<span data-ttu-id="35cd4-139">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-139">x</span></span>|<span data-ttu-id="35cd4-140">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-140">x</span></span>|
|[<span data-ttu-id="35cd4-141">FormSettings</span><span class="sxs-lookup"><span data-stu-id="35cd4-141">FormSettings</span></span>](formsettings.md)||<span data-ttu-id="35cd4-142">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-142">x</span></span>||
|[<span data-ttu-id="35cd4-143">アクセス許可</span><span class="sxs-lookup"><span data-stu-id="35cd4-143">Permissions</span></span>](permissions.md)|<span data-ttu-id="35cd4-144">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-144">x</span></span>||<span data-ttu-id="35cd4-145">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-145">x</span></span>|
|[<span data-ttu-id="35cd4-146">Rule</span><span class="sxs-lookup"><span data-stu-id="35cd4-146">Rule</span></span>](rule.md)||<span data-ttu-id="35cd4-147">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-147">x</span></span>||

## <a name="can-contain"></a><span data-ttu-id="35cd4-148">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="35cd4-148">Can contain</span></span>

|<span data-ttu-id="35cd4-149">**Element**</span><span class="sxs-lookup"><span data-stu-id="35cd4-149">**Element**</span></span>|<span data-ttu-id="35cd4-150">**コンテンツ**</span><span class="sxs-lookup"><span data-stu-id="35cd4-150">**Content**</span></span>|<span data-ttu-id="35cd4-151">**メール**</span><span class="sxs-lookup"><span data-stu-id="35cd4-151">**Mail**</span></span>|<span data-ttu-id="35cd4-152">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="35cd4-152">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="35cd4-153">AlternateId</span><span class="sxs-lookup"><span data-stu-id="35cd4-153">AlternateId</span></span>](alternateid.md)|<span data-ttu-id="35cd4-154">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-154">x</span></span>|<span data-ttu-id="35cd4-155">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-155">x</span></span>|<span data-ttu-id="35cd4-156">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-156">x</span></span>|
|[<span data-ttu-id="35cd4-157">IconUrl</span><span class="sxs-lookup"><span data-stu-id="35cd4-157">IconUrl</span></span>](iconurl.md)|<span data-ttu-id="35cd4-158">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-158">x</span></span>|<span data-ttu-id="35cd4-159">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-159">x</span></span>|<span data-ttu-id="35cd4-160">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-160">x</span></span>|
|[<span data-ttu-id="35cd4-161">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="35cd4-161">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|<span data-ttu-id="35cd4-162">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-162">x</span></span>|<span data-ttu-id="35cd4-163">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-163">x</span></span>|<span data-ttu-id="35cd4-164">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-164">x</span></span>|
|[<span data-ttu-id="35cd4-165">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="35cd4-165">SupportUrl</span></span>](supporturl.md)|<span data-ttu-id="35cd4-166">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-166">x</span></span>|<span data-ttu-id="35cd4-167">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-167">x</span></span>|<span data-ttu-id="35cd4-168">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-168">x</span></span>|
|[<span data-ttu-id="35cd4-169">AppDomains</span><span class="sxs-lookup"><span data-stu-id="35cd4-169">AppDomains</span></span>](appdomains.md)|<span data-ttu-id="35cd4-170">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-170">x</span></span>|<span data-ttu-id="35cd4-171">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-171">x</span></span>|<span data-ttu-id="35cd4-172">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-172">x</span></span>|
|[<span data-ttu-id="35cd4-173">Hosts</span><span class="sxs-lookup"><span data-stu-id="35cd4-173">Hosts</span></span>](hosts.md)|<span data-ttu-id="35cd4-174">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-174">x</span></span>|<span data-ttu-id="35cd4-175">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-175">x</span></span>|<span data-ttu-id="35cd4-176">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-176">x</span></span>|
|[<span data-ttu-id="35cd4-177">Requirements</span><span class="sxs-lookup"><span data-stu-id="35cd4-177">Requirements</span></span>](requirements.md)|<span data-ttu-id="35cd4-178">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-178">x</span></span>|<span data-ttu-id="35cd4-179">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-179">x</span></span>|<span data-ttu-id="35cd4-180">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-180">x</span></span>|
|[<span data-ttu-id="35cd4-181">AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="35cd4-181">AllowSnapshot</span></span>](allowsnapshot.md)|<span data-ttu-id="35cd4-182">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-182">x</span></span>|||
|[<span data-ttu-id="35cd4-183">アクセス許可</span><span class="sxs-lookup"><span data-stu-id="35cd4-183">Permissions</span></span>](permissions.md)||<span data-ttu-id="35cd4-184">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-184">x</span></span>||
|[<span data-ttu-id="35cd4-185">DisableEntityHighlighting</span><span class="sxs-lookup"><span data-stu-id="35cd4-185">DisableEntityHighlighting</span></span>](disableentityhighlighting.md)||<span data-ttu-id="35cd4-186">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-186">x</span></span>||
|[<span data-ttu-id="35cd4-187">Dictionary</span><span class="sxs-lookup"><span data-stu-id="35cd4-187">Dictionary</span></span>](dictionary.md)|||<span data-ttu-id="35cd4-188">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-188">x</span></span>|
|[<span data-ttu-id="35cd4-189">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="35cd4-189">VersionOverrides</span></span>](versionoverrides.md)|<span data-ttu-id="35cd4-190">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-190">x</span></span>|<span data-ttu-id="35cd4-191">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-191">x</span></span>|<span data-ttu-id="35cd4-192">x</span><span class="sxs-lookup"><span data-stu-id="35cd4-192">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="35cd4-193">属性</span><span class="sxs-lookup"><span data-stu-id="35cd4-193">Attributes</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="35cd4-194">xmlns</span><span class="sxs-lookup"><span data-stu-id="35cd4-194">xmlns</span></span>|<span data-ttu-id="35cd4-p101">Office アドイン マニフェストの名前空間とスキーマ バージョンを定義します。この属性は常に `"http://schemas.microsoft.com/office/appforoffice/1.1"` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="35cd4-p101">Defines the Office Add-in manifest namespace and schema version. This attribute should always be set to  `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span></span>|
|<span data-ttu-id="35cd4-197">xmlns:xsi</span><span class="sxs-lookup"><span data-stu-id="35cd4-197">xmlns:xsi</span></span>|<span data-ttu-id="35cd4-p102">XMLSchema インスタンスを定義します。この属性は常に `"http://www.w3.org/2001/XMLSchema-instance"` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="35cd4-p102">Defines the XMLSchema instance. This attribute should always be set to  `"http://www.w3.org/2001/XMLSchema-instance"`</span></span>|
|<span data-ttu-id="35cd4-200">xsi:type</span><span class="sxs-lookup"><span data-stu-id="35cd4-200">xsi:type</span></span>|<span data-ttu-id="35cd4-p103">Office アドインの種類を定義します。この属性は、`"ContentApp"`、`"MailApp"`、または `"TaskPaneApp"` のいずれかに設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="35cd4-p103">Defines the kind of Office Add-in. This attribute should be set to one of:  `"ContentApp"`,  `"MailApp"`, or  `"TaskPaneApp"`</span></span>|
