---
title: マニフェスト ファイルの OfficeApp 要素
description: OfficeApp 要素は、Office アドインマニフェストのルート要素です。
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: c5786343173d0e130df4b786f28a8689d573b6ca
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996320"
---
# <a name="officeapp-element"></a><span data-ttu-id="16659-103">OfficeApp 要素</span><span class="sxs-lookup"><span data-stu-id="16659-103">OfficeApp element</span></span>

<span data-ttu-id="16659-104">Office アドインのマニフェストのルート要素。</span><span class="sxs-lookup"><span data-stu-id="16659-104">The root element in the manifest of an Office Add-in.</span></span>

<span data-ttu-id="16659-105">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="16659-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="16659-106">構文</span><span class="sxs-lookup"><span data-stu-id="16659-106">Syntax</span></span>

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="16659-107">次に含まれる</span><span class="sxs-lookup"><span data-stu-id="16659-107">Contained in</span></span>

 <span data-ttu-id="16659-108">_none_</span><span class="sxs-lookup"><span data-stu-id="16659-108">_none_</span></span>

## <a name="must-contain"></a><span data-ttu-id="16659-109">含める必要があるもの</span><span class="sxs-lookup"><span data-stu-id="16659-109">Must contain</span></span>

|<span data-ttu-id="16659-110">要素</span><span class="sxs-lookup"><span data-stu-id="16659-110">Element</span></span>|<span data-ttu-id="16659-111">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="16659-111">Content</span></span>|<span data-ttu-id="16659-112">メール</span><span class="sxs-lookup"><span data-stu-id="16659-112">Mail</span></span>|<span data-ttu-id="16659-113">TaskPane</span><span class="sxs-lookup"><span data-stu-id="16659-113">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="16659-114">Id</span><span class="sxs-lookup"><span data-stu-id="16659-114">Id</span></span>](id.md)|<span data-ttu-id="16659-115">x</span><span class="sxs-lookup"><span data-stu-id="16659-115">x</span></span>|<span data-ttu-id="16659-116">x</span><span class="sxs-lookup"><span data-stu-id="16659-116">x</span></span>|<span data-ttu-id="16659-117">x</span><span class="sxs-lookup"><span data-stu-id="16659-117">x</span></span>|
|[<span data-ttu-id="16659-118">バージョン</span><span class="sxs-lookup"><span data-stu-id="16659-118">Version</span></span>](version.md)|<span data-ttu-id="16659-119">x</span><span class="sxs-lookup"><span data-stu-id="16659-119">x</span></span>|<span data-ttu-id="16659-120">x</span><span class="sxs-lookup"><span data-stu-id="16659-120">x</span></span>|<span data-ttu-id="16659-121">x</span><span class="sxs-lookup"><span data-stu-id="16659-121">x</span></span>|
|[<span data-ttu-id="16659-122">ProviderName</span><span class="sxs-lookup"><span data-stu-id="16659-122">ProviderName</span></span>](providername.md)|<span data-ttu-id="16659-123">x</span><span class="sxs-lookup"><span data-stu-id="16659-123">x</span></span>|<span data-ttu-id="16659-124">x</span><span class="sxs-lookup"><span data-stu-id="16659-124">x</span></span>|<span data-ttu-id="16659-125">x</span><span class="sxs-lookup"><span data-stu-id="16659-125">x</span></span>|
|[<span data-ttu-id="16659-126">DefaultLocale</span><span class="sxs-lookup"><span data-stu-id="16659-126">DefaultLocale</span></span>](defaultlocale.md)|<span data-ttu-id="16659-127">x</span><span class="sxs-lookup"><span data-stu-id="16659-127">x</span></span>|<span data-ttu-id="16659-128">x</span><span class="sxs-lookup"><span data-stu-id="16659-128">x</span></span>|<span data-ttu-id="16659-129">x</span><span class="sxs-lookup"><span data-stu-id="16659-129">x</span></span>|
|[<span data-ttu-id="16659-130">DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="16659-130">DefaultSettings</span></span>](defaultsettings.md)|<span data-ttu-id="16659-131">x</span><span class="sxs-lookup"><span data-stu-id="16659-131">x</span></span>||<span data-ttu-id="16659-132">x</span><span class="sxs-lookup"><span data-stu-id="16659-132">x</span></span>|
|[<span data-ttu-id="16659-133">DisplayName</span><span class="sxs-lookup"><span data-stu-id="16659-133">DisplayName</span></span>](displayname.md)|<span data-ttu-id="16659-134">x</span><span class="sxs-lookup"><span data-stu-id="16659-134">x</span></span>|<span data-ttu-id="16659-135">x</span><span class="sxs-lookup"><span data-stu-id="16659-135">x</span></span>|<span data-ttu-id="16659-136">x</span><span class="sxs-lookup"><span data-stu-id="16659-136">x</span></span>|
|[<span data-ttu-id="16659-137">説明</span><span class="sxs-lookup"><span data-stu-id="16659-137">Description</span></span>](description.md)|<span data-ttu-id="16659-138">x</span><span class="sxs-lookup"><span data-stu-id="16659-138">x</span></span>|<span data-ttu-id="16659-139">x</span><span class="sxs-lookup"><span data-stu-id="16659-139">x</span></span>|<span data-ttu-id="16659-140">x</span><span class="sxs-lookup"><span data-stu-id="16659-140">x</span></span>|
|[<span data-ttu-id="16659-141">FormSettings</span><span class="sxs-lookup"><span data-stu-id="16659-141">FormSettings</span></span>](formsettings.md)||<span data-ttu-id="16659-142">x</span><span class="sxs-lookup"><span data-stu-id="16659-142">x</span></span>||
|[<span data-ttu-id="16659-143">アクセス許可</span><span class="sxs-lookup"><span data-stu-id="16659-143">Permissions</span></span>](permissions.md)|<span data-ttu-id="16659-144">x</span><span class="sxs-lookup"><span data-stu-id="16659-144">x</span></span>||<span data-ttu-id="16659-145">x</span><span class="sxs-lookup"><span data-stu-id="16659-145">x</span></span>|
|[<span data-ttu-id="16659-146">Rule</span><span class="sxs-lookup"><span data-stu-id="16659-146">Rule</span></span>](rule.md)||<span data-ttu-id="16659-147">x</span><span class="sxs-lookup"><span data-stu-id="16659-147">x</span></span>||

## <a name="can-contain"></a><span data-ttu-id="16659-148">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="16659-148">Can contain</span></span>

|<span data-ttu-id="16659-149">要素</span><span class="sxs-lookup"><span data-stu-id="16659-149">Element</span></span>|<span data-ttu-id="16659-150">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="16659-150">Content</span></span>|<span data-ttu-id="16659-151">メール</span><span class="sxs-lookup"><span data-stu-id="16659-151">Mail</span></span>|<span data-ttu-id="16659-152">TaskPane</span><span class="sxs-lookup"><span data-stu-id="16659-152">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="16659-153">AlternateId</span><span class="sxs-lookup"><span data-stu-id="16659-153">AlternateId</span></span>](alternateid.md)|<span data-ttu-id="16659-154">x</span><span class="sxs-lookup"><span data-stu-id="16659-154">x</span></span>|<span data-ttu-id="16659-155">x</span><span class="sxs-lookup"><span data-stu-id="16659-155">x</span></span>|<span data-ttu-id="16659-156">x</span><span class="sxs-lookup"><span data-stu-id="16659-156">x</span></span>|
|[<span data-ttu-id="16659-157">IconUrl</span><span class="sxs-lookup"><span data-stu-id="16659-157">IconUrl</span></span>](iconurl.md)|<span data-ttu-id="16659-158">x</span><span class="sxs-lookup"><span data-stu-id="16659-158">x</span></span>|<span data-ttu-id="16659-159">x</span><span class="sxs-lookup"><span data-stu-id="16659-159">x</span></span>|<span data-ttu-id="16659-160">x</span><span class="sxs-lookup"><span data-stu-id="16659-160">x</span></span>|
|[<span data-ttu-id="16659-161">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="16659-161">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|<span data-ttu-id="16659-162">x</span><span class="sxs-lookup"><span data-stu-id="16659-162">x</span></span>|<span data-ttu-id="16659-163">x</span><span class="sxs-lookup"><span data-stu-id="16659-163">x</span></span>|<span data-ttu-id="16659-164">x</span><span class="sxs-lookup"><span data-stu-id="16659-164">x</span></span>|
|[<span data-ttu-id="16659-165">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="16659-165">SupportUrl</span></span>](supporturl.md)|<span data-ttu-id="16659-166">x</span><span class="sxs-lookup"><span data-stu-id="16659-166">x</span></span>|<span data-ttu-id="16659-167">x</span><span class="sxs-lookup"><span data-stu-id="16659-167">x</span></span>|<span data-ttu-id="16659-168">x</span><span class="sxs-lookup"><span data-stu-id="16659-168">x</span></span>|
|[<span data-ttu-id="16659-169">AppDomains</span><span class="sxs-lookup"><span data-stu-id="16659-169">AppDomains</span></span>](appdomains.md)|<span data-ttu-id="16659-170">x</span><span class="sxs-lookup"><span data-stu-id="16659-170">x</span></span>|<span data-ttu-id="16659-171">x</span><span class="sxs-lookup"><span data-stu-id="16659-171">x</span></span>|<span data-ttu-id="16659-172">x</span><span class="sxs-lookup"><span data-stu-id="16659-172">x</span></span>|
|[<span data-ttu-id="16659-173">Hosts</span><span class="sxs-lookup"><span data-stu-id="16659-173">Hosts</span></span>](hosts.md)|<span data-ttu-id="16659-174">x</span><span class="sxs-lookup"><span data-stu-id="16659-174">x</span></span>|<span data-ttu-id="16659-175">x</span><span class="sxs-lookup"><span data-stu-id="16659-175">x</span></span>|<span data-ttu-id="16659-176">x</span><span class="sxs-lookup"><span data-stu-id="16659-176">x</span></span>|
|[<span data-ttu-id="16659-177">Requirements</span><span class="sxs-lookup"><span data-stu-id="16659-177">Requirements</span></span>](requirements.md)|<span data-ttu-id="16659-178">x</span><span class="sxs-lookup"><span data-stu-id="16659-178">x</span></span>|<span data-ttu-id="16659-179">x</span><span class="sxs-lookup"><span data-stu-id="16659-179">x</span></span>|<span data-ttu-id="16659-180">x</span><span class="sxs-lookup"><span data-stu-id="16659-180">x</span></span>|
|[<span data-ttu-id="16659-181">AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="16659-181">AllowSnapshot</span></span>](allowsnapshot.md)|<span data-ttu-id="16659-182">x</span><span class="sxs-lookup"><span data-stu-id="16659-182">x</span></span>|||
|[<span data-ttu-id="16659-183">アクセス許可</span><span class="sxs-lookup"><span data-stu-id="16659-183">Permissions</span></span>](permissions.md)||<span data-ttu-id="16659-184">x</span><span class="sxs-lookup"><span data-stu-id="16659-184">x</span></span>||
|[<span data-ttu-id="16659-185">DisableEntityHighlighting</span><span class="sxs-lookup"><span data-stu-id="16659-185">DisableEntityHighlighting</span></span>](disableentityhighlighting.md)||<span data-ttu-id="16659-186">x</span><span class="sxs-lookup"><span data-stu-id="16659-186">x</span></span>||
|[<span data-ttu-id="16659-187">Dictionary</span><span class="sxs-lookup"><span data-stu-id="16659-187">Dictionary</span></span>](dictionary.md)|||<span data-ttu-id="16659-188">x</span><span class="sxs-lookup"><span data-stu-id="16659-188">x</span></span>|
|[<span data-ttu-id="16659-189">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="16659-189">VersionOverrides</span></span>](versionoverrides.md)|<span data-ttu-id="16659-190">x</span><span class="sxs-lookup"><span data-stu-id="16659-190">x</span></span>|<span data-ttu-id="16659-191">x</span><span class="sxs-lookup"><span data-stu-id="16659-191">x</span></span>|<span data-ttu-id="16659-192">x</span><span class="sxs-lookup"><span data-stu-id="16659-192">x</span></span>|
|[<span data-ttu-id="16659-193">ExtendedOverrides</span><span class="sxs-lookup"><span data-stu-id="16659-193">ExtendedOverrides</span></span>](extendedoverrides.md)|||<span data-ttu-id="16659-194">x</span><span class="sxs-lookup"><span data-stu-id="16659-194">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="16659-195">属性</span><span class="sxs-lookup"><span data-stu-id="16659-195">Attributes</span></span>

|<span data-ttu-id="16659-196">属性</span><span class="sxs-lookup"><span data-stu-id="16659-196">Attribute</span></span>|<span data-ttu-id="16659-197">説明</span><span class="sxs-lookup"><span data-stu-id="16659-197">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="16659-198">xmlns</span><span class="sxs-lookup"><span data-stu-id="16659-198">xmlns</span></span>|<span data-ttu-id="16659-p101">Office アドイン マニフェストの名前空間とスキーマ バージョンを定義します。この属性は常に `"http://schemas.microsoft.com/office/appforoffice/1.1"` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="16659-p101">Defines the Office Add-in manifest namespace and schema version. This attribute should always be set to  `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span></span>|
|<span data-ttu-id="16659-201">xmlns:xsi</span><span class="sxs-lookup"><span data-stu-id="16659-201">xmlns:xsi</span></span>|<span data-ttu-id="16659-p102">XMLSchema インスタンスを定義します。この属性は常に `"http://www.w3.org/2001/XMLSchema-instance"` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="16659-p102">Defines the XMLSchema instance. This attribute should always be set to  `"http://www.w3.org/2001/XMLSchema-instance"`</span></span>|
|<span data-ttu-id="16659-204">xsi:type</span><span class="sxs-lookup"><span data-stu-id="16659-204">xsi:type</span></span>|<span data-ttu-id="16659-p103">Office アドインの種類を定義します。この属性は、`"ContentApp"`、`"MailApp"`、または `"TaskPaneApp"` のいずれかに設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="16659-p103">Defines the kind of Office Add-in. This attribute should be set to one of:  `"ContentApp"`,  `"MailApp"`, or  `"TaskPaneApp"`</span></span>|
