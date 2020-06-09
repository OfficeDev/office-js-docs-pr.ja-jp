---
title: マニフェスト ファイルの OfficeApp 要素
description: OfficeApp 要素は、Office アドインマニフェストのルート要素です。
ms.date: 02/04/2020
localization_priority: Normal
ms.openlocfilehash: b6f3102a97794a19366b06734789e01fc4bc4f9d
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611527"
---
# <a name="officeapp-element"></a><span data-ttu-id="f642e-103">OfficeApp 要素</span><span class="sxs-lookup"><span data-stu-id="f642e-103">OfficeApp element</span></span>

<span data-ttu-id="f642e-104">Office アドインのマニフェストのルート要素。</span><span class="sxs-lookup"><span data-stu-id="f642e-104">The root element in the manifest of an Office Add-in.</span></span>

<span data-ttu-id="f642e-105">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="f642e-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="f642e-106">構文</span><span class="sxs-lookup"><span data-stu-id="f642e-106">Syntax</span></span>

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="f642e-107">次に含まれる</span><span class="sxs-lookup"><span data-stu-id="f642e-107">Contained in</span></span>

 <span data-ttu-id="f642e-108">_none_</span><span class="sxs-lookup"><span data-stu-id="f642e-108">_none_</span></span>

## <a name="must-contain"></a><span data-ttu-id="f642e-109">含める必要があるもの</span><span class="sxs-lookup"><span data-stu-id="f642e-109">Must contain</span></span>

|<span data-ttu-id="f642e-110">**要素**</span><span class="sxs-lookup"><span data-stu-id="f642e-110">**Element**</span></span>|<span data-ttu-id="f642e-111">**コンテンツ**</span><span class="sxs-lookup"><span data-stu-id="f642e-111">**Content**</span></span>|<span data-ttu-id="f642e-112">**メール**</span><span class="sxs-lookup"><span data-stu-id="f642e-112">**Mail**</span></span>|<span data-ttu-id="f642e-113">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="f642e-113">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="f642e-114">Id</span><span class="sxs-lookup"><span data-stu-id="f642e-114">Id</span></span>](id.md)|<span data-ttu-id="f642e-115">x</span><span class="sxs-lookup"><span data-stu-id="f642e-115">x</span></span>|<span data-ttu-id="f642e-116">x</span><span class="sxs-lookup"><span data-stu-id="f642e-116">x</span></span>|<span data-ttu-id="f642e-117">x</span><span class="sxs-lookup"><span data-stu-id="f642e-117">x</span></span>|
|[<span data-ttu-id="f642e-118">バージョン</span><span class="sxs-lookup"><span data-stu-id="f642e-118">Version</span></span>](version.md)|<span data-ttu-id="f642e-119">x</span><span class="sxs-lookup"><span data-stu-id="f642e-119">x</span></span>|<span data-ttu-id="f642e-120">x</span><span class="sxs-lookup"><span data-stu-id="f642e-120">x</span></span>|<span data-ttu-id="f642e-121">x</span><span class="sxs-lookup"><span data-stu-id="f642e-121">x</span></span>|
|[<span data-ttu-id="f642e-122">ProviderName</span><span class="sxs-lookup"><span data-stu-id="f642e-122">ProviderName</span></span>](providername.md)|<span data-ttu-id="f642e-123">x</span><span class="sxs-lookup"><span data-stu-id="f642e-123">x</span></span>|<span data-ttu-id="f642e-124">x</span><span class="sxs-lookup"><span data-stu-id="f642e-124">x</span></span>|<span data-ttu-id="f642e-125">x</span><span class="sxs-lookup"><span data-stu-id="f642e-125">x</span></span>|
|[<span data-ttu-id="f642e-126">DefaultLocale</span><span class="sxs-lookup"><span data-stu-id="f642e-126">DefaultLocale</span></span>](defaultlocale.md)|<span data-ttu-id="f642e-127">x</span><span class="sxs-lookup"><span data-stu-id="f642e-127">x</span></span>|<span data-ttu-id="f642e-128">x</span><span class="sxs-lookup"><span data-stu-id="f642e-128">x</span></span>|<span data-ttu-id="f642e-129">x</span><span class="sxs-lookup"><span data-stu-id="f642e-129">x</span></span>|
|[<span data-ttu-id="f642e-130">DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="f642e-130">DefaultSettings</span></span>](defaultsettings.md)|<span data-ttu-id="f642e-131">x</span><span class="sxs-lookup"><span data-stu-id="f642e-131">x</span></span>||<span data-ttu-id="f642e-132">x</span><span class="sxs-lookup"><span data-stu-id="f642e-132">x</span></span>|
|[<span data-ttu-id="f642e-133">DisplayName</span><span class="sxs-lookup"><span data-stu-id="f642e-133">DisplayName</span></span>](displayname.md)|<span data-ttu-id="f642e-134">x</span><span class="sxs-lookup"><span data-stu-id="f642e-134">x</span></span>|<span data-ttu-id="f642e-135">x</span><span class="sxs-lookup"><span data-stu-id="f642e-135">x</span></span>|<span data-ttu-id="f642e-136">x</span><span class="sxs-lookup"><span data-stu-id="f642e-136">x</span></span>|
|[<span data-ttu-id="f642e-137">説明</span><span class="sxs-lookup"><span data-stu-id="f642e-137">Description</span></span>](description.md)|<span data-ttu-id="f642e-138">x</span><span class="sxs-lookup"><span data-stu-id="f642e-138">x</span></span>|<span data-ttu-id="f642e-139">x</span><span class="sxs-lookup"><span data-stu-id="f642e-139">x</span></span>|<span data-ttu-id="f642e-140">x</span><span class="sxs-lookup"><span data-stu-id="f642e-140">x</span></span>|
|[<span data-ttu-id="f642e-141">FormSettings</span><span class="sxs-lookup"><span data-stu-id="f642e-141">FormSettings</span></span>](formsettings.md)||<span data-ttu-id="f642e-142">x</span><span class="sxs-lookup"><span data-stu-id="f642e-142">x</span></span>||
|[<span data-ttu-id="f642e-143">アクセス許可</span><span class="sxs-lookup"><span data-stu-id="f642e-143">Permissions</span></span>](permissions.md)|<span data-ttu-id="f642e-144">x</span><span class="sxs-lookup"><span data-stu-id="f642e-144">x</span></span>||<span data-ttu-id="f642e-145">x</span><span class="sxs-lookup"><span data-stu-id="f642e-145">x</span></span>|
|[<span data-ttu-id="f642e-146">Rule</span><span class="sxs-lookup"><span data-stu-id="f642e-146">Rule</span></span>](rule.md)||<span data-ttu-id="f642e-147">x</span><span class="sxs-lookup"><span data-stu-id="f642e-147">x</span></span>||

## <a name="can-contain"></a><span data-ttu-id="f642e-148">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="f642e-148">Can contain</span></span>

|<span data-ttu-id="f642e-149">**Element**</span><span class="sxs-lookup"><span data-stu-id="f642e-149">**Element**</span></span>|<span data-ttu-id="f642e-150">**コンテンツ**</span><span class="sxs-lookup"><span data-stu-id="f642e-150">**Content**</span></span>|<span data-ttu-id="f642e-151">**メール**</span><span class="sxs-lookup"><span data-stu-id="f642e-151">**Mail**</span></span>|<span data-ttu-id="f642e-152">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="f642e-152">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="f642e-153">AlternateId</span><span class="sxs-lookup"><span data-stu-id="f642e-153">AlternateId</span></span>](alternateid.md)|<span data-ttu-id="f642e-154">x</span><span class="sxs-lookup"><span data-stu-id="f642e-154">x</span></span>|<span data-ttu-id="f642e-155">x</span><span class="sxs-lookup"><span data-stu-id="f642e-155">x</span></span>|<span data-ttu-id="f642e-156">x</span><span class="sxs-lookup"><span data-stu-id="f642e-156">x</span></span>|
|[<span data-ttu-id="f642e-157">IconUrl</span><span class="sxs-lookup"><span data-stu-id="f642e-157">IconUrl</span></span>](iconurl.md)|<span data-ttu-id="f642e-158">x</span><span class="sxs-lookup"><span data-stu-id="f642e-158">x</span></span>|<span data-ttu-id="f642e-159">x</span><span class="sxs-lookup"><span data-stu-id="f642e-159">x</span></span>|<span data-ttu-id="f642e-160">x</span><span class="sxs-lookup"><span data-stu-id="f642e-160">x</span></span>|
|[<span data-ttu-id="f642e-161">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="f642e-161">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|<span data-ttu-id="f642e-162">x</span><span class="sxs-lookup"><span data-stu-id="f642e-162">x</span></span>|<span data-ttu-id="f642e-163">x</span><span class="sxs-lookup"><span data-stu-id="f642e-163">x</span></span>|<span data-ttu-id="f642e-164">x</span><span class="sxs-lookup"><span data-stu-id="f642e-164">x</span></span>|
|[<span data-ttu-id="f642e-165">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="f642e-165">SupportUrl</span></span>](supporturl.md)|<span data-ttu-id="f642e-166">x</span><span class="sxs-lookup"><span data-stu-id="f642e-166">x</span></span>|<span data-ttu-id="f642e-167">x</span><span class="sxs-lookup"><span data-stu-id="f642e-167">x</span></span>|<span data-ttu-id="f642e-168">x</span><span class="sxs-lookup"><span data-stu-id="f642e-168">x</span></span>|
|[<span data-ttu-id="f642e-169">AppDomains</span><span class="sxs-lookup"><span data-stu-id="f642e-169">AppDomains</span></span>](appdomains.md)|<span data-ttu-id="f642e-170">x</span><span class="sxs-lookup"><span data-stu-id="f642e-170">x</span></span>|<span data-ttu-id="f642e-171">x</span><span class="sxs-lookup"><span data-stu-id="f642e-171">x</span></span>|<span data-ttu-id="f642e-172">x</span><span class="sxs-lookup"><span data-stu-id="f642e-172">x</span></span>|
|[<span data-ttu-id="f642e-173">Hosts</span><span class="sxs-lookup"><span data-stu-id="f642e-173">Hosts</span></span>](hosts.md)|<span data-ttu-id="f642e-174">x</span><span class="sxs-lookup"><span data-stu-id="f642e-174">x</span></span>|<span data-ttu-id="f642e-175">x</span><span class="sxs-lookup"><span data-stu-id="f642e-175">x</span></span>|<span data-ttu-id="f642e-176">x</span><span class="sxs-lookup"><span data-stu-id="f642e-176">x</span></span>|
|[<span data-ttu-id="f642e-177">Requirements</span><span class="sxs-lookup"><span data-stu-id="f642e-177">Requirements</span></span>](requirements.md)|<span data-ttu-id="f642e-178">x</span><span class="sxs-lookup"><span data-stu-id="f642e-178">x</span></span>|<span data-ttu-id="f642e-179">x</span><span class="sxs-lookup"><span data-stu-id="f642e-179">x</span></span>|<span data-ttu-id="f642e-180">x</span><span class="sxs-lookup"><span data-stu-id="f642e-180">x</span></span>|
|[<span data-ttu-id="f642e-181">AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="f642e-181">AllowSnapshot</span></span>](allowsnapshot.md)|<span data-ttu-id="f642e-182">x</span><span class="sxs-lookup"><span data-stu-id="f642e-182">x</span></span>|||
|[<span data-ttu-id="f642e-183">アクセス許可</span><span class="sxs-lookup"><span data-stu-id="f642e-183">Permissions</span></span>](permissions.md)||<span data-ttu-id="f642e-184">x</span><span class="sxs-lookup"><span data-stu-id="f642e-184">x</span></span>||
|[<span data-ttu-id="f642e-185">DisableEntityHighlighting</span><span class="sxs-lookup"><span data-stu-id="f642e-185">DisableEntityHighlighting</span></span>](disableentityhighlighting.md)||<span data-ttu-id="f642e-186">x</span><span class="sxs-lookup"><span data-stu-id="f642e-186">x</span></span>||
|[<span data-ttu-id="f642e-187">Dictionary</span><span class="sxs-lookup"><span data-stu-id="f642e-187">Dictionary</span></span>](dictionary.md)|||<span data-ttu-id="f642e-188">x</span><span class="sxs-lookup"><span data-stu-id="f642e-188">x</span></span>|
|[<span data-ttu-id="f642e-189">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="f642e-189">VersionOverrides</span></span>](versionoverrides.md)|<span data-ttu-id="f642e-190">x</span><span class="sxs-lookup"><span data-stu-id="f642e-190">x</span></span>|<span data-ttu-id="f642e-191">x</span><span class="sxs-lookup"><span data-stu-id="f642e-191">x</span></span>|<span data-ttu-id="f642e-192">x</span><span class="sxs-lookup"><span data-stu-id="f642e-192">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="f642e-193">属性</span><span class="sxs-lookup"><span data-stu-id="f642e-193">Attributes</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="f642e-194">xmlns</span><span class="sxs-lookup"><span data-stu-id="f642e-194">xmlns</span></span>|<span data-ttu-id="f642e-p101">Office アドイン マニフェストの名前空間とスキーマ バージョンを定義します。この属性は常に `"http://schemas.microsoft.com/office/appforoffice/1.1"` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f642e-p101">Defines the Office Add-in manifest namespace and schema version. This attribute should always be set to  `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span></span>|
|<span data-ttu-id="f642e-197">xmlns:xsi</span><span class="sxs-lookup"><span data-stu-id="f642e-197">xmlns:xsi</span></span>|<span data-ttu-id="f642e-p102">XMLSchema インスタンスを定義します。この属性は常に `"http://www.w3.org/2001/XMLSchema-instance"` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f642e-p102">Defines the XMLSchema instance. This attribute should always be set to  `"http://www.w3.org/2001/XMLSchema-instance"`</span></span>|
|<span data-ttu-id="f642e-200">xsi:type</span><span class="sxs-lookup"><span data-stu-id="f642e-200">xsi:type</span></span>|<span data-ttu-id="f642e-p103">Office アドインの種類を定義します。この属性は、`"ContentApp"`、`"MailApp"`、または `"TaskPaneApp"` のいずれかに設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="f642e-p103">Defines the kind of Office Add-in. This attribute should be set to one of:  `"ContentApp"`,  `"MailApp"`, or  `"TaskPaneApp"`</span></span>|
