---
title: マニフェスト ファイルの OfficeApp 要素
description: ''
ms.date: 02/04/2020
localization_priority: Normal
ms.openlocfilehash: 080025e62a56421dff942792f99ee672ce1db69a
ms.sourcegitcommit: c1dbea577ae6183523fb663d364422d2adbc8bcf
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/05/2020
ms.locfileid: "41773580"
---
# <a name="officeapp-element"></a><span data-ttu-id="3c472-102">OfficeApp 要素</span><span class="sxs-lookup"><span data-stu-id="3c472-102">OfficeApp element</span></span>

<span data-ttu-id="3c472-103">Office アドインのマニフェストのルート要素。</span><span class="sxs-lookup"><span data-stu-id="3c472-103">The root element in the manifest of an Office Add-in.</span></span>

<span data-ttu-id="3c472-104">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="3c472-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="3c472-105">構文</span><span class="sxs-lookup"><span data-stu-id="3c472-105">Syntax</span></span>

```XML
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xsi:type= ["ContentApp" |"MailApp"| "TaskPaneApp"]>
  ...
</OfficeApp>
```

## <a name="contained-in"></a><span data-ttu-id="3c472-106">次に含まれる</span><span class="sxs-lookup"><span data-stu-id="3c472-106">Contained in</span></span>

 <span data-ttu-id="3c472-107">_none_</span><span class="sxs-lookup"><span data-stu-id="3c472-107">_none_</span></span>

## <a name="must-contain"></a><span data-ttu-id="3c472-108">含める必要があるもの</span><span class="sxs-lookup"><span data-stu-id="3c472-108">Must contain</span></span>

|<span data-ttu-id="3c472-109">**要素**</span><span class="sxs-lookup"><span data-stu-id="3c472-109">**Element**</span></span>|<span data-ttu-id="3c472-110">**コンテンツ**</span><span class="sxs-lookup"><span data-stu-id="3c472-110">**Content**</span></span>|<span data-ttu-id="3c472-111">**メール**</span><span class="sxs-lookup"><span data-stu-id="3c472-111">**Mail**</span></span>|<span data-ttu-id="3c472-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="3c472-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="3c472-113">Id</span><span class="sxs-lookup"><span data-stu-id="3c472-113">Id</span></span>](id.md)|<span data-ttu-id="3c472-114">x</span><span class="sxs-lookup"><span data-stu-id="3c472-114">x</span></span>|<span data-ttu-id="3c472-115">x</span><span class="sxs-lookup"><span data-stu-id="3c472-115">x</span></span>|<span data-ttu-id="3c472-116">x</span><span class="sxs-lookup"><span data-stu-id="3c472-116">x</span></span>|
|[<span data-ttu-id="3c472-117">バージョン</span><span class="sxs-lookup"><span data-stu-id="3c472-117">Version</span></span>](version.md)|<span data-ttu-id="3c472-118">x</span><span class="sxs-lookup"><span data-stu-id="3c472-118">x</span></span>|<span data-ttu-id="3c472-119">x</span><span class="sxs-lookup"><span data-stu-id="3c472-119">x</span></span>|<span data-ttu-id="3c472-120">x</span><span class="sxs-lookup"><span data-stu-id="3c472-120">x</span></span>|
|[<span data-ttu-id="3c472-121">ProviderName</span><span class="sxs-lookup"><span data-stu-id="3c472-121">ProviderName</span></span>](providername.md)|<span data-ttu-id="3c472-122">x</span><span class="sxs-lookup"><span data-stu-id="3c472-122">x</span></span>|<span data-ttu-id="3c472-123">x</span><span class="sxs-lookup"><span data-stu-id="3c472-123">x</span></span>|<span data-ttu-id="3c472-124">x</span><span class="sxs-lookup"><span data-stu-id="3c472-124">x</span></span>|
|[<span data-ttu-id="3c472-125">DefaultLocale</span><span class="sxs-lookup"><span data-stu-id="3c472-125">DefaultLocale</span></span>](defaultlocale.md)|<span data-ttu-id="3c472-126">x</span><span class="sxs-lookup"><span data-stu-id="3c472-126">x</span></span>|<span data-ttu-id="3c472-127">x</span><span class="sxs-lookup"><span data-stu-id="3c472-127">x</span></span>|<span data-ttu-id="3c472-128">x</span><span class="sxs-lookup"><span data-stu-id="3c472-128">x</span></span>|
|[<span data-ttu-id="3c472-129">DefaultSettings</span><span class="sxs-lookup"><span data-stu-id="3c472-129">DefaultSettings</span></span>](defaultsettings.md)|<span data-ttu-id="3c472-130">x</span><span class="sxs-lookup"><span data-stu-id="3c472-130">x</span></span>||<span data-ttu-id="3c472-131">x</span><span class="sxs-lookup"><span data-stu-id="3c472-131">x</span></span>|
|[<span data-ttu-id="3c472-132">DisplayName</span><span class="sxs-lookup"><span data-stu-id="3c472-132">DisplayName</span></span>](displayname.md)|<span data-ttu-id="3c472-133">x</span><span class="sxs-lookup"><span data-stu-id="3c472-133">x</span></span>|<span data-ttu-id="3c472-134">x</span><span class="sxs-lookup"><span data-stu-id="3c472-134">x</span></span>|<span data-ttu-id="3c472-135">x</span><span class="sxs-lookup"><span data-stu-id="3c472-135">x</span></span>|
|[<span data-ttu-id="3c472-136">説明</span><span class="sxs-lookup"><span data-stu-id="3c472-136">Description</span></span>](description.md)|<span data-ttu-id="3c472-137">x</span><span class="sxs-lookup"><span data-stu-id="3c472-137">x</span></span>|<span data-ttu-id="3c472-138">x</span><span class="sxs-lookup"><span data-stu-id="3c472-138">x</span></span>|<span data-ttu-id="3c472-139">x</span><span class="sxs-lookup"><span data-stu-id="3c472-139">x</span></span>|
|[<span data-ttu-id="3c472-140">FormSettings</span><span class="sxs-lookup"><span data-stu-id="3c472-140">FormSettings</span></span>](formsettings.md)||<span data-ttu-id="3c472-141">x</span><span class="sxs-lookup"><span data-stu-id="3c472-141">x</span></span>||
|[<span data-ttu-id="3c472-142">アクセス許可</span><span class="sxs-lookup"><span data-stu-id="3c472-142">Permissions</span></span>](permissions.md)|<span data-ttu-id="3c472-143">x</span><span class="sxs-lookup"><span data-stu-id="3c472-143">x</span></span>||<span data-ttu-id="3c472-144">x</span><span class="sxs-lookup"><span data-stu-id="3c472-144">x</span></span>|
|[<span data-ttu-id="3c472-145">Rule</span><span class="sxs-lookup"><span data-stu-id="3c472-145">Rule</span></span>](rule.md)||<span data-ttu-id="3c472-146">x</span><span class="sxs-lookup"><span data-stu-id="3c472-146">x</span></span>||

## <a name="can-contain"></a><span data-ttu-id="3c472-147">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="3c472-147">Can contain</span></span>

|<span data-ttu-id="3c472-148">**Element**</span><span class="sxs-lookup"><span data-stu-id="3c472-148">**Element**</span></span>|<span data-ttu-id="3c472-149">**コンテンツ**</span><span class="sxs-lookup"><span data-stu-id="3c472-149">**Content**</span></span>|<span data-ttu-id="3c472-150">**メール**</span><span class="sxs-lookup"><span data-stu-id="3c472-150">**Mail**</span></span>|<span data-ttu-id="3c472-151">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="3c472-151">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="3c472-152">AlternateId</span><span class="sxs-lookup"><span data-stu-id="3c472-152">AlternateId</span></span>](alternateid.md)|<span data-ttu-id="3c472-153">x</span><span class="sxs-lookup"><span data-stu-id="3c472-153">x</span></span>|<span data-ttu-id="3c472-154">x</span><span class="sxs-lookup"><span data-stu-id="3c472-154">x</span></span>|<span data-ttu-id="3c472-155">x</span><span class="sxs-lookup"><span data-stu-id="3c472-155">x</span></span>|
|[<span data-ttu-id="3c472-156">IconUrl</span><span class="sxs-lookup"><span data-stu-id="3c472-156">IconUrl</span></span>](iconurl.md)|<span data-ttu-id="3c472-157">x</span><span class="sxs-lookup"><span data-stu-id="3c472-157">x</span></span>|<span data-ttu-id="3c472-158">x</span><span class="sxs-lookup"><span data-stu-id="3c472-158">x</span></span>|<span data-ttu-id="3c472-159">x</span><span class="sxs-lookup"><span data-stu-id="3c472-159">x</span></span>|
|[<span data-ttu-id="3c472-160">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="3c472-160">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|<span data-ttu-id="3c472-161">x</span><span class="sxs-lookup"><span data-stu-id="3c472-161">x</span></span>|<span data-ttu-id="3c472-162">x</span><span class="sxs-lookup"><span data-stu-id="3c472-162">x</span></span>|<span data-ttu-id="3c472-163">x</span><span class="sxs-lookup"><span data-stu-id="3c472-163">x</span></span>|
|[<span data-ttu-id="3c472-164">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="3c472-164">SupportUrl</span></span>](supporturl.md)|<span data-ttu-id="3c472-165">x</span><span class="sxs-lookup"><span data-stu-id="3c472-165">x</span></span>|<span data-ttu-id="3c472-166">x</span><span class="sxs-lookup"><span data-stu-id="3c472-166">x</span></span>|<span data-ttu-id="3c472-167">x</span><span class="sxs-lookup"><span data-stu-id="3c472-167">x</span></span>|
|[<span data-ttu-id="3c472-168">AppDomains</span><span class="sxs-lookup"><span data-stu-id="3c472-168">AppDomains</span></span>](appdomains.md)|<span data-ttu-id="3c472-169">x</span><span class="sxs-lookup"><span data-stu-id="3c472-169">x</span></span>|<span data-ttu-id="3c472-170">x</span><span class="sxs-lookup"><span data-stu-id="3c472-170">x</span></span>|<span data-ttu-id="3c472-171">x</span><span class="sxs-lookup"><span data-stu-id="3c472-171">x</span></span>|
|[<span data-ttu-id="3c472-172">Hosts</span><span class="sxs-lookup"><span data-stu-id="3c472-172">Hosts</span></span>](hosts.md)|<span data-ttu-id="3c472-173">x</span><span class="sxs-lookup"><span data-stu-id="3c472-173">x</span></span>|<span data-ttu-id="3c472-174">x</span><span class="sxs-lookup"><span data-stu-id="3c472-174">x</span></span>|<span data-ttu-id="3c472-175">x</span><span class="sxs-lookup"><span data-stu-id="3c472-175">x</span></span>|
|[<span data-ttu-id="3c472-176">Requirements</span><span class="sxs-lookup"><span data-stu-id="3c472-176">Requirements</span></span>](requirements.md)|<span data-ttu-id="3c472-177">x</span><span class="sxs-lookup"><span data-stu-id="3c472-177">x</span></span>|<span data-ttu-id="3c472-178">x</span><span class="sxs-lookup"><span data-stu-id="3c472-178">x</span></span>|<span data-ttu-id="3c472-179">x</span><span class="sxs-lookup"><span data-stu-id="3c472-179">x</span></span>|
|[<span data-ttu-id="3c472-180">AllowSnapshot</span><span class="sxs-lookup"><span data-stu-id="3c472-180">AllowSnapshot</span></span>](allowsnapshot.md)|<span data-ttu-id="3c472-181">x</span><span class="sxs-lookup"><span data-stu-id="3c472-181">x</span></span>|||
|[<span data-ttu-id="3c472-182">アクセス許可</span><span class="sxs-lookup"><span data-stu-id="3c472-182">Permissions</span></span>](permissions.md)||<span data-ttu-id="3c472-183">x</span><span class="sxs-lookup"><span data-stu-id="3c472-183">x</span></span>||
|[<span data-ttu-id="3c472-184">DisableEntityHighlighting</span><span class="sxs-lookup"><span data-stu-id="3c472-184">DisableEntityHighlighting</span></span>](disableentityhighlighting.md)||<span data-ttu-id="3c472-185">x</span><span class="sxs-lookup"><span data-stu-id="3c472-185">x</span></span>||
|[<span data-ttu-id="3c472-186">Dictionary</span><span class="sxs-lookup"><span data-stu-id="3c472-186">Dictionary</span></span>](dictionary.md)|||<span data-ttu-id="3c472-187">x</span><span class="sxs-lookup"><span data-stu-id="3c472-187">x</span></span>|
|[<span data-ttu-id="3c472-188">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="3c472-188">VersionOverrides</span></span>](versionoverrides.md)|<span data-ttu-id="3c472-189">x</span><span class="sxs-lookup"><span data-stu-id="3c472-189">x</span></span>|<span data-ttu-id="3c472-190">x</span><span class="sxs-lookup"><span data-stu-id="3c472-190">x</span></span>|<span data-ttu-id="3c472-191">x</span><span class="sxs-lookup"><span data-stu-id="3c472-191">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="3c472-192">属性</span><span class="sxs-lookup"><span data-stu-id="3c472-192">Attributes</span></span>

|||
|:-----|:-----|
|<span data-ttu-id="3c472-193">xmlns</span><span class="sxs-lookup"><span data-stu-id="3c472-193">xmlns</span></span>|<span data-ttu-id="3c472-p101">Office アドイン マニフェストの名前空間とスキーマ バージョンを定義します。この属性は常に `"http://schemas.microsoft.com/office/appforoffice/1.1"` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="3c472-p101">Defines the Office Add-in manifest namespace and schema version. This attribute should always be set to  `"http://schemas.microsoft.com/office/appforoffice/1.1"`</span></span>|
|<span data-ttu-id="3c472-196">xmlns:xsi</span><span class="sxs-lookup"><span data-stu-id="3c472-196">xmlns:xsi</span></span>|<span data-ttu-id="3c472-p102">XMLSchema インスタンスを定義します。この属性は常に `"http://www.w3.org/2001/XMLSchema-instance"` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="3c472-p102">Defines the XMLSchema instance. This attribute should always be set to  `"http://www.w3.org/2001/XMLSchema-instance"`</span></span>|
|<span data-ttu-id="3c472-199">xsi:type</span><span class="sxs-lookup"><span data-stu-id="3c472-199">xsi:type</span></span>|<span data-ttu-id="3c472-p103">Office アドインの種類を定義します。この属性は、`"ContentApp"`、`"MailApp"`、または `"TaskPaneApp"` のいずれかに設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="3c472-p103">Defines the kind of Office Add-in. This attribute should be set to one of:  `"ContentApp"`,  `"MailApp"`, or  `"TaskPaneApp"`</span></span>|
