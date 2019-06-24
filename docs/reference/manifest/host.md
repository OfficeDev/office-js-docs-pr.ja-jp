---
title: マニフェスト ファイルの Host 要素
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: debb4d59f75ce974ffb21d853c6b65a579c4e685
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127570"
---
# <a name="host-element"></a><span data-ttu-id="ebd33-102">Host 要素</span><span class="sxs-lookup"><span data-stu-id="ebd33-102">Host element</span></span>

<span data-ttu-id="ebd33-103">アドインでアクティブ化する Office アプリケーションの種類を個別に指定します。</span><span class="sxs-lookup"><span data-stu-id="ebd33-103">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="ebd33-104">**Host** 要素の構文は、要素が[基本のマニフェスト](#basic-manifest)で定義されているか、[VersionOverrides](#versionoverrides-node) ノードで定義されているかによって異なります。</span><span class="sxs-lookup"><span data-stu-id="ebd33-104">The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="ebd33-105">ただし、機能は変わりません。</span><span class="sxs-lookup"><span data-stu-id="ebd33-105">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="ebd33-106">基本のマニフェスト</span><span class="sxs-lookup"><span data-stu-id="ebd33-106">Basic manifest</span></span>

<span data-ttu-id="ebd33-107">基本のマニフェストで定義されている場合 ([OfficeApp](officeapp.md) の下)、ホストの種類は `Name` 属性によって決定されます。</span><span class="sxs-lookup"><span data-stu-id="ebd33-107">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="ebd33-108">属性</span><span class="sxs-lookup"><span data-stu-id="ebd33-108">Attributes</span></span>

| <span data-ttu-id="ebd33-109">属性</span><span class="sxs-lookup"><span data-stu-id="ebd33-109">Attribute</span></span>     | <span data-ttu-id="ebd33-110">型</span><span class="sxs-lookup"><span data-stu-id="ebd33-110">Type</span></span>   | <span data-ttu-id="ebd33-111">必須</span><span class="sxs-lookup"><span data-stu-id="ebd33-111">Required</span></span> | <span data-ttu-id="ebd33-112">説明</span><span class="sxs-lookup"><span data-stu-id="ebd33-112">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="ebd33-113">名前</span><span class="sxs-lookup"><span data-stu-id="ebd33-113">Name</span></span>](#name) | <span data-ttu-id="ebd33-114">string</span><span class="sxs-lookup"><span data-stu-id="ebd33-114">string</span></span> | <span data-ttu-id="ebd33-115">必須</span><span class="sxs-lookup"><span data-stu-id="ebd33-115">required</span></span> | <span data-ttu-id="ebd33-116">Office ホスト アプリケーションの種類の名前。</span><span class="sxs-lookup"><span data-stu-id="ebd33-116">The name of the type of Office host application.</span></span> |

### <a name="name"></a><span data-ttu-id="ebd33-117">名前</span><span class="sxs-lookup"><span data-stu-id="ebd33-117">Name</span></span>
<span data-ttu-id="ebd33-p102">このアドインが対象にするホストの種類を指定します。この値は、次のいずれかである必要があります。</span><span class="sxs-lookup"><span data-stu-id="ebd33-p102">Specifies the Host type targeted by this add-in. The value must be one of the following:</span></span>

- <span data-ttu-id="ebd33-120">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="ebd33-120">`Document` (Word)</span></span>
- <span data-ttu-id="ebd33-121">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="ebd33-121">`Database` (Access)</span></span>
- <span data-ttu-id="ebd33-122">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="ebd33-122">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="ebd33-123">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="ebd33-123">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="ebd33-124">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="ebd33-124">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="ebd33-125">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="ebd33-125">`Project` (Project)</span></span>
- <span data-ttu-id="ebd33-126">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="ebd33-126">`Workbook` (Excel)</span></span>

### <a name="example"></a><span data-ttu-id="ebd33-127">例</span><span class="sxs-lookup"><span data-stu-id="ebd33-127">Example</span></span>
```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="ebd33-128">VersionOverrides ノード</span><span class="sxs-lookup"><span data-stu-id="ebd33-128">VersionOverrides node</span></span>
<span data-ttu-id="ebd33-129">[VersionOverrides](versionoverrides.md) で定義されている場合、ホストの種類は `xsi:type` 属性によって決定されます。</span><span class="sxs-lookup"><span data-stu-id="ebd33-129">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span> 

### <a name="attributes"></a><span data-ttu-id="ebd33-130">属性</span><span class="sxs-lookup"><span data-stu-id="ebd33-130">Attributes</span></span>

|  <span data-ttu-id="ebd33-131">属性</span><span class="sxs-lookup"><span data-stu-id="ebd33-131">Attribute</span></span>  |  <span data-ttu-id="ebd33-132">必須</span><span class="sxs-lookup"><span data-stu-id="ebd33-132">Required</span></span>  |  <span data-ttu-id="ebd33-133">説明</span><span class="sxs-lookup"><span data-stu-id="ebd33-133">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="ebd33-134">xsi:type</span><span class="sxs-lookup"><span data-stu-id="ebd33-134">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="ebd33-135">はい</span><span class="sxs-lookup"><span data-stu-id="ebd33-135">Yes</span></span>  | <span data-ttu-id="ebd33-136">これらの設定を適用する Office ホストについて説明します。</span><span class="sxs-lookup"><span data-stu-id="ebd33-136">Describes the Office host where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="ebd33-137">子要素</span><span class="sxs-lookup"><span data-stu-id="ebd33-137">Child elements</span></span>

|  <span data-ttu-id="ebd33-138">要素</span><span class="sxs-lookup"><span data-stu-id="ebd33-138">Element</span></span> |  <span data-ttu-id="ebd33-139">必須</span><span class="sxs-lookup"><span data-stu-id="ebd33-139">Required</span></span>  |  <span data-ttu-id="ebd33-140">説明</span><span class="sxs-lookup"><span data-stu-id="ebd33-140">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="ebd33-141">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="ebd33-141">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="ebd33-142">はい</span><span class="sxs-lookup"><span data-stu-id="ebd33-142">Yes</span></span>   |  <span data-ttu-id="ebd33-143">デスクトップ フォーム ファクターの設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="ebd33-143">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="ebd33-144">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="ebd33-144">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="ebd33-145">いいえ</span><span class="sxs-lookup"><span data-stu-id="ebd33-145">No</span></span>   |  <span data-ttu-id="ebd33-146">モバイルフォームファクターの設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="ebd33-146">Defines the settings for the mobile form factor.</span></span> <span data-ttu-id="ebd33-147">**注:** この要素は、iOS の Outlook でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="ebd33-147">**Note:** This element is only supported in Outlook on iOS.</span></span> |
|  [<span data-ttu-id="ebd33-148">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="ebd33-148">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="ebd33-149">いいえ</span><span class="sxs-lookup"><span data-stu-id="ebd33-149">No</span></span>   |  <span data-ttu-id="ebd33-150">すべてのフォーム ファクターの設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="ebd33-150">Defines the settings for all form factors.</span></span> <span data-ttu-id="ebd33-151">Excel のカスタム関数でのみ使用します。</span><span class="sxs-lookup"><span data-stu-id="ebd33-151">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="ebd33-152">xsi:type</span><span class="sxs-lookup"><span data-stu-id="ebd33-152">xsi:type</span></span>

<span data-ttu-id="ebd33-153">含まれている設定を適用する Office ホスト (Word、Excel、PowerPoint、Outlook、OneNote) を制御します。</span><span class="sxs-lookup"><span data-stu-id="ebd33-153">Controls which Office host (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="ebd33-154">この値は、次のいずれかである必要があります。</span><span class="sxs-lookup"><span data-stu-id="ebd33-154">The value must be one of the following:</span></span>

- <span data-ttu-id="ebd33-155">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="ebd33-155">`Document` (Word)</span></span>
- <span data-ttu-id="ebd33-156">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="ebd33-156">`MailHost` (Outlook)</span></span>
- <span data-ttu-id="ebd33-157">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="ebd33-157">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="ebd33-158">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="ebd33-158">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="ebd33-159">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="ebd33-159">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="ebd33-160">ホストの例</span><span class="sxs-lookup"><span data-stu-id="ebd33-160">Host example</span></span> 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
