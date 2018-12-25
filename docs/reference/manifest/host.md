---
title: マニフェスト ファイルの Host 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 37b772261ad82b4f899e73314a08ffd1dd03b442
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432810"
---
# <a name="host-element"></a><span data-ttu-id="0916a-102">Host 要素</span><span class="sxs-lookup"><span data-stu-id="0916a-102">Host element</span></span>

<span data-ttu-id="0916a-103">アドインでアクティブ化する Office アプリケーションの種類を個別に指定します。</span><span class="sxs-lookup"><span data-stu-id="0916a-103">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="0916a-104">**Host** 要素の構文は、要素が[基本のマニフェスト](#basic-manifest)で定義されているか、[VersionOverrides](#versionoverrides-node) ノードで定義されているかによって異なります。</span><span class="sxs-lookup"><span data-stu-id="0916a-104">Important: The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="0916a-105">ただし、機能は変わりません。</span><span class="sxs-lookup"><span data-stu-id="0916a-105">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="0916a-106">基本のマニフェスト</span><span class="sxs-lookup"><span data-stu-id="0916a-106">Basic manifest</span></span>

<span data-ttu-id="0916a-107">基本のマニフェストで定義されている場合 ([OfficeApp](officeapp.md) の下)、ホストの種類は `Name` 属性によって決定されます。</span><span class="sxs-lookup"><span data-stu-id="0916a-107">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>   

### <a name="attributes"></a><span data-ttu-id="0916a-108">属性</span><span class="sxs-lookup"><span data-stu-id="0916a-108">Attributes</span></span>

| <span data-ttu-id="0916a-109">属性</span><span class="sxs-lookup"><span data-stu-id="0916a-109">Attribute</span></span>     | <span data-ttu-id="0916a-110">種類</span><span class="sxs-lookup"><span data-stu-id="0916a-110">Type</span></span>   | <span data-ttu-id="0916a-111">必須</span><span class="sxs-lookup"><span data-stu-id="0916a-111">Required</span></span> | <span data-ttu-id="0916a-112">説明</span><span class="sxs-lookup"><span data-stu-id="0916a-112">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="0916a-113">Name</span><span class="sxs-lookup"><span data-stu-id="0916a-113">Name</span></span>](#name) | <span data-ttu-id="0916a-114">string</span><span class="sxs-lookup"><span data-stu-id="0916a-114">string</span></span> | <span data-ttu-id="0916a-115">必須</span><span class="sxs-lookup"><span data-stu-id="0916a-115">required</span></span> | <span data-ttu-id="0916a-116">Office ホスト アプリケーションの種類の名前。</span><span class="sxs-lookup"><span data-stu-id="0916a-116">The name of the type of Office host application.</span></span> |

### <a name="name"></a><span data-ttu-id="0916a-117">名前</span><span class="sxs-lookup"><span data-stu-id="0916a-117">Name</span></span>
<span data-ttu-id="0916a-p102">このアドインが対象にするホストの種類を指定します。この値は、次のいずれかである必要があります。</span><span class="sxs-lookup"><span data-stu-id="0916a-p102">Specifies the Host type targeted by this add-in. The value must be one of the following:</span></span>

- <span data-ttu-id="0916a-120">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="0916a-120">`Document` (Word)</span></span>
- <span data-ttu-id="0916a-121">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="0916a-121">`Database` (Access)</span></span>
- <span data-ttu-id="0916a-122">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="0916a-122">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="0916a-123">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="0916a-123">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="0916a-124">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="0916a-124">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="0916a-125">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="0916a-125">`Project` (Project)</span></span>
- <span data-ttu-id="0916a-126">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="0916a-126">`Workbook` (Excel)</span></span>

### <a name="example"></a><span data-ttu-id="0916a-127">例</span><span class="sxs-lookup"><span data-stu-id="0916a-127">Example</span></span>
```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="0916a-128">VersionOverrides ノード</span><span class="sxs-lookup"><span data-stu-id="0916a-128">VersionOverrides node</span></span>
<span data-ttu-id="0916a-129">[VersionOverrides](versionoverrides.md) で定義されている場合、ホストの種類は `xsi:type` 属性によって決定されます。</span><span class="sxs-lookup"><span data-stu-id="0916a-129">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span> 

### <a name="attributes"></a><span data-ttu-id="0916a-130">属性</span><span class="sxs-lookup"><span data-stu-id="0916a-130">Attributes</span></span>

|  <span data-ttu-id="0916a-131">属性</span><span class="sxs-lookup"><span data-stu-id="0916a-131">Attribute</span></span>  |  <span data-ttu-id="0916a-132">必須</span><span class="sxs-lookup"><span data-stu-id="0916a-132">Required</span></span>  |  <span data-ttu-id="0916a-133">説明</span><span class="sxs-lookup"><span data-stu-id="0916a-133">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="0916a-134">xsi:type</span><span class="sxs-lookup"><span data-stu-id="0916a-134">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="0916a-135">はい</span><span class="sxs-lookup"><span data-stu-id="0916a-135">Yes</span></span>  | <span data-ttu-id="0916a-136">これらの設定を適用する Office ホストについて説明します。</span><span class="sxs-lookup"><span data-stu-id="0916a-136">Describes the Office host where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="0916a-137">子要素</span><span class="sxs-lookup"><span data-stu-id="0916a-137">Child elements</span></span>

|  <span data-ttu-id="0916a-138">要素</span><span class="sxs-lookup"><span data-stu-id="0916a-138">Element</span></span> |  <span data-ttu-id="0916a-139">必須</span><span class="sxs-lookup"><span data-stu-id="0916a-139">Required</span></span>  |  <span data-ttu-id="0916a-140">説明</span><span class="sxs-lookup"><span data-stu-id="0916a-140">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="0916a-141">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="0916a-141">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="0916a-142">はい</span><span class="sxs-lookup"><span data-stu-id="0916a-142">Yes</span></span>   |  <span data-ttu-id="0916a-143">デスクトップ フォーム ファクターの設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="0916a-143">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="0916a-144">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="0916a-144">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="0916a-145">いいえ</span><span class="sxs-lookup"><span data-stu-id="0916a-145">No</span></span>   |  <span data-ttu-id="0916a-p103">モバイル フォーム ファクターの設定を定義します。**注:** この要素は、Outlook for iOS でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="0916a-p103">Defines the settings for the mobile form factor. **Note:** this element is only supported in Outlook for iOS.</span></span> |
|  [<span data-ttu-id="0916a-148">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="0916a-148">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="0916a-149">いいえ</span><span class="sxs-lookup"><span data-stu-id="0916a-149">No</span></span>   |  <span data-ttu-id="0916a-150">すべてのフォーム ファクターの設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="0916a-150">Defines the settings for all form factors.</span></span> <span data-ttu-id="0916a-151">Excel のカスタム関数でのみ使用します。</span><span class="sxs-lookup"><span data-stu-id="0916a-151">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="0916a-152">xsi:type</span><span class="sxs-lookup"><span data-stu-id="0916a-152">xsi:type</span></span>

<span data-ttu-id="0916a-153">含まれている設定を適用する Office ホスト (Word、Excel、PowerPoint、Outlook、OneNote) を制御します。</span><span class="sxs-lookup"><span data-stu-id="0916a-153">Controls which Office host (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="0916a-154">この値は、次のいずれかである必要があります。</span><span class="sxs-lookup"><span data-stu-id="0916a-154">The value must be one of the following:</span></span>

- <span data-ttu-id="0916a-155">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="0916a-155">`Document` (Word)</span></span>
- <span data-ttu-id="0916a-156">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="0916a-156">`MailHost` (Outlook)</span></span>    
- <span data-ttu-id="0916a-157">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="0916a-157">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="0916a-158">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="0916a-158">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="0916a-159">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="0916a-159">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="0916a-160">ホストの例</span><span class="sxs-lookup"><span data-stu-id="0916a-160">Host example</span></span> 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
