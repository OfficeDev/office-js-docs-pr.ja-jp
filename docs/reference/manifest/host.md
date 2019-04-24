---
title: マニフェスト ファイルの Host 要素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: f496e3e0c16f24d20e1d1db76208e61267235131
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450507"
---
# <a name="host-element"></a><span data-ttu-id="12b14-102">Host 要素</span><span class="sxs-lookup"><span data-stu-id="12b14-102">Host element</span></span>

<span data-ttu-id="12b14-103">アドインでアクティブ化する Office アプリケーションの種類を個別に指定します。</span><span class="sxs-lookup"><span data-stu-id="12b14-103">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="12b14-104">**Host** 要素の構文は、要素が[基本のマニフェスト](#basic-manifest)で定義されているか、[VersionOverrides](#versionoverrides-node) ノードで定義されているかによって異なります。</span><span class="sxs-lookup"><span data-stu-id="12b14-104">The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="12b14-105">ただし、機能は変わりません。</span><span class="sxs-lookup"><span data-stu-id="12b14-105">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="12b14-106">基本のマニフェスト</span><span class="sxs-lookup"><span data-stu-id="12b14-106">Basic manifest</span></span>

<span data-ttu-id="12b14-107">基本のマニフェストで定義されている場合 ([OfficeApp](officeapp.md) の下)、ホストの種類は `Name` 属性によって決定されます。</span><span class="sxs-lookup"><span data-stu-id="12b14-107">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>   

### <a name="attributes"></a><span data-ttu-id="12b14-108">属性</span><span class="sxs-lookup"><span data-stu-id="12b14-108">Attributes</span></span>

| <span data-ttu-id="12b14-109">属性</span><span class="sxs-lookup"><span data-stu-id="12b14-109">Attribute</span></span>     | <span data-ttu-id="12b14-110">Type</span><span class="sxs-lookup"><span data-stu-id="12b14-110">Type</span></span>   | <span data-ttu-id="12b14-111">必須</span><span class="sxs-lookup"><span data-stu-id="12b14-111">Required</span></span> | <span data-ttu-id="12b14-112">説明</span><span class="sxs-lookup"><span data-stu-id="12b14-112">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="12b14-113">名前</span><span class="sxs-lookup"><span data-stu-id="12b14-113">Name</span></span>](#name) | <span data-ttu-id="12b14-114">string</span><span class="sxs-lookup"><span data-stu-id="12b14-114">string</span></span> | <span data-ttu-id="12b14-115">必須</span><span class="sxs-lookup"><span data-stu-id="12b14-115">required</span></span> | <span data-ttu-id="12b14-116">Office ホスト アプリケーションの種類の名前。</span><span class="sxs-lookup"><span data-stu-id="12b14-116">The name of the type of Office host application.</span></span> |

### <a name="name"></a><span data-ttu-id="12b14-117">名前</span><span class="sxs-lookup"><span data-stu-id="12b14-117">Name</span></span>
<span data-ttu-id="12b14-p102">このアドインが対象にするホストの種類を指定します。この値は、次のいずれかである必要があります。</span><span class="sxs-lookup"><span data-stu-id="12b14-p102">Specifies the Host type targeted by this add-in. The value must be one of the following:</span></span>

- <span data-ttu-id="12b14-120">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="12b14-120">`Document` (Word)</span></span>
- <span data-ttu-id="12b14-121">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="12b14-121">`Database` (Access)</span></span>
- <span data-ttu-id="12b14-122">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="12b14-122">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="12b14-123">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="12b14-123">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="12b14-124">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="12b14-124">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="12b14-125">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="12b14-125">`Project` (Project)</span></span>
- <span data-ttu-id="12b14-126">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="12b14-126">`Workbook` (Excel)</span></span>

### <a name="example"></a><span data-ttu-id="12b14-127">例</span><span class="sxs-lookup"><span data-stu-id="12b14-127">Example</span></span>
```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="12b14-128">VersionOverrides ノード</span><span class="sxs-lookup"><span data-stu-id="12b14-128">VersionOverrides node</span></span>
<span data-ttu-id="12b14-129">[VersionOverrides](versionoverrides.md) で定義されている場合、ホストの種類は `xsi:type` 属性によって決定されます。</span><span class="sxs-lookup"><span data-stu-id="12b14-129">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span> 

### <a name="attributes"></a><span data-ttu-id="12b14-130">属性</span><span class="sxs-lookup"><span data-stu-id="12b14-130">Attributes</span></span>

|  <span data-ttu-id="12b14-131">属性</span><span class="sxs-lookup"><span data-stu-id="12b14-131">Attribute</span></span>  |  <span data-ttu-id="12b14-132">必須</span><span class="sxs-lookup"><span data-stu-id="12b14-132">Required</span></span>  |  <span data-ttu-id="12b14-133">説明</span><span class="sxs-lookup"><span data-stu-id="12b14-133">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="12b14-134">xsi:type</span><span class="sxs-lookup"><span data-stu-id="12b14-134">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="12b14-135">はい</span><span class="sxs-lookup"><span data-stu-id="12b14-135">Yes</span></span>  | <span data-ttu-id="12b14-136">これらの設定を適用する Office ホストについて説明します。</span><span class="sxs-lookup"><span data-stu-id="12b14-136">Describes the Office host where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="12b14-137">子要素</span><span class="sxs-lookup"><span data-stu-id="12b14-137">Child elements</span></span>

|  <span data-ttu-id="12b14-138">要素</span><span class="sxs-lookup"><span data-stu-id="12b14-138">Element</span></span> |  <span data-ttu-id="12b14-139">必須</span><span class="sxs-lookup"><span data-stu-id="12b14-139">Required</span></span>  |  <span data-ttu-id="12b14-140">説明</span><span class="sxs-lookup"><span data-stu-id="12b14-140">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="12b14-141">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="12b14-141">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="12b14-142">はい</span><span class="sxs-lookup"><span data-stu-id="12b14-142">Yes</span></span>   |  <span data-ttu-id="12b14-143">デスクトップ フォーム ファクターの設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="12b14-143">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="12b14-144">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="12b14-144">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="12b14-145">いいえ</span><span class="sxs-lookup"><span data-stu-id="12b14-145">No</span></span>   |  <span data-ttu-id="12b14-p103">モバイル フォーム ファクターの設定を定義します。**注:** この要素は、Outlook for iOS でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="12b14-p103">Defines the settings for the mobile form factor. **Note:** this element is only supported in Outlook for iOS.</span></span> |
|  [<span data-ttu-id="12b14-148">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="12b14-148">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="12b14-149">いいえ</span><span class="sxs-lookup"><span data-stu-id="12b14-149">No</span></span>   |  <span data-ttu-id="12b14-150">すべてのフォーム ファクターの設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="12b14-150">Defines the settings for all form factors.</span></span> <span data-ttu-id="12b14-151">Excel のカスタム関数でのみ使用します。</span><span class="sxs-lookup"><span data-stu-id="12b14-151">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="12b14-152">xsi:type</span><span class="sxs-lookup"><span data-stu-id="12b14-152">xsi:type</span></span>

<span data-ttu-id="12b14-153">含まれている設定を適用する Office ホスト (Word、Excel、PowerPoint、Outlook、OneNote) を制御します。</span><span class="sxs-lookup"><span data-stu-id="12b14-153">Controls which Office host (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="12b14-154">この値は、次のいずれかである必要があります。</span><span class="sxs-lookup"><span data-stu-id="12b14-154">The value must be one of the following:</span></span>

- <span data-ttu-id="12b14-155">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="12b14-155">`Document` (Word)</span></span>
- <span data-ttu-id="12b14-156">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="12b14-156">`MailHost` (Outlook)</span></span>    
- <span data-ttu-id="12b14-157">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="12b14-157">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="12b14-158">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="12b14-158">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="12b14-159">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="12b14-159">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="12b14-160">ホストの例</span><span class="sxs-lookup"><span data-stu-id="12b14-160">Host example</span></span> 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
