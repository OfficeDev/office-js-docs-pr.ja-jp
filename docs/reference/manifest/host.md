---
title: マニフェスト ファイルの Host 要素
description: アドインでアクティブ化する Office アプリケーションの種類を個別に指定します。
ms.date: 11/05/2019
localization_priority: Normal
ms.openlocfilehash: b9f03e6d6b028ca6f4616ae81b8fd76601256793
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718133"
---
# <a name="host-element"></a><span data-ttu-id="6a225-103">Host 要素</span><span class="sxs-lookup"><span data-stu-id="6a225-103">Host element</span></span>

<span data-ttu-id="6a225-104">アドインでアクティブ化する Office アプリケーションの種類を個別に指定します。</span><span class="sxs-lookup"><span data-stu-id="6a225-104">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="6a225-105">**Host** 要素の構文は、要素が[基本のマニフェスト](#basic-manifest)で定義されているか、[VersionOverrides](#versionoverrides-node) ノードで定義されているかによって異なります。</span><span class="sxs-lookup"><span data-stu-id="6a225-105">The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="6a225-106">ただし、機能は変わりません。</span><span class="sxs-lookup"><span data-stu-id="6a225-106">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="6a225-107">基本のマニフェスト</span><span class="sxs-lookup"><span data-stu-id="6a225-107">Basic manifest</span></span>

<span data-ttu-id="6a225-108">基本のマニフェストで定義されている場合 ([OfficeApp](officeapp.md) の下)、ホストの種類は `Name` 属性によって決定されます。</span><span class="sxs-lookup"><span data-stu-id="6a225-108">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="6a225-109">属性</span><span class="sxs-lookup"><span data-stu-id="6a225-109">Attributes</span></span>

| <span data-ttu-id="6a225-110">属性</span><span class="sxs-lookup"><span data-stu-id="6a225-110">Attribute</span></span>     | <span data-ttu-id="6a225-111">型</span><span class="sxs-lookup"><span data-stu-id="6a225-111">Type</span></span>   | <span data-ttu-id="6a225-112">必須</span><span class="sxs-lookup"><span data-stu-id="6a225-112">Required</span></span> | <span data-ttu-id="6a225-113">説明</span><span class="sxs-lookup"><span data-stu-id="6a225-113">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="6a225-114">名前</span><span class="sxs-lookup"><span data-stu-id="6a225-114">Name</span></span>](#name) | <span data-ttu-id="6a225-115">string</span><span class="sxs-lookup"><span data-stu-id="6a225-115">string</span></span> | <span data-ttu-id="6a225-116">必須</span><span class="sxs-lookup"><span data-stu-id="6a225-116">required</span></span> | <span data-ttu-id="6a225-117">Office ホスト アプリケーションの種類の名前。</span><span class="sxs-lookup"><span data-stu-id="6a225-117">The name of the type of Office host application.</span></span> |

### <a name="name"></a><span data-ttu-id="6a225-118">名前</span><span class="sxs-lookup"><span data-stu-id="6a225-118">Name</span></span>

<span data-ttu-id="6a225-119">このアドインが対象にするホストの種類を指定します。</span><span class="sxs-lookup"><span data-stu-id="6a225-119">Specifies the Host type targeted by this add-in.</span></span> <span data-ttu-id="6a225-120">この値は、次のいずれかであることが必要です。</span><span class="sxs-lookup"><span data-stu-id="6a225-120">The value must be one of the following.</span></span>

- <span data-ttu-id="6a225-121">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="6a225-121">`Document` (Word)</span></span>
- <span data-ttu-id="6a225-122">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="6a225-122">`Database` (Access)</span></span>
- <span data-ttu-id="6a225-123">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="6a225-123">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="6a225-124">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="6a225-124">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="6a225-125">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="6a225-125">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="6a225-126">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="6a225-126">`Project` (Project)</span></span>
- <span data-ttu-id="6a225-127">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="6a225-127">`Workbook` (Excel)</span></span>

> [!IMPORTANT]
> <span data-ttu-id="6a225-128">SharePoint で Access Web アプリとデータベースを作成して使用することは推奨されなくなりました。</span><span class="sxs-lookup"><span data-stu-id="6a225-128">We no longer recommend that you create and use Access web apps and databases in SharePoint.</span></span> <span data-ttu-id="6a225-129">代わりに、[Microsoft PowerApps](https://powerapps.microsoft.com/) を使用して、コード作成が不要な Web とモバイル デバイス用ビジネス ソリューションをビルドすることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="6a225-129">As an alternative, we recommend that you use [Microsoft PowerApps](https://powerapps.microsoft.com/) to build no-code business solutions for web and mobile devices.</span></span>

### <a name="example"></a><span data-ttu-id="6a225-130">例</span><span class="sxs-lookup"><span data-stu-id="6a225-130">Example</span></span>

```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="6a225-131">VersionOverrides ノード</span><span class="sxs-lookup"><span data-stu-id="6a225-131">VersionOverrides node</span></span>

<span data-ttu-id="6a225-132">[VersionOverrides](versionoverrides.md) で定義されている場合、ホストの種類は `xsi:type` 属性によって決定されます。</span><span class="sxs-lookup"><span data-stu-id="6a225-132">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="6a225-133">属性</span><span class="sxs-lookup"><span data-stu-id="6a225-133">Attributes</span></span>

|  <span data-ttu-id="6a225-134">属性</span><span class="sxs-lookup"><span data-stu-id="6a225-134">Attribute</span></span>  |  <span data-ttu-id="6a225-135">必須</span><span class="sxs-lookup"><span data-stu-id="6a225-135">Required</span></span>  |  <span data-ttu-id="6a225-136">説明</span><span class="sxs-lookup"><span data-stu-id="6a225-136">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="6a225-137">xsi:type</span><span class="sxs-lookup"><span data-stu-id="6a225-137">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="6a225-138">はい</span><span class="sxs-lookup"><span data-stu-id="6a225-138">Yes</span></span>  | <span data-ttu-id="6a225-139">これらの設定を適用する Office ホストについて説明します。</span><span class="sxs-lookup"><span data-stu-id="6a225-139">Describes the Office host where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="6a225-140">子要素</span><span class="sxs-lookup"><span data-stu-id="6a225-140">Child elements</span></span>

|  <span data-ttu-id="6a225-141">要素</span><span class="sxs-lookup"><span data-stu-id="6a225-141">Element</span></span> |  <span data-ttu-id="6a225-142">必須</span><span class="sxs-lookup"><span data-stu-id="6a225-142">Required</span></span>  |  <span data-ttu-id="6a225-143">説明</span><span class="sxs-lookup"><span data-stu-id="6a225-143">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="6a225-144">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="6a225-144">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="6a225-145">はい</span><span class="sxs-lookup"><span data-stu-id="6a225-145">Yes</span></span>   |  <span data-ttu-id="6a225-146">デスクトップ フォーム ファクターの設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="6a225-146">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="6a225-147">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="6a225-147">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="6a225-148">いいえ</span><span class="sxs-lookup"><span data-stu-id="6a225-148">No</span></span>   |  <span data-ttu-id="6a225-149">モバイルフォームファクターの設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="6a225-149">Defines the settings for the mobile form factor.</span></span> <span data-ttu-id="6a225-150">**注:** この要素は、iOS および Android の Outlook でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="6a225-150">**Note:** This element is only supported in Outlook on iOS and Android.</span></span> |
|  [<span data-ttu-id="6a225-151">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="6a225-151">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="6a225-152">いいえ</span><span class="sxs-lookup"><span data-stu-id="6a225-152">No</span></span>   |  <span data-ttu-id="6a225-153">すべてのフォーム ファクターの設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="6a225-153">Defines the settings for all form factors.</span></span> <span data-ttu-id="6a225-154">Excel のカスタム関数でのみ使用します。</span><span class="sxs-lookup"><span data-stu-id="6a225-154">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="6a225-155">xsi:type</span><span class="sxs-lookup"><span data-stu-id="6a225-155">xsi:type</span></span>

<span data-ttu-id="6a225-156">含まれている設定を適用する Office ホスト (Word、Excel、PowerPoint、Outlook、OneNote) を制御します。</span><span class="sxs-lookup"><span data-stu-id="6a225-156">Controls which Office host (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="6a225-157">この値は、次のいずれかである必要があります。</span><span class="sxs-lookup"><span data-stu-id="6a225-157">The value must be one of the following:</span></span>

- <span data-ttu-id="6a225-158">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="6a225-158">`Document` (Word)</span></span>
- <span data-ttu-id="6a225-159">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="6a225-159">`MailHost` (Outlook)</span></span>
- <span data-ttu-id="6a225-160">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="6a225-160">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="6a225-161">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="6a225-161">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="6a225-162">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="6a225-162">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="6a225-163">ホストの例</span><span class="sxs-lookup"><span data-stu-id="6a225-163">Host example</span></span>

```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
