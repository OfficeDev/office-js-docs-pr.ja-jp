---
title: マニフェスト ファイルの Host 要素
description: アドインでアクティブ化する Office アプリケーションの種類を個別に指定します。
ms.date: 11/05/2019
localization_priority: Normal
ms.openlocfilehash: 45d4ed42946038699be235ff3912c071a92ff226
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348329"
---
# <a name="host-element"></a><span data-ttu-id="baf98-103">Host 要素</span><span class="sxs-lookup"><span data-stu-id="baf98-103">Host element</span></span>

<span data-ttu-id="baf98-104">アドインでアクティブ化する Office アプリケーションの種類を個別に指定します。</span><span class="sxs-lookup"><span data-stu-id="baf98-104">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="baf98-105">**Host** 要素の構文は、要素が [基本のマニフェスト](#basic-manifest)で定義されているか、[VersionOverrides](#versionoverrides-node) ノードで定義されているかによって異なります。</span><span class="sxs-lookup"><span data-stu-id="baf98-105">The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="baf98-106">ただし、機能は変わりません。</span><span class="sxs-lookup"><span data-stu-id="baf98-106">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="baf98-107">基本のマニフェスト</span><span class="sxs-lookup"><span data-stu-id="baf98-107">Basic manifest</span></span>

<span data-ttu-id="baf98-108">基本のマニフェストで定義されている場合 ([OfficeApp](officeapp.md) の下)、ホストの種類は `Name` 属性によって決定されます。</span><span class="sxs-lookup"><span data-stu-id="baf98-108">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="baf98-109">属性</span><span class="sxs-lookup"><span data-stu-id="baf98-109">Attributes</span></span>

| <span data-ttu-id="baf98-110">属性</span><span class="sxs-lookup"><span data-stu-id="baf98-110">Attribute</span></span>     | <span data-ttu-id="baf98-111">型</span><span class="sxs-lookup"><span data-stu-id="baf98-111">Type</span></span>   | <span data-ttu-id="baf98-112">必須</span><span class="sxs-lookup"><span data-stu-id="baf98-112">Required</span></span> | <span data-ttu-id="baf98-113">説明</span><span class="sxs-lookup"><span data-stu-id="baf98-113">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="baf98-114">名前</span><span class="sxs-lookup"><span data-stu-id="baf98-114">Name</span></span>](#name) | <span data-ttu-id="baf98-115">string</span><span class="sxs-lookup"><span data-stu-id="baf98-115">string</span></span> | <span data-ttu-id="baf98-116">必須</span><span class="sxs-lookup"><span data-stu-id="baf98-116">required</span></span> | <span data-ttu-id="baf98-117">クライアント アプリケーションの種類Officeします。</span><span class="sxs-lookup"><span data-stu-id="baf98-117">The name of the type of Office client application.</span></span> |

### <a name="name"></a><span data-ttu-id="baf98-118">名前</span><span class="sxs-lookup"><span data-stu-id="baf98-118">Name</span></span>

<span data-ttu-id="baf98-p102">このアドインが対象にするホストの種類を指定します。この値は、次のいずれかである必要があります。</span><span class="sxs-lookup"><span data-stu-id="baf98-p102">Specifies the Host type targeted by this add-in. The value must be one of the following:</span></span>

- <span data-ttu-id="baf98-121">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="baf98-121">`Document` (Word)</span></span>
- <span data-ttu-id="baf98-122">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="baf98-122">`Database` (Access)</span></span>
- <span data-ttu-id="baf98-123">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="baf98-123">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="baf98-124">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="baf98-124">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="baf98-125">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="baf98-125">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="baf98-126">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="baf98-126">`Project` (Project)</span></span>
- <span data-ttu-id="baf98-127">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="baf98-127">`Workbook` (Excel)</span></span>

> [!IMPORTANT]
> <span data-ttu-id="baf98-128">SharePoint で Access Web アプリとデータベースを作成して使用することは推奨されなくなりました。</span><span class="sxs-lookup"><span data-stu-id="baf98-128">We no longer recommend that you create and use Access web apps and databases in SharePoint.</span></span> <span data-ttu-id="baf98-129">代わりに、[Microsoft PowerApps](https://powerapps.microsoft.com/) を使用して、コード作成が不要な Web とモバイル デバイス用ビジネス ソリューションをビルドすることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="baf98-129">As an alternative, we recommend that you use [Microsoft PowerApps](https://powerapps.microsoft.com/) to build no-code business solutions for web and mobile devices.</span></span>

### <a name="example"></a><span data-ttu-id="baf98-130">例</span><span class="sxs-lookup"><span data-stu-id="baf98-130">Example</span></span>

```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="baf98-131">VersionOverrides ノード</span><span class="sxs-lookup"><span data-stu-id="baf98-131">VersionOverrides node</span></span>

<span data-ttu-id="baf98-132">[VersionOverrides](versionoverrides.md) で定義されている場合、ホストの種類は `xsi:type` 属性によって決定されます。</span><span class="sxs-lookup"><span data-stu-id="baf98-132">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="baf98-133">属性</span><span class="sxs-lookup"><span data-stu-id="baf98-133">Attributes</span></span>

|  <span data-ttu-id="baf98-134">属性</span><span class="sxs-lookup"><span data-stu-id="baf98-134">Attribute</span></span>  |  <span data-ttu-id="baf98-135">必須</span><span class="sxs-lookup"><span data-stu-id="baf98-135">Required</span></span>  |  <span data-ttu-id="baf98-136">説明</span><span class="sxs-lookup"><span data-stu-id="baf98-136">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="baf98-137">xsi:type</span><span class="sxs-lookup"><span data-stu-id="baf98-137">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="baf98-138">はい</span><span class="sxs-lookup"><span data-stu-id="baf98-138">Yes</span></span>  | <span data-ttu-id="baf98-139">これらの設定が適用Officeアプリケーションについて説明します。</span><span class="sxs-lookup"><span data-stu-id="baf98-139">Describes the Office application where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="baf98-140">子要素</span><span class="sxs-lookup"><span data-stu-id="baf98-140">Child elements</span></span>

|  <span data-ttu-id="baf98-141">要素</span><span class="sxs-lookup"><span data-stu-id="baf98-141">Element</span></span> |  <span data-ttu-id="baf98-142">必須</span><span class="sxs-lookup"><span data-stu-id="baf98-142">Required</span></span>  |  <span data-ttu-id="baf98-143">説明</span><span class="sxs-lookup"><span data-stu-id="baf98-143">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="baf98-144">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="baf98-144">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="baf98-145">はい</span><span class="sxs-lookup"><span data-stu-id="baf98-145">Yes</span></span>   |  <span data-ttu-id="baf98-146">デスクトップ フォーム ファクターの設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="baf98-146">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="baf98-147">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="baf98-147">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="baf98-148">いいえ</span><span class="sxs-lookup"><span data-stu-id="baf98-148">No</span></span>   |  <span data-ttu-id="baf98-149">モバイル フォーム ファクターの設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="baf98-149">Defines the settings for the mobile form factor.</span></span> <span data-ttu-id="baf98-150">**注:** この要素は、iOS Outlook Android でのみサポートされます。</span><span class="sxs-lookup"><span data-stu-id="baf98-150">**Note:** This element is only supported in Outlook on iOS and Android.</span></span> |
|  [<span data-ttu-id="baf98-151">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="baf98-151">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="baf98-152">いいえ</span><span class="sxs-lookup"><span data-stu-id="baf98-152">No</span></span>   |  <span data-ttu-id="baf98-153">すべてのフォーム ファクターの設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="baf98-153">Defines the settings for all form factors.</span></span> <span data-ttu-id="baf98-154">Excel のカスタム関数でのみ使用します。</span><span class="sxs-lookup"><span data-stu-id="baf98-154">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="baf98-155">xsi:type</span><span class="sxs-lookup"><span data-stu-id="baf98-155">xsi:type</span></span>

<span data-ttu-id="baf98-156">含まれているOffice適用するアプリケーション (Word、Excel、PowerPoint、Outlook、OneNote) を制御します。</span><span class="sxs-lookup"><span data-stu-id="baf98-156">Controls which Office application (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="baf98-157">この値は、次のいずれかである必要があります。</span><span class="sxs-lookup"><span data-stu-id="baf98-157">The value must be one of the following:</span></span>

- <span data-ttu-id="baf98-158">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="baf98-158">`Document` (Word)</span></span>
- <span data-ttu-id="baf98-159">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="baf98-159">`MailHost` (Outlook)</span></span>
- <span data-ttu-id="baf98-160">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="baf98-160">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="baf98-161">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="baf98-161">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="baf98-162">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="baf98-162">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="baf98-163">ホストの例</span><span class="sxs-lookup"><span data-stu-id="baf98-163">Host example</span></span>

```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
