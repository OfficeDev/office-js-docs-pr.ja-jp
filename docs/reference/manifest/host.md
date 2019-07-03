---
title: マニフェスト ファイルの Host 要素
description: ''
ms.date: 07/01/2019
localization_priority: Normal
ms.openlocfilehash: e7b557034f70b03ed57598b7ffb9f43878db7392
ms.sourcegitcommit: 90c2d8236c6b30d80ac2b13950028a208ef60973
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/02/2019
ms.locfileid: "35454896"
---
# <a name="host-element"></a><span data-ttu-id="ea01f-102">Host 要素</span><span class="sxs-lookup"><span data-stu-id="ea01f-102">Host element</span></span>

<span data-ttu-id="ea01f-103">アドインでアクティブ化する Office アプリケーションの種類を個別に指定します。</span><span class="sxs-lookup"><span data-stu-id="ea01f-103">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="ea01f-104">**Host** 要素の構文は、要素が[基本のマニフェスト](#basic-manifest)で定義されているか、[VersionOverrides](#versionoverrides-node) ノードで定義されているかによって異なります。</span><span class="sxs-lookup"><span data-stu-id="ea01f-104">The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="ea01f-105">ただし、機能は変わりません。</span><span class="sxs-lookup"><span data-stu-id="ea01f-105">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="ea01f-106">基本のマニフェスト</span><span class="sxs-lookup"><span data-stu-id="ea01f-106">Basic manifest</span></span>

<span data-ttu-id="ea01f-107">基本のマニフェストで定義されている場合 ([OfficeApp](officeapp.md) の下)、ホストの種類は `Name` 属性によって決定されます。</span><span class="sxs-lookup"><span data-stu-id="ea01f-107">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>

### <a name="attributes"></a><span data-ttu-id="ea01f-108">属性</span><span class="sxs-lookup"><span data-stu-id="ea01f-108">Attributes</span></span>

| <span data-ttu-id="ea01f-109">属性</span><span class="sxs-lookup"><span data-stu-id="ea01f-109">Attribute</span></span>     | <span data-ttu-id="ea01f-110">型</span><span class="sxs-lookup"><span data-stu-id="ea01f-110">Type</span></span>   | <span data-ttu-id="ea01f-111">必須</span><span class="sxs-lookup"><span data-stu-id="ea01f-111">Required</span></span> | <span data-ttu-id="ea01f-112">説明</span><span class="sxs-lookup"><span data-stu-id="ea01f-112">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="ea01f-113">名前</span><span class="sxs-lookup"><span data-stu-id="ea01f-113">Name</span></span>](#name) | <span data-ttu-id="ea01f-114">string</span><span class="sxs-lookup"><span data-stu-id="ea01f-114">string</span></span> | <span data-ttu-id="ea01f-115">必須</span><span class="sxs-lookup"><span data-stu-id="ea01f-115">required</span></span> | <span data-ttu-id="ea01f-116">Office ホスト アプリケーションの種類の名前。</span><span class="sxs-lookup"><span data-stu-id="ea01f-116">The name of the type of Office host application.</span></span> |

### <a name="name"></a><span data-ttu-id="ea01f-117">名前</span><span class="sxs-lookup"><span data-stu-id="ea01f-117">Name</span></span>

<span data-ttu-id="ea01f-118">このアドインが対象にするホストの種類を指定します。</span><span class="sxs-lookup"><span data-stu-id="ea01f-118">Specifies the Host type targeted by this add-in.</span></span> <span data-ttu-id="ea01f-119">この値は、次のいずれかであることが必要です。</span><span class="sxs-lookup"><span data-stu-id="ea01f-119">The value must be one of the following.</span></span>

- <span data-ttu-id="ea01f-120">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="ea01f-120">`Document` (Word)</span></span>
- <span data-ttu-id="ea01f-121">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="ea01f-121">`Database` (Access)</span></span>
- <span data-ttu-id="ea01f-122">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="ea01f-122">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="ea01f-123">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="ea01f-123">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="ea01f-124">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="ea01f-124">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="ea01f-125">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="ea01f-125">`Project` (Project)</span></span>
- <span data-ttu-id="ea01f-126">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="ea01f-126">`Workbook` (Excel)</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ea01f-127">SharePoint で Access Web アプリとデータベースを作成して使用することは推奨されなくなりました。</span><span class="sxs-lookup"><span data-stu-id="ea01f-127">We no longer recommend that you create and use Access web apps and databases in SharePoint.</span></span> <span data-ttu-id="ea01f-128">代わりに、[Microsoft PowerApps](https://powerapps.microsoft.com/) を使用して、コード作成が不要な Web とモバイル デバイス用ビジネス ソリューションをビルドすることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="ea01f-128">As an alternative, we recommend that you use [Microsoft PowerApps](https://powerapps.microsoft.com/) to build no-code business solutions for web and mobile devices.</span></span>

### <a name="example"></a><span data-ttu-id="ea01f-129">例</span><span class="sxs-lookup"><span data-stu-id="ea01f-129">Example</span></span>

```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="ea01f-130">VersionOverrides ノード</span><span class="sxs-lookup"><span data-stu-id="ea01f-130">VersionOverrides node</span></span>

<span data-ttu-id="ea01f-131">[VersionOverrides](versionoverrides.md) で定義されている場合、ホストの種類は `xsi:type` 属性によって決定されます。</span><span class="sxs-lookup"><span data-stu-id="ea01f-131">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span> 

### <a name="attributes"></a><span data-ttu-id="ea01f-132">属性</span><span class="sxs-lookup"><span data-stu-id="ea01f-132">Attributes</span></span>

|  <span data-ttu-id="ea01f-133">属性</span><span class="sxs-lookup"><span data-stu-id="ea01f-133">Attribute</span></span>  |  <span data-ttu-id="ea01f-134">必須</span><span class="sxs-lookup"><span data-stu-id="ea01f-134">Required</span></span>  |  <span data-ttu-id="ea01f-135">説明</span><span class="sxs-lookup"><span data-stu-id="ea01f-135">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="ea01f-136">xsi:type</span><span class="sxs-lookup"><span data-stu-id="ea01f-136">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="ea01f-137">はい</span><span class="sxs-lookup"><span data-stu-id="ea01f-137">Yes</span></span>  | <span data-ttu-id="ea01f-138">これらの設定を適用する Office ホストについて説明します。</span><span class="sxs-lookup"><span data-stu-id="ea01f-138">Describes the Office host where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="ea01f-139">子要素</span><span class="sxs-lookup"><span data-stu-id="ea01f-139">Child elements</span></span>

|  <span data-ttu-id="ea01f-140">要素</span><span class="sxs-lookup"><span data-stu-id="ea01f-140">Element</span></span> |  <span data-ttu-id="ea01f-141">必須</span><span class="sxs-lookup"><span data-stu-id="ea01f-141">Required</span></span>  |  <span data-ttu-id="ea01f-142">説明</span><span class="sxs-lookup"><span data-stu-id="ea01f-142">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="ea01f-143">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="ea01f-143">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="ea01f-144">はい</span><span class="sxs-lookup"><span data-stu-id="ea01f-144">Yes</span></span>   |  <span data-ttu-id="ea01f-145">デスクトップ フォーム ファクターの設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="ea01f-145">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="ea01f-146">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="ea01f-146">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="ea01f-147">いいえ</span><span class="sxs-lookup"><span data-stu-id="ea01f-147">No</span></span>   |  <span data-ttu-id="ea01f-148">モバイルフォームファクターの設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="ea01f-148">Defines the settings for the mobile form factor.</span></span> <span data-ttu-id="ea01f-149">**注:** この要素は、iOS の Outlook でのみサポートされています。</span><span class="sxs-lookup"><span data-stu-id="ea01f-149">**Note:** This element is only supported in Outlook on iOS.</span></span> |
|  [<span data-ttu-id="ea01f-150">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="ea01f-150">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="ea01f-151">いいえ</span><span class="sxs-lookup"><span data-stu-id="ea01f-151">No</span></span>   |  <span data-ttu-id="ea01f-152">すべてのフォーム ファクターの設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="ea01f-152">Defines the settings for all form factors.</span></span> <span data-ttu-id="ea01f-153">Excel のカスタム関数でのみ使用します。</span><span class="sxs-lookup"><span data-stu-id="ea01f-153">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="ea01f-154">xsi:type</span><span class="sxs-lookup"><span data-stu-id="ea01f-154">xsi:type</span></span>

<span data-ttu-id="ea01f-155">含まれている設定を適用する Office ホスト (Word、Excel、PowerPoint、Outlook、OneNote) を制御します。</span><span class="sxs-lookup"><span data-stu-id="ea01f-155">Controls which Office host (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="ea01f-156">この値は、次のいずれかである必要があります。</span><span class="sxs-lookup"><span data-stu-id="ea01f-156">The value must be one of the following:</span></span>

- <span data-ttu-id="ea01f-157">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="ea01f-157">`Document` (Word)</span></span>
- <span data-ttu-id="ea01f-158">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="ea01f-158">`MailHost` (Outlook)</span></span>
- <span data-ttu-id="ea01f-159">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="ea01f-159">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="ea01f-160">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="ea01f-160">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="ea01f-161">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="ea01f-161">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="ea01f-162">ホストの例</span><span class="sxs-lookup"><span data-stu-id="ea01f-162">Host example</span></span> 

```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
