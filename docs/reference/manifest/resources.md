---
title: マニフェスト ファイルの Resources 要素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 7e1cd9fcb383fa4d5881917b3dd8d6dec3bbe4f8
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324828"
---
# <a name="resources-element"></a><span data-ttu-id="1643a-102">Resources 要素</span><span class="sxs-lookup"><span data-stu-id="1643a-102">Resources element</span></span>

<span data-ttu-id="1643a-p101">[VersionOverrides](versionoverrides.md) ノードのアイコン、文字列、および URL が含まれます。 マニフェスト要素によりリソースが指定されます。リソースの **id** を使用します。 それにより、特にリソースにさまざまなロケールのバージョンがあるとき、マニフェストのサイズが管理できる大きさに抑えられます。 **id** はマニフェスト内で一意にする必要があり、最大 32 文字を使用できます。</span><span class="sxs-lookup"><span data-stu-id="1643a-p101">Contains icons, strings, and URLs for the [VersionOverrides](versionoverrides.md) node. A manifest element specifies a resource by using the **id** of the resource. This helps to keep the size of the manifest manageable, especially when resources have versions for different locales. An **id** must be unique within the manifest and can have a maximum of 32 characters.</span></span>

<span data-ttu-id="1643a-107">各リソースは、特定のロケールに異なるリソースを定義する 1 つ以上の **Override** 子要素を持つことができます。</span><span class="sxs-lookup"><span data-stu-id="1643a-107">Each resource can have one or more **Override** child elements to define a different resource for a specific locale.</span></span>

## <a name="child-elements"></a><span data-ttu-id="1643a-108">子要素</span><span class="sxs-lookup"><span data-stu-id="1643a-108">Child elements</span></span>

|  <span data-ttu-id="1643a-109">要素</span><span class="sxs-lookup"><span data-stu-id="1643a-109">Element</span></span> |  <span data-ttu-id="1643a-110">型</span><span class="sxs-lookup"><span data-stu-id="1643a-110">Type</span></span>  |  <span data-ttu-id="1643a-111">説明</span><span class="sxs-lookup"><span data-stu-id="1643a-111">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="1643a-112">Images</span><span class="sxs-lookup"><span data-stu-id="1643a-112">Images</span></span>](#images)            |  <span data-ttu-id="1643a-113">image</span><span class="sxs-lookup"><span data-stu-id="1643a-113">image</span></span>   |  <span data-ttu-id="1643a-114">アイコンの画像への HTTPS URL を指定します。</span><span class="sxs-lookup"><span data-stu-id="1643a-114">Provides the HTTPS URL to an image for an icon.</span></span> |
|  <span data-ttu-id="1643a-115">**Urls**</span><span class="sxs-lookup"><span data-stu-id="1643a-115">**Urls**</span></span>                |  <span data-ttu-id="1643a-116">url</span><span class="sxs-lookup"><span data-stu-id="1643a-116">url</span></span>     |  <span data-ttu-id="1643a-p102">HTTPS URL の場所を指定します。 URL の長さは最大で 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="1643a-p102">Provides an HTTPS URL location. A URL can have a maximum of 2048 characters.</span></span> |
|  <span data-ttu-id="1643a-119">**ShortStrings**</span><span class="sxs-lookup"><span data-stu-id="1643a-119">**ShortStrings**</span></span> |  <span data-ttu-id="1643a-120">string</span><span class="sxs-lookup"><span data-stu-id="1643a-120">string</span></span>  |  <span data-ttu-id="1643a-p103">**Label** 要素と **Title** 要素のテキスト。 各 **String** には、最大 125 文字を使用できます。</span><span class="sxs-lookup"><span data-stu-id="1643a-p103">The text for **Label** and **Title** elements. Each **String** contains a maximum of 125 characters.</span></span>|
|  <span data-ttu-id="1643a-123">**LongStrings**</span><span class="sxs-lookup"><span data-stu-id="1643a-123">**LongStrings**</span></span>  |  <span data-ttu-id="1643a-124">string</span><span class="sxs-lookup"><span data-stu-id="1643a-124">string</span></span>  | <span data-ttu-id="1643a-p104">**Description** 属性のテキスト。 各 **String** には、最大 250 文字を使用できます。</span><span class="sxs-lookup"><span data-stu-id="1643a-p104">The text for **Description** attributes. Each **String** contains a maximum of 250 characters.</span></span>|

> [!NOTE]
> <span data-ttu-id="1643a-127">**Image** 要素と **Url** 要素のすべての URL で Secure Sockets Layer (SSL) を使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="1643a-127">You must use Secure Sockets Layer (SSL) for all URLs in the **Image** and **Url** elements.</span></span>

### <a name="images"></a><span data-ttu-id="1643a-128">画像</span><span class="sxs-lookup"><span data-stu-id="1643a-128">Images</span></span>
<span data-ttu-id="1643a-129">各アイコンには3つの**Images**要素があり、それぞれに3つの必須サイズがあります。</span><span class="sxs-lookup"><span data-stu-id="1643a-129">Each icon must have three **Images** elements, one for each of the three mandatory sizes:</span></span>

- <span data-ttu-id="1643a-130">16x16</span><span class="sxs-lookup"><span data-stu-id="1643a-130">16x16</span></span>
- <span data-ttu-id="1643a-131">32x32</span><span class="sxs-lookup"><span data-stu-id="1643a-131">32x32</span></span>
- <span data-ttu-id="1643a-132">80x80</span><span class="sxs-lookup"><span data-stu-id="1643a-132">80x80</span></span>

<span data-ttu-id="1643a-133">上記の他に次のサイズもサポートされていますが、指定する必要はありません。</span><span class="sxs-lookup"><span data-stu-id="1643a-133">The following additional sizes are also supported, but not required:</span></span>

- <span data-ttu-id="1643a-134">20x20</span><span class="sxs-lookup"><span data-stu-id="1643a-134">20x20</span></span>
- <span data-ttu-id="1643a-135">24x24</span><span class="sxs-lookup"><span data-stu-id="1643a-135">24x24</span></span>
- <span data-ttu-id="1643a-136">40x40</span><span class="sxs-lookup"><span data-stu-id="1643a-136">40x40</span></span>
- <span data-ttu-id="1643a-137">48x48</span><span class="sxs-lookup"><span data-stu-id="1643a-137">48x48</span></span>
- <span data-ttu-id="1643a-138">64x64</span><span class="sxs-lookup"><span data-stu-id="1643a-138">64x64</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="1643a-139">Outlook では、パフォーマンス向上のために画像リソースをキャッシュする機能が必要です。</span><span class="sxs-lookup"><span data-stu-id="1643a-139">Outlook requires the ability to cache image resources for performance purposes.</span></span> <span data-ttu-id="1643a-140">このため、画像リソースをホストするサーバーは、どんな CACHE-CONTROL ディレクティブも応答ヘッダーに追加することはできません。</span><span class="sxs-lookup"><span data-stu-id="1643a-140">For this reason, the server hosting an image resource must not add any CACHE-CONTROL directives to the response header.</span></span> <span data-ttu-id="1643a-141">これは、Outlook が汎用の画像や既定の画像を自動的に代用する原因になります。</span><span class="sxs-lookup"><span data-stu-id="1643a-141">This will result in Outlook automatically substituting a generic or default image.</span></span>    

## <a name="resources-examples"></a><span data-ttu-id="1643a-142">リソースの例</span><span class="sxs-lookup"><span data-stu-id="1643a-142">Resources examples</span></span> 

```XML
<Resources>
      <bt:Images>
        <bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp16-icon_default.png" />
        </bt:Image>
        <bt:Image id="icon1_32x32" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp32-icon_default.png" />
        </bt:Image>
        <bt:Image id="icon1_80x80" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp80-icon_default.png" />
        </bt:Image>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="residDesktopFuncUrl" DefaultValue="https://www.contoso.com/Pages/Home.aspx">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Pages/Home.aspx" />
        </bt:Url>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="residLabel" DefaultValue="GetData">
          <bt:Override Locale="ja-jp" Value="JA-JP-GetData" />
        </bt:String>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="residToolTip" DefaultValue="Get data for your document.">
          <bt:Override Locale="ja-jp" Value="JA-JP - Get data for your document." />
        </bt:String>
      </bt:LongStrings>
    </Resources>
```

```xml
<Resources>
  <bt:Images>
    <!-- Blue icon -->
    <bt:Image id="blue-icon-16" DefaultValue="YOUR_WEB_SERVER/blue-16.png"/>
    <bt:Image id="blue-icon-32" DefaultValue="YOUR_WEB_SERVER//blue-32.png"/>
    <bt:Image id="blue-icon-80" DefaultValue="YOUR_WEB_SERVER/blue-80.png"/>
  </bt:Images>
  <bt:Urls>
    <bt:Url id="functionFile" DefaultValue="YOUR_WEB_SERVER/FunctionFile/Functions.html"/>
    <!-- other URLs -->
  </bt:Urls>
  <bt:ShortStrings>
    <bt:String id="groupLabel" DefaultValue="Add-in Demo">
      <bt:Override Locale="ar-sa" Value="<Localized text>" />
    </bt:String>
    <!-- Other short strings -->
  </bt:ShortStrings>
  <bt:LongStrings>
    <bt:String id="funcReadSuperTipDescription" DefaultValue="Gets the subject of the message or appointment.">
      <bt:Override Locale="ar-sa" Value="<Localized text>." />
    </bt:String>
    <!-- Other long strings -->
  </bt:LongStrings>
</Resources>
```
