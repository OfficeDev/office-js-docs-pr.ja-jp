---
title: マニフェスト ファイルの Resources 要素
description: Resources 要素には、VersionOverrides ノードのアイコン、文字列、URL が含まれます。
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: bdf73420345ca4d054438bfba5217254e6682e6d
ms.sourcegitcommit: 0bff0411d8cfefd4bb00c189643358e6fb1df95e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/07/2021
ms.locfileid: "51604618"
---
# <a name="resources-element"></a><span data-ttu-id="dd92d-103">リソース要素</span><span class="sxs-lookup"><span data-stu-id="dd92d-103">Resources element</span></span>

<span data-ttu-id="dd92d-p101">[VersionOverrides](versionoverrides.md) ノードのアイコン、文字列、および URL が含まれます。 マニフェスト要素によりリソースが指定されます。リソースの **id** を使用します。 それにより、特にリソースにさまざまなロケールのバージョンがあるとき、マニフェストのサイズが管理できる大きさに抑えられます。 **id** はマニフェスト内で一意にする必要があり、最大 32 文字を使用できます。</span><span class="sxs-lookup"><span data-stu-id="dd92d-p101">Contains icons, strings, and URLs for the [VersionOverrides](versionoverrides.md) node. A manifest element specifies a resource by using the **id** of the resource. This helps to keep the size of the manifest manageable, especially when resources have versions for different locales. An **id** must be unique within the manifest and can have a maximum of 32 characters.</span></span>

<span data-ttu-id="dd92d-108">各リソースは、特定のロケールに異なるリソースを定義する 1 つ以上の **Override** 子要素を持つことができます。</span><span class="sxs-lookup"><span data-stu-id="dd92d-108">Each resource can have one or more **Override** child elements to define a different resource for a specific locale.</span></span>

## <a name="child-elements"></a><span data-ttu-id="dd92d-109">子要素</span><span class="sxs-lookup"><span data-stu-id="dd92d-109">Child elements</span></span>

|  <span data-ttu-id="dd92d-110">要素</span><span class="sxs-lookup"><span data-stu-id="dd92d-110">Element</span></span> |  <span data-ttu-id="dd92d-111">型</span><span class="sxs-lookup"><span data-stu-id="dd92d-111">Type</span></span>  |  <span data-ttu-id="dd92d-112">説明</span><span class="sxs-lookup"><span data-stu-id="dd92d-112">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="dd92d-113">Images</span><span class="sxs-lookup"><span data-stu-id="dd92d-113">Images</span></span>](#images)            |  <span data-ttu-id="dd92d-114">image</span><span class="sxs-lookup"><span data-stu-id="dd92d-114">image</span></span>   |  <span data-ttu-id="dd92d-115">アイコンの画像への HTTPS URL を指定します。</span><span class="sxs-lookup"><span data-stu-id="dd92d-115">Provides the HTTPS URL to an image for an icon.</span></span> |
|  <span data-ttu-id="dd92d-116">**Urls**</span><span class="sxs-lookup"><span data-stu-id="dd92d-116">**Urls**</span></span>                |  <span data-ttu-id="dd92d-117">url</span><span class="sxs-lookup"><span data-stu-id="dd92d-117">url</span></span>     |  <span data-ttu-id="dd92d-p102">HTTPS URL の場所を指定します。 URL の長さは最大で 2048 文字です。</span><span class="sxs-lookup"><span data-stu-id="dd92d-p102">Provides an HTTPS URL location. A URL can have a maximum of 2048 characters.</span></span> |
|  <span data-ttu-id="dd92d-120">**ShortStrings**</span><span class="sxs-lookup"><span data-stu-id="dd92d-120">**ShortStrings**</span></span> |  <span data-ttu-id="dd92d-121">string</span><span class="sxs-lookup"><span data-stu-id="dd92d-121">string</span></span>  |  <span data-ttu-id="dd92d-p103">**Label** 要素と **Title** 要素のテキスト。 各 **String** には、最大 125 文字を使用できます。</span><span class="sxs-lookup"><span data-stu-id="dd92d-p103">The text for **Label** and **Title** elements. Each **String** contains a maximum of 125 characters.</span></span>|
|  <span data-ttu-id="dd92d-124">**LongStrings**</span><span class="sxs-lookup"><span data-stu-id="dd92d-124">**LongStrings**</span></span>  |  <span data-ttu-id="dd92d-125">string</span><span class="sxs-lookup"><span data-stu-id="dd92d-125">string</span></span>  | <span data-ttu-id="dd92d-p104">**Description** 属性のテキスト。 各 **String** には、最大 250 文字を使用できます。</span><span class="sxs-lookup"><span data-stu-id="dd92d-p104">The text for **Description** attributes. Each **String** contains a maximum of 250 characters.</span></span>|

> [!NOTE]
> <span data-ttu-id="dd92d-128">**Image** 要素と **Url** 要素のすべての URL で Secure Sockets Layer (SSL) を使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="dd92d-128">You must use Secure Sockets Layer (SSL) for all URLs in the **Image** and **Url** elements.</span></span>

### <a name="images"></a><span data-ttu-id="dd92d-129">画像</span><span class="sxs-lookup"><span data-stu-id="dd92d-129">Images</span></span>

<span data-ttu-id="dd92d-130">各アイコンには、3 つの必須サイズごとに 1 つの **3** つの Images 要素が必要です。</span><span class="sxs-lookup"><span data-stu-id="dd92d-130">Each icon must have three **Images** elements, one for each of the three mandatory sizes:</span></span>

- <span data-ttu-id="dd92d-131">16x16</span><span class="sxs-lookup"><span data-stu-id="dd92d-131">16x16</span></span>
- <span data-ttu-id="dd92d-132">32x32</span><span class="sxs-lookup"><span data-stu-id="dd92d-132">32x32</span></span>
- <span data-ttu-id="dd92d-133">80x80</span><span class="sxs-lookup"><span data-stu-id="dd92d-133">80x80</span></span>

<span data-ttu-id="dd92d-134">上記の他に次のサイズもサポートされていますが、指定する必要はありません。</span><span class="sxs-lookup"><span data-stu-id="dd92d-134">The following additional sizes are also supported, but not required:</span></span>

- <span data-ttu-id="dd92d-135">20x20</span><span class="sxs-lookup"><span data-stu-id="dd92d-135">20x20</span></span>
- <span data-ttu-id="dd92d-136">24x24</span><span class="sxs-lookup"><span data-stu-id="dd92d-136">24x24</span></span>
- <span data-ttu-id="dd92d-137">40x40</span><span class="sxs-lookup"><span data-stu-id="dd92d-137">40x40</span></span>
- <span data-ttu-id="dd92d-138">48x48</span><span class="sxs-lookup"><span data-stu-id="dd92d-138">48x48</span></span>
- <span data-ttu-id="dd92d-139">64x64</span><span class="sxs-lookup"><span data-stu-id="dd92d-139">64x64</span></span>

> [!IMPORTANT]
>
> - <span data-ttu-id="dd92d-140">この画像がアドインの代表的なアイコンである場合は、「Create effective listings in [AppSource](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in) and within Office サイズと他の要件」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="dd92d-140">If this image is your add-in's representative icon, see [Create effective listings in AppSource and within Office](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in) for size and other requirements.</span></span>
> - <span data-ttu-id="dd92d-141">Outlook では、パフォーマンス向上のために画像リソースをキャッシュする機能が必要です。</span><span class="sxs-lookup"><span data-stu-id="dd92d-141">Outlook requires the ability to cache image resources for performance purposes.</span></span> <span data-ttu-id="dd92d-142">このため、画像リソースをホストするサーバーは、どんな CACHE-CONTROL ディレクティブも応答ヘッダーに追加することはできません。</span><span class="sxs-lookup"><span data-stu-id="dd92d-142">For this reason, the server hosting an image resource must not add any CACHE-CONTROL directives to the response header.</span></span> <span data-ttu-id="dd92d-143">これは、Outlook が汎用の画像や既定の画像を自動的に代用する原因になります。</span><span class="sxs-lookup"><span data-stu-id="dd92d-143">This will result in Outlook automatically substituting a generic or default image.</span></span>

## <a name="resources-examples"></a><span data-ttu-id="dd92d-144">リソースの例</span><span class="sxs-lookup"><span data-stu-id="dd92d-144">Resources examples</span></span>

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
