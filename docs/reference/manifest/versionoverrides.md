---
title: マニフェスト ファイルの VersionOverrides 要素
description: ''
ms.date: 01/29/2019
localization_priority: Normal
ms.openlocfilehash: 897c2203ef6ae84911b7f269ee8a2c88aec36bd0
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452068"
---
# <a name="versionoverrides-element"></a><span data-ttu-id="5ae0f-102">VersionOverrides 要素</span><span class="sxs-lookup"><span data-stu-id="5ae0f-102">VersionOverrides element</span></span>

<span data-ttu-id="5ae0f-p101">アドインによって実装されたアドイン コマンドに関する情報を格納するルート要素です。**VersionOverrides** は、マニフェスト内の [OfficeApp](./officeapp.md) 要素の子要素です。この要素は、マニフェスト スキーマ v1.1 以降でサポートされていますが、VersionOverrides v1.0 または v1.1 スキーマで定義されています。</span><span class="sxs-lookup"><span data-stu-id="5ae0f-p101">The root element that contains information for the add-in commands implemented by the add-in. **VersionOverrides** is a child element of the [OfficeApp](./officeapp.md) element in the manifest. This element is supported in manifest schema v1.1 and later but is defined in the VersionOverrides v1.0 or v1.1 schema.</span></span>

## <a name="attributes"></a><span data-ttu-id="5ae0f-106">属性</span><span class="sxs-lookup"><span data-stu-id="5ae0f-106">Attributes</span></span>

|  <span data-ttu-id="5ae0f-107">属性</span><span class="sxs-lookup"><span data-stu-id="5ae0f-107">Attribute</span></span>  |  <span data-ttu-id="5ae0f-108">必須</span><span class="sxs-lookup"><span data-stu-id="5ae0f-108">Required</span></span>  |  <span data-ttu-id="5ae0f-109">説明</span><span class="sxs-lookup"><span data-stu-id="5ae0f-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="5ae0f-110">**xmlns**</span><span class="sxs-lookup"><span data-stu-id="5ae0f-110">**xmlns**</span></span>       |  <span data-ttu-id="5ae0f-111">はい</span><span class="sxs-lookup"><span data-stu-id="5ae0f-111">Yes</span></span>  |  <span data-ttu-id="5ae0f-112">スキーマの場所。`http://schemas.microsoft.com/office/mailappversionoverrides` が `xsi:type` の場合は `VersionOverridesV1_0` にする必要があり、`http://schemas.microsoft.com/office/mailappversionoverrides/1.1` が `xsi:type` の場合は `VersionOverridesV1_1` にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="5ae0f-112">The schema location, which must be `http://schemas.microsoft.com/office/mailappversionoverrides` when `xsi:type` is `VersionOverridesV1_0`, and `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` when `xsi:type` is `VersionOverridesV1_1`.</span></span>|
|  <span data-ttu-id="5ae0f-113">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="5ae0f-113">**xsi:type**</span></span>  |  <span data-ttu-id="5ae0f-114">はい</span><span class="sxs-lookup"><span data-stu-id="5ae0f-114">Yes</span></span>  | <span data-ttu-id="5ae0f-p102">スキーマのバージョン。現時点では、`VersionOverridesV1_0` および `VersionOverridesV1_1` のみが有効な値になります。</span><span class="sxs-lookup"><span data-stu-id="5ae0f-p102">The schema version. At this time, the only valid values are `VersionOverridesV1_0` and `VersionOverridesV1_1`.</span></span> |

> [!NOTE]
> <span data-ttu-id="5ae0f-117">現在、Outlook 2016 以降では、versionoverrides v1.1 スキーマと`VersionOverridesV1_1`種類をサポートしています。</span><span class="sxs-lookup"><span data-stu-id="5ae0f-117">Currently only Outlook 2016 or later supports the VersionOverrides v1.1 schema and the `VersionOverridesV1_1` type.</span></span>

## <a name="child-elements"></a><span data-ttu-id="5ae0f-118">子要素</span><span class="sxs-lookup"><span data-stu-id="5ae0f-118">Child elements</span></span>

|  <span data-ttu-id="5ae0f-119">要素</span><span class="sxs-lookup"><span data-stu-id="5ae0f-119">Element</span></span> |  <span data-ttu-id="5ae0f-120">必須</span><span class="sxs-lookup"><span data-stu-id="5ae0f-120">Required</span></span>  |  <span data-ttu-id="5ae0f-121">説明</span><span class="sxs-lookup"><span data-stu-id="5ae0f-121">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="5ae0f-122">**説明**</span><span class="sxs-lookup"><span data-stu-id="5ae0f-122">**Description**</span></span>    |  <span data-ttu-id="5ae0f-123">No/しない</span><span class="sxs-lookup"><span data-stu-id="5ae0f-123">No</span></span>   |  <span data-ttu-id="5ae0f-p103">アドインについての説明。これは、マニフェスト内の任意の親部分の `Description` 要素を上書きします。説明のテキストは、**Resources** 要素の [LongString](./resources.md) 要素の子要素に含まれています。`resid` 要素の \*\*\*\* の属性は、テキストを含む `id` 要素の `String` 属性の値に設定されています。</span><span class="sxs-lookup"><span data-stu-id="5ae0f-p103">Describes the add-in. This overrides the `Description` element in any parent portion of the manifest. The text of the description is contained in a child element of the **LongString** element contained in the [Resources](./resources.md) element. The `resid` attribute of the **Description** element is set to the value of the `id` attribute of the `String` element that contains the text.</span></span>|
|  <span data-ttu-id="5ae0f-128">**Requirements**</span><span class="sxs-lookup"><span data-stu-id="5ae0f-128">**Requirements**</span></span>  |  <span data-ttu-id="5ae0f-129">いいえ</span><span class="sxs-lookup"><span data-stu-id="5ae0f-129">No</span></span>   |  <span data-ttu-id="5ae0f-p104">アドインに必要な最小の Office.js のセットおよびバージョンを指定します。これは、マニフェストの親部分の `Requirements` 要素を上書きします。</span><span class="sxs-lookup"><span data-stu-id="5ae0f-p104">Specifies the minimum requirement set and version of Office.js that the add-in requires. This overrides the  `Requirements` element in the parent portion of the manifest.</span></span>|
|  [<span data-ttu-id="5ae0f-132">Hosts</span><span class="sxs-lookup"><span data-stu-id="5ae0f-132">Hosts</span></span>](./hosts.md)                |  <span data-ttu-id="5ae0f-133">はい</span><span class="sxs-lookup"><span data-stu-id="5ae0f-133">Yes</span></span>  |  <span data-ttu-id="5ae0f-p105">Office ホストのコレクションを指定します。子の Host 要素は、マニフェストの親部分の Host 要素を上書きします。</span><span class="sxs-lookup"><span data-stu-id="5ae0f-p105">Specifies a collection of Office hosts. The child  Hosts element overrides the Hosts element in the parent portion of the manifest.</span></span>  |
|  [<span data-ttu-id="5ae0f-136">Resources</span><span class="sxs-lookup"><span data-stu-id="5ae0f-136">Resources</span></span>](./resources.md)    |  <span data-ttu-id="5ae0f-137">はい</span><span class="sxs-lookup"><span data-stu-id="5ae0f-137">Yes</span></span>  | <span data-ttu-id="5ae0f-138">マニフェストの他の要素によって参照されるリソースのコレクション (文字列、URL、画像) を定義します。</span><span class="sxs-lookup"><span data-stu-id="5ae0f-138">Defines a collection of resources (strings, URLs, and images) that other manifest elements reference.</span></span>|
|  <span data-ttu-id="5ae0f-139">**VersionOverrides**</span><span class="sxs-lookup"><span data-stu-id="5ae0f-139">**VersionOverrides**</span></span>    |  <span data-ttu-id="5ae0f-140">いいえ</span><span class="sxs-lookup"><span data-stu-id="5ae0f-140">No</span></span>  | <span data-ttu-id="5ae0f-p106">より新しいスキーマ バージョンでアドイン コマンドを定義します。詳細については、「[複数のバージョンを実装する](#implementing-multiple-versions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5ae0f-p106">Defines add-in commands under a newer schema version. See [Implementing multiple versions](#implementing-multiple-versions) for details.</span></span> |
|  <span data-ttu-id="5ae0f-143">**WebApplicationInfo**</span><span class="sxs-lookup"><span data-stu-id="5ae0f-143">**WebApplicationInfo**</span></span>    |  <span data-ttu-id="5ae0f-144">いいえ</span><span class="sxs-lookup"><span data-stu-id="5ae0f-144">No</span></span>  | <span data-ttu-id="5ae0f-145">アドインの関連 Web アプリケーションについての詳細を指定します。</span><span class="sxs-lookup"><span data-stu-id="5ae0f-145">Specifies details about the add-in's associated Web application.</span></span> |

### <a name="versionoverrides-example"></a><span data-ttu-id="5ae0f-146">VersionOverrides の例</span><span class="sxs-lookup"><span data-stu-id="5ae0f-146">VersionOverrides example</span></span>

<span data-ttu-id="5ae0f-147">通常、必須ではありません`<VersionOverrides>`が通常使用される子要素を含む一般的な要素の例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="5ae0f-147">The following is an example of a typical `<VersionOverrides>` element, including some child elements that are not required but are typically used.</span></span>

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources>
      <!-- add information on resources -->
    </Resources>
  </VersionOverrides>
...
</OfficeApp>
```

## <a name="implementing-multiple-versions"></a><span data-ttu-id="5ae0f-148">複数のバージョンを実装する</span><span class="sxs-lookup"><span data-stu-id="5ae0f-148">Implementing multiple versions</span></span>

<span data-ttu-id="5ae0f-p107">1 つのマニフェストで、複数のバージョンの `VersionOverrides` 要素を実装することで、異なるバージョンの VersionOverrides スキーマをサポートできます。これは、新しいスキーマの新機能をオプションでサポートしながら、新機能をサポートしていない古いクライアントもサポートすることで実現できます。</span><span class="sxs-lookup"><span data-stu-id="5ae0f-p107">A manifest can implement multiple versions of the `VersionOverrides` element which support different versions of the VersionOverrides schema. This can be done to optionally support new features in a newer schema while still supporting older clients that do not support the new features.</span></span>

<span data-ttu-id="5ae0f-151">複数のバージョンを実装するために、新しいバージョンの `VersionOverrides` 要素は、古いバージョンの `VersionOverrides` 要素の子にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="5ae0f-151">In order to implement multiple versions, the `VersionOverrides` element for the newer version must be a child of the `VersionOverrides` element for the older version.</span></span> <span data-ttu-id="5ae0f-152">子の `VersionOverrides` 要素は、どの値も親から継承しません。</span><span class="sxs-lookup"><span data-stu-id="5ae0f-152">The child `VersionOverrides` element doesn't inherit any values from the parent.</span></span>

<span data-ttu-id="5ae0f-153">VersionOverrides v1.0 と v1.1 の両方のスキーマを実装するためのマニフェストは、次に示す例のようになります。</span><span class="sxs-lookup"><span data-stu-id="5ae0f-153">To implement both the VersionOverrides v1.0 and v1.1 schema, the manifest would look similar to the following example:</span></span>

```xml
<OfficeApp>
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources>
      <!-- add information on resources -->
    </Resources>

    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
      <Description resid="residDescription" />
      <Requirements>
        <!-- add information on requirements -->
      </Requirements>
      <Hosts>
        <Host xsi:type="MailHost">
          <!-- add information on form factors -->
        </Host>
      </Hosts>
      <Resources>
        <!-- add information on resources -->
      </Resources>
    </VersionOverrides>  
  </VersionOverrides>
...
</OfficeApp>
```
