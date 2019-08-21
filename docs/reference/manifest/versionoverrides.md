---
title: マニフェスト ファイルの VersionOverrides 要素
description: ''
ms.date: 08/12/2019
localization_priority: Normal
ms.openlocfilehash: ce65cdced1b3cf885cee09732c2cda0081a53cfc
ms.sourcegitcommit: da8e6148f4bd9884ab9702db3033273a383d15f0
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/20/2019
ms.locfileid: "36477881"
---
# <a name="versionoverrides-element"></a><span data-ttu-id="056f5-102">VersionOverrides 要素</span><span class="sxs-lookup"><span data-stu-id="056f5-102">VersionOverrides element</span></span>

<span data-ttu-id="056f5-p101">アドインによって実装されたアドイン コマンドに関する情報を格納するルート要素です。**VersionOverrides** は、マニフェスト内の [OfficeApp](./officeapp.md) 要素の子要素です。この要素は、マニフェスト スキーマ v1.1 以降でサポートされていますが、VersionOverrides v1.0 または v1.1 スキーマで定義されています。</span><span class="sxs-lookup"><span data-stu-id="056f5-p101">The root element that contains information for the add-in commands implemented by the add-in. **VersionOverrides** is a child element of the [OfficeApp](./officeapp.md) element in the manifest. This element is supported in manifest schema v1.1 and later but is defined in the VersionOverrides v1.0 or v1.1 schema.</span></span>

## <a name="attributes"></a><span data-ttu-id="056f5-106">属性</span><span class="sxs-lookup"><span data-stu-id="056f5-106">Attributes</span></span>

|  <span data-ttu-id="056f5-107">属性</span><span class="sxs-lookup"><span data-stu-id="056f5-107">Attribute</span></span>  |  <span data-ttu-id="056f5-108">必須</span><span class="sxs-lookup"><span data-stu-id="056f5-108">Required</span></span>  |  <span data-ttu-id="056f5-109">説明</span><span class="sxs-lookup"><span data-stu-id="056f5-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="056f5-110">**xmlns**</span><span class="sxs-lookup"><span data-stu-id="056f5-110">**xmlns**</span></span>       |  <span data-ttu-id="056f5-111">はい</span><span class="sxs-lookup"><span data-stu-id="056f5-111">Yes</span></span>  |  <span data-ttu-id="056f5-112">スキーマの場所。`http://schemas.microsoft.com/office/mailappversionoverrides` が `xsi:type` の場合は `VersionOverridesV1_0` にする必要があり、`http://schemas.microsoft.com/office/mailappversionoverrides/1.1` が `xsi:type` の場合は `VersionOverridesV1_1` にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="056f5-112">The schema location, which must be `http://schemas.microsoft.com/office/mailappversionoverrides` when `xsi:type` is `VersionOverridesV1_0`, and `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` when `xsi:type` is `VersionOverridesV1_1`.</span></span>|
|  <span data-ttu-id="056f5-113">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="056f5-113">**xsi:type**</span></span>  |  <span data-ttu-id="056f5-114">はい</span><span class="sxs-lookup"><span data-stu-id="056f5-114">Yes</span></span>  | <span data-ttu-id="056f5-p102">スキーマのバージョン。現時点では、`VersionOverridesV1_0` および `VersionOverridesV1_1` のみが有効な値になります。</span><span class="sxs-lookup"><span data-stu-id="056f5-p102">The schema version. At this time, the only valid values are `VersionOverridesV1_0` and `VersionOverridesV1_1`.</span></span> |

> [!NOTE]
> <span data-ttu-id="056f5-117">現在、Outlook 2016 以降では、VersionOverrides v1.1 スキーマと`VersionOverridesV1_1`種類をサポートしています。</span><span class="sxs-lookup"><span data-stu-id="056f5-117">Currently only Outlook 2016 or later supports the VersionOverrides v1.1 schema and the `VersionOverridesV1_1` type.</span></span>

## <a name="child-elements"></a><span data-ttu-id="056f5-118">子要素</span><span class="sxs-lookup"><span data-stu-id="056f5-118">Child elements</span></span>

|  <span data-ttu-id="056f5-119">要素</span><span class="sxs-lookup"><span data-stu-id="056f5-119">Element</span></span> |  <span data-ttu-id="056f5-120">必須</span><span class="sxs-lookup"><span data-stu-id="056f5-120">Required</span></span>  |  <span data-ttu-id="056f5-121">説明</span><span class="sxs-lookup"><span data-stu-id="056f5-121">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="056f5-122">**説明**</span><span class="sxs-lookup"><span data-stu-id="056f5-122">**Description**</span></span>    |  <span data-ttu-id="056f5-123">No/しない</span><span class="sxs-lookup"><span data-stu-id="056f5-123">No</span></span>   |  <span data-ttu-id="056f5-p103">アドインについての説明。これは、マニフェスト内の任意の親部分の `Description` 要素を上書きします。説明のテキストは、**Resources** 要素の [LongString](./resources.md) 要素の子要素に含まれています。`resid` 要素の \*\*\*\* の属性は、テキストを含む `id` 要素の `String` 属性の値に設定されています。</span><span class="sxs-lookup"><span data-stu-id="056f5-p103">Describes the add-in. This overrides the `Description` element in any parent portion of the manifest. The text of the description is contained in a child element of the **LongString** element contained in the [Resources](./resources.md) element. The `resid` attribute of the **Description** element is set to the value of the `id` attribute of the `String` element that contains the text.</span></span>|
| <span data-ttu-id="056f5-128">**EquivalentAddins**</span><span class="sxs-lookup"><span data-stu-id="056f5-128">**EquivalentAddins**</span></span> | <span data-ttu-id="056f5-129">いいえ</span><span class="sxs-lookup"><span data-stu-id="056f5-129">No</span></span> | <span data-ttu-id="056f5-130">同等の COM アドイン、XLL、またはその両方との下位互換性を指定します。</span><span class="sxs-lookup"><span data-stu-id="056f5-130">Specifies backwards compatibility with an equivalent COM add-in, XLL, or both.</span></span> |
|  <span data-ttu-id="056f5-131">**Requirements**</span><span class="sxs-lookup"><span data-stu-id="056f5-131">**Requirements**</span></span>  |  <span data-ttu-id="056f5-132">いいえ</span><span class="sxs-lookup"><span data-stu-id="056f5-132">No</span></span>   |  <span data-ttu-id="056f5-p104">アドインに必要な最小の Office.js のセットおよびバージョンを指定します。これは、マニフェストの親部分の `Requirements` 要素を上書きします。</span><span class="sxs-lookup"><span data-stu-id="056f5-p104">Specifies the minimum requirement set and version of Office.js that the add-in requires. This overrides the  `Requirements` element in the parent portion of the manifest.</span></span>|
|  [<span data-ttu-id="056f5-135">Hosts</span><span class="sxs-lookup"><span data-stu-id="056f5-135">Hosts</span></span>](./hosts.md)                |  <span data-ttu-id="056f5-136">はい</span><span class="sxs-lookup"><span data-stu-id="056f5-136">Yes</span></span>  |  <span data-ttu-id="056f5-p105">Office ホストのコレクションを指定します。子の Host 要素は、マニフェストの親部分の Host 要素を上書きします。</span><span class="sxs-lookup"><span data-stu-id="056f5-p105">Specifies a collection of Office hosts. The child  Hosts element overrides the Hosts element in the parent portion of the manifest.</span></span>  |
|  [<span data-ttu-id="056f5-139">Resources</span><span class="sxs-lookup"><span data-stu-id="056f5-139">Resources</span></span>](./resources.md)    |  <span data-ttu-id="056f5-140">はい</span><span class="sxs-lookup"><span data-stu-id="056f5-140">Yes</span></span>  | <span data-ttu-id="056f5-141">マニフェストの他の要素によって参照されるリソースのコレクション (文字列、URL、画像) を定義します。</span><span class="sxs-lookup"><span data-stu-id="056f5-141">Defines a collection of resources (strings, URLs, and images) that other manifest elements reference.</span></span>|
|  [<span data-ttu-id="056f5-142">EquivalentAddins</span><span class="sxs-lookup"><span data-stu-id="056f5-142">EquivalentAddins</span></span>](./equivalentaddins.md)    |  <span data-ttu-id="056f5-143">いいえ</span><span class="sxs-lookup"><span data-stu-id="056f5-143">No</span></span>  | <span data-ttu-id="056f5-144">Web アドインと同等のネイティブ (COM/XLL) アドインを指定します。</span><span class="sxs-lookup"><span data-stu-id="056f5-144">Specifies the native (COM/XLL) add-ins that are equivalent to the web add-in.</span></span> <span data-ttu-id="056f5-145">同等のネイティブアドインがインストールされている場合、web アドインはアクティブ化されません。</span><span class="sxs-lookup"><span data-stu-id="056f5-145">The web add-in is not activated if an equivalent native add-in is installed.</span></span>|
|  <span data-ttu-id="056f5-146">**VersionOverrides**</span><span class="sxs-lookup"><span data-stu-id="056f5-146">**VersionOverrides**</span></span>    |  <span data-ttu-id="056f5-147">いいえ</span><span class="sxs-lookup"><span data-stu-id="056f5-147">No</span></span>  | <span data-ttu-id="056f5-p107">より新しいスキーマ バージョンでアドイン コマンドを定義します。詳細については、「[複数のバージョンを実装する](#implementing-multiple-versions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="056f5-p107">Defines add-in commands under a newer schema version. See [Implementing multiple versions](#implementing-multiple-versions) for details.</span></span> |
|  [<span data-ttu-id="056f5-150">WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="056f5-150">WebApplicationInfo</span></span>](./webapplicationinfo.md)    |  <span data-ttu-id="056f5-151">いいえ</span><span class="sxs-lookup"><span data-stu-id="056f5-151">No</span></span>  | <span data-ttu-id="056f5-152">Azure Active Directory v2.0 など、セキュリティで保護されたトークン発行者によるアドインの登録に関する詳細を指定します。</span><span class="sxs-lookup"><span data-stu-id="056f5-152">Specifies details about the add-in's registration with secure token issuers, such as Azure Active Directory V2.0.</span></span> |

### <a name="versionoverrides-example"></a><span data-ttu-id="056f5-153">VersionOverrides の例</span><span class="sxs-lookup"><span data-stu-id="056f5-153">VersionOverrides example</span></span>

<span data-ttu-id="056f5-154">通常、必須ではありません`<VersionOverrides>`が通常使用される子要素を含む一般的な要素の例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="056f5-154">The following is an example of a typical `<VersionOverrides>` element, including some child elements that are not required but are typically used.</span></span>

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

## <a name="implementing-multiple-versions"></a><span data-ttu-id="056f5-155">複数のバージョンを実装する</span><span class="sxs-lookup"><span data-stu-id="056f5-155">Implementing multiple versions</span></span>

<span data-ttu-id="056f5-p108">1 つのマニフェストで、複数のバージョンの `VersionOverrides` 要素を実装することで、異なるバージョンの VersionOverrides スキーマをサポートできます。これは、新しいスキーマの新機能をオプションでサポートしながら、新機能をサポートしていない古いクライアントもサポートすることで実現できます。</span><span class="sxs-lookup"><span data-stu-id="056f5-p108">A manifest can implement multiple versions of the `VersionOverrides` element which support different versions of the VersionOverrides schema. This can be done to optionally support new features in a newer schema while still supporting older clients that do not support the new features.</span></span>

<span data-ttu-id="056f5-158">複数のバージョンを実装するために、新しいバージョンの `VersionOverrides` 要素は、古いバージョンの `VersionOverrides` 要素の子にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="056f5-158">In order to implement multiple versions, the `VersionOverrides` element for the newer version must be a child of the `VersionOverrides` element for the older version.</span></span> <span data-ttu-id="056f5-159">子の `VersionOverrides` 要素は、どの値も親から継承しません。</span><span class="sxs-lookup"><span data-stu-id="056f5-159">The child `VersionOverrides` element doesn't inherit any values from the parent.</span></span>

<span data-ttu-id="056f5-160">VersionOverrides v1.0 と v1.1 の両方のスキーマを実装するためのマニフェストは、次に示す例のようになります。</span><span class="sxs-lookup"><span data-stu-id="056f5-160">To implement both the VersionOverrides v1.0 and v1.1 schema, the manifest would look similar to the following example:</span></span>

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
