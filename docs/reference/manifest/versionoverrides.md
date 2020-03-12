---
title: マニフェスト ファイルの VersionOverrides 要素
description: ''
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: 5dc1013f24ef6e0cc4f000128b6f5d28ccae4432
ms.sourcegitcommit: a0262ea40cd23f221e69bcb0223110f011265d13
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/12/2020
ms.locfileid: "42605681"
---
# <a name="versionoverrides-element"></a><span data-ttu-id="bc892-102">VersionOverrides 要素</span><span class="sxs-lookup"><span data-stu-id="bc892-102">VersionOverrides element</span></span>

<span data-ttu-id="bc892-p101">アドインによって実装されたアドイン コマンドに関する情報を格納するルート要素です。**VersionOverrides** は、マニフェスト内の [OfficeApp](./officeapp.md) 要素の子要素です。この要素は、マニフェスト スキーマ v1.1 以降でサポートされていますが、VersionOverrides v1.0 または v1.1 スキーマで定義されています。</span><span class="sxs-lookup"><span data-stu-id="bc892-p101">The root element that contains information for the add-in commands implemented by the add-in. **VersionOverrides** is a child element of the [OfficeApp](./officeapp.md) element in the manifest. This element is supported in manifest schema v1.1 and later but is defined in the VersionOverrides v1.0 or v1.1 schema.</span></span>

## <a name="attributes"></a><span data-ttu-id="bc892-106">属性</span><span class="sxs-lookup"><span data-stu-id="bc892-106">Attributes</span></span>

|  <span data-ttu-id="bc892-107">属性</span><span class="sxs-lookup"><span data-stu-id="bc892-107">Attribute</span></span>  |  <span data-ttu-id="bc892-108">必須</span><span class="sxs-lookup"><span data-stu-id="bc892-108">Required</span></span>  |  <span data-ttu-id="bc892-109">説明</span><span class="sxs-lookup"><span data-stu-id="bc892-109">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="bc892-110">**xmlns**</span><span class="sxs-lookup"><span data-stu-id="bc892-110">**xmlns**</span></span>       |  <span data-ttu-id="bc892-111">必要</span><span class="sxs-lookup"><span data-stu-id="bc892-111">Yes</span></span>  |  <span data-ttu-id="bc892-112">VersionOverrides スキーマ名前空間。</span><span class="sxs-lookup"><span data-stu-id="bc892-112">The VersionOverrides schema namespace.</span></span> <span data-ttu-id="bc892-113">指定できる値は、 `<VersionOverrides>`この要素の**xsi: type**値と親`<OfficeApp>`要素の**xsi: type**値によって異なります。</span><span class="sxs-lookup"><span data-stu-id="bc892-113">The allowed values vary depending on  this `<VersionOverrides>` element's **xsi:type** value and the **xsi:type** value of the parent `<OfficeApp>` element.</span></span> <span data-ttu-id="bc892-114">以下の[名前空間の値](#namespace-values)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="bc892-114">See [Namespace values](#namespace-values) below.</span></span>|
|  <span data-ttu-id="bc892-115">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="bc892-115">**xsi:type**</span></span>  |  <span data-ttu-id="bc892-116">はい</span><span class="sxs-lookup"><span data-stu-id="bc892-116">Yes</span></span>  | <span data-ttu-id="bc892-p103">スキーマのバージョン。現時点では、`VersionOverridesV1_0` および `VersionOverridesV1_1` のみが有効な値になります。</span><span class="sxs-lookup"><span data-stu-id="bc892-p103">The schema version. At this time, the only valid values are `VersionOverridesV1_0` and `VersionOverridesV1_1`.</span></span> |

### <a name="namespace-values"></a><span data-ttu-id="bc892-119">名前空間の値</span><span class="sxs-lookup"><span data-stu-id="bc892-119">Namespace values</span></span>

<span data-ttu-id="bc892-120">次に、親`<OfficeApp>`要素の**xsi: type**値に応じて、 **xmlns**値に必要な値を示します。</span><span class="sxs-lookup"><span data-stu-id="bc892-120">The following lists the required value of the **xmlns** value depending on the **xsi:type** value of the parent `<OfficeApp>` element.</span></span>

- <span data-ttu-id="bc892-121">**Task区画アプリ**は、バージョン1.0 の versionoverrides のみをサポート**xmlns**し、xmlns `http://schemas.microsoft.com/office/taskpaneappversionoverrides`はにする必要があります。</span><span class="sxs-lookup"><span data-stu-id="bc892-121">**TaskPaneApp** supports only version 1.0 of VersionOverrides, and the **xmlns** should be `http://schemas.microsoft.com/office/taskpaneappversionoverrides`.</span></span>
- <span data-ttu-id="bc892-122">**Contentapp**はバージョン1.0 の versionoverrides のみをサポートし、 **xmlns**は`http://schemas.microsoft.com/office/contentappversionoverrides`である必要があります。</span><span class="sxs-lookup"><span data-stu-id="bc892-122">**ContentApp** supports only version 1.0 of VersionOverrides, and the **xmlns** should be `http://schemas.microsoft.com/office/contentappversionoverrides`.</span></span>
- <span data-ttu-id="bc892-123">**Mailapp**はバージョン1.0 および1.1 の versionoverrides をサポートしているため、 **xmlns**の値`<VersionOverrides>`は次の要素の**xsi: type**値に応じて異なります。</span><span class="sxs-lookup"><span data-stu-id="bc892-123">**MailApp** supports versions 1.0 and 1.1 of VersionOverrides, so the value of **xmlns** varies depending on this `<VersionOverrides>` element's **xsi:type** value:</span></span>
    - <span data-ttu-id="bc892-124">**Xsi: type**が`VersionOverridesV1_0`の場合、 **xmlns**は`http://schemas.microsoft.com/office/mailappversionoverrides`でなければなりません。</span><span class="sxs-lookup"><span data-stu-id="bc892-124">When **xsi:type** is `VersionOverridesV1_0`, **xmlns** must be `http://schemas.microsoft.com/office/mailappversionoverrides`.</span></span>
    - <span data-ttu-id="bc892-125">**Xsi: type**が`VersionOverridesV1_1`の場合、 **xmlns**は`http://schemas.microsoft.com/office/mailappversionoverrides/1.1`でなければなりません。</span><span class="sxs-lookup"><span data-stu-id="bc892-125">When **xsi:type** is `VersionOverridesV1_1`, **xmlns** must be `http://schemas.microsoft.com/office/mailappversionoverrides/1.1`.</span></span>

> [!NOTE]
> <span data-ttu-id="bc892-126">現在、Outlook 2016 以降では、VersionOverrides v1.1 スキーマと`VersionOverridesV1_1`種類をサポートしています。</span><span class="sxs-lookup"><span data-stu-id="bc892-126">Currently only Outlook 2016 or later supports the VersionOverrides v1.1 schema and the `VersionOverridesV1_1` type.</span></span>

## <a name="child-elements"></a><span data-ttu-id="bc892-127">子要素</span><span class="sxs-lookup"><span data-stu-id="bc892-127">Child elements</span></span>

|  <span data-ttu-id="bc892-128">要素</span><span class="sxs-lookup"><span data-stu-id="bc892-128">Element</span></span> |  <span data-ttu-id="bc892-129">必須</span><span class="sxs-lookup"><span data-stu-id="bc892-129">Required</span></span>  |  <span data-ttu-id="bc892-130">説明</span><span class="sxs-lookup"><span data-stu-id="bc892-130">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="bc892-131">**説明**</span><span class="sxs-lookup"><span data-stu-id="bc892-131">**Description**</span></span>    |  <span data-ttu-id="bc892-132">No/しない</span><span class="sxs-lookup"><span data-stu-id="bc892-132">No</span></span>   |  <span data-ttu-id="bc892-p104">アドインについての説明。これは、マニフェスト内の任意の親部分の `Description` 要素を上書きします。説明のテキストは、**Resources** 要素の [LongString](resources.md) 要素の子要素に含まれています。`resid` 要素の \*\*\*\* の属性は、テキストを含む `id` 要素の `String` 属性の値に設定されています。</span><span class="sxs-lookup"><span data-stu-id="bc892-p104">Describes the add-in. This overrides the `Description` element in any parent portion of the manifest. The text of the description is contained in a child element of the **LongString** element contained in the [Resources](resources.md) element. The `resid` attribute of the **Description** element is set to the value of the `id` attribute of the `String` element that contains the text.</span></span>|
|  <span data-ttu-id="bc892-137">**Requirements**</span><span class="sxs-lookup"><span data-stu-id="bc892-137">**Requirements**</span></span>  |  <span data-ttu-id="bc892-138">いいえ</span><span class="sxs-lookup"><span data-stu-id="bc892-138">No</span></span>   |  <span data-ttu-id="bc892-p105">アドインに必要な最小の Office.js のセットおよびバージョンを指定します。これは、マニフェストの親部分の `Requirements` 要素を上書きします。</span><span class="sxs-lookup"><span data-stu-id="bc892-p105">Specifies the minimum requirement set and version of Office.js that the add-in requires. This overrides the  `Requirements` element in the parent portion of the manifest.</span></span>|
|  [<span data-ttu-id="bc892-141">Hosts</span><span class="sxs-lookup"><span data-stu-id="bc892-141">Hosts</span></span>](hosts.md)                |  <span data-ttu-id="bc892-142">必要</span><span class="sxs-lookup"><span data-stu-id="bc892-142">Yes</span></span>  |  <span data-ttu-id="bc892-p106">Office ホストのコレクションを指定します。子の Host 要素は、マニフェストの親部分の Host 要素を上書きします。</span><span class="sxs-lookup"><span data-stu-id="bc892-p106">Specifies a collection of Office hosts. The child  Hosts element overrides the Hosts element in the parent portion of the manifest.</span></span>  |
|  [<span data-ttu-id="bc892-145">Resources</span><span class="sxs-lookup"><span data-stu-id="bc892-145">Resources</span></span>](resources.md)    |  <span data-ttu-id="bc892-146">はい</span><span class="sxs-lookup"><span data-stu-id="bc892-146">Yes</span></span>  | <span data-ttu-id="bc892-147">マニフェストの他の要素によって参照されるリソースのコレクション (文字列、URL、画像) を定義します。</span><span class="sxs-lookup"><span data-stu-id="bc892-147">Defines a collection of resources (strings, URLs, and images) that other manifest elements reference.</span></span>|
|  [<span data-ttu-id="bc892-148">EquivalentAddins</span><span class="sxs-lookup"><span data-stu-id="bc892-148">EquivalentAddins</span></span>](equivalentaddins.md)    |  <span data-ttu-id="bc892-149">いいえ</span><span class="sxs-lookup"><span data-stu-id="bc892-149">No</span></span>  | <span data-ttu-id="bc892-150">Web アドインと同等のネイティブ (COM/XLL) アドインを指定します。</span><span class="sxs-lookup"><span data-stu-id="bc892-150">Specifies the native (COM/XLL) add-ins that are equivalent to the web add-in.</span></span> <span data-ttu-id="bc892-151">同等のネイティブアドインがインストールされている場合、web アドインはアクティブ化されません。</span><span class="sxs-lookup"><span data-stu-id="bc892-151">The web add-in is not activated if an equivalent native add-in is installed.</span></span>|
|  <span data-ttu-id="bc892-152">**VersionOverrides**</span><span class="sxs-lookup"><span data-stu-id="bc892-152">**VersionOverrides**</span></span>    |  <span data-ttu-id="bc892-153">いいえ</span><span class="sxs-lookup"><span data-stu-id="bc892-153">No</span></span>  | <span data-ttu-id="bc892-p108">より新しいスキーマ バージョンでアドイン コマンドを定義します。詳細については、「[複数のバージョンを実装する](#implementing-multiple-versions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="bc892-p108">Defines add-in commands under a newer schema version. See [Implementing multiple versions](#implementing-multiple-versions) for details.</span></span> |
|  [<span data-ttu-id="bc892-156">WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="bc892-156">WebApplicationInfo</span></span>](webapplicationinfo.md)    |  <span data-ttu-id="bc892-157">いいえ</span><span class="sxs-lookup"><span data-stu-id="bc892-157">No</span></span>  | <span data-ttu-id="bc892-158">Azure Active Directory v2.0 など、セキュリティで保護されたトークン発行者によるアドインの登録に関する詳細を指定します。</span><span class="sxs-lookup"><span data-stu-id="bc892-158">Specifies details about the add-in's registration with secure token issuers, such as Azure Active Directory V2.0.</span></span> |
|  [<span data-ttu-id="bc892-159">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="bc892-159">ExtendedPermissions</span></span>](extendedpermissions.md) |  <span data-ttu-id="bc892-160">いいえ</span><span class="sxs-lookup"><span data-stu-id="bc892-160">No</span></span>  |  <span data-ttu-id="bc892-161">拡張アクセス許可のコレクションを指定します。</span><span class="sxs-lookup"><span data-stu-id="bc892-161">Specifies a collection of extended permissions.</span></span><br><br><span data-ttu-id="bc892-162">**重要**: [Office. appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-) API は現在プレビュー段階のため、この`ExtendedPermissions`要素を使用するアドインは、appsource に発行することも、一元展開によって展開することもできません。</span><span class="sxs-lookup"><span data-stu-id="bc892-162">**Important**: Because the [Office.Body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-) API is currently in preview, add-ins that use the `ExtendedPermissions` element can't be published to AppSource or deployed via centralized deployment.</span></span> |

### <a name="versionoverrides-example"></a><span data-ttu-id="bc892-163">VersionOverrides の例</span><span class="sxs-lookup"><span data-stu-id="bc892-163">VersionOverrides example</span></span>

<span data-ttu-id="bc892-164">通常、必須ではありません`<VersionOverrides>`が通常使用される子要素を含む一般的な要素の例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="bc892-164">The following is an example of a typical `<VersionOverrides>` element, including some child elements that are not required but are typically used.</span></span>

```xml
<OfficeApp ... xsi:type="MailApp">
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

## <a name="implementing-multiple-versions"></a><span data-ttu-id="bc892-165">複数のバージョンを実装する</span><span class="sxs-lookup"><span data-stu-id="bc892-165">Implementing multiple versions</span></span>

<span data-ttu-id="bc892-p109">1 つのマニフェストで、複数のバージョンの `VersionOverrides` 要素を実装することで、異なるバージョンの VersionOverrides スキーマをサポートできます。これは、新しいスキーマの新機能をオプションでサポートしながら、新機能をサポートしていない古いクライアントもサポートすることで実現できます。</span><span class="sxs-lookup"><span data-stu-id="bc892-p109">A manifest can implement multiple versions of the `VersionOverrides` element which support different versions of the VersionOverrides schema. This can be done to optionally support new features in a newer schema while still supporting older clients that do not support the new features.</span></span>

<span data-ttu-id="bc892-168">複数のバージョンを実装するために、新しいバージョンの `VersionOverrides` 要素は、古いバージョンの `VersionOverrides` 要素の子にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="bc892-168">In order to implement multiple versions, the `VersionOverrides` element for the newer version must be a child of the `VersionOverrides` element for the older version.</span></span> <span data-ttu-id="bc892-169">子の `VersionOverrides` 要素は、どの値も親から継承しません。</span><span class="sxs-lookup"><span data-stu-id="bc892-169">The child `VersionOverrides` element doesn't inherit any values from the parent.</span></span>

<span data-ttu-id="bc892-170">VersionOverrides v1.0 と v1.1 の両方のスキーマを実装するためのマニフェストは、次に示す例のようになります。</span><span class="sxs-lookup"><span data-stu-id="bc892-170">To implement both the VersionOverrides v1.0 and v1.1 schema, the manifest would look similar to the following example:</span></span>

```xml
<OfficeApp ... xsi:type="MailApp">
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
