---
title: マニフェスト ファイルの VersionOverrides 要素
description: Office アドインのマニフェスト (XML) ファイルの VersionOverrides 要素の参照ドキュメント。
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: 588f0074941b41a617dd912d78ed2ef2c59f0886
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293836"
---
# <a name="versionoverrides-element"></a><span data-ttu-id="56f61-103">VersionOverrides 要素</span><span class="sxs-lookup"><span data-stu-id="56f61-103">VersionOverrides element</span></span>

<span data-ttu-id="56f61-p101">アドインによって実装されたアドイン コマンドに関する情報を格納するルート要素です。**VersionOverrides** は、マニフェスト内の [OfficeApp](./officeapp.md) 要素の子要素です。この要素は、マニフェスト スキーマ v1.1 以降でサポートされていますが、VersionOverrides v1.0 または v1.1 スキーマで定義されています。</span><span class="sxs-lookup"><span data-stu-id="56f61-p101">The root element that contains information for the add-in commands implemented by the add-in. **VersionOverrides** is a child element of the [OfficeApp](./officeapp.md) element in the manifest. This element is supported in manifest schema v1.1 and later but is defined in the VersionOverrides v1.0 or v1.1 schema.</span></span>

## <a name="attributes"></a><span data-ttu-id="56f61-107">属性</span><span class="sxs-lookup"><span data-stu-id="56f61-107">Attributes</span></span>

|  <span data-ttu-id="56f61-108">属性</span><span class="sxs-lookup"><span data-stu-id="56f61-108">Attribute</span></span>  |  <span data-ttu-id="56f61-109">必須</span><span class="sxs-lookup"><span data-stu-id="56f61-109">Required</span></span>  |  <span data-ttu-id="56f61-110">説明</span><span class="sxs-lookup"><span data-stu-id="56f61-110">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="56f61-111">**xmlns**</span><span class="sxs-lookup"><span data-stu-id="56f61-111">**xmlns**</span></span>       |  <span data-ttu-id="56f61-112">はい</span><span class="sxs-lookup"><span data-stu-id="56f61-112">Yes</span></span>  |  <span data-ttu-id="56f61-113">VersionOverrides スキーマ名前空間。</span><span class="sxs-lookup"><span data-stu-id="56f61-113">The VersionOverrides schema namespace.</span></span> <span data-ttu-id="56f61-114">指定できる値は、この `<VersionOverrides>` 要素の **xsi: type** 値と親要素の **xsi: type** 値によって異なり `<OfficeApp>` ます。</span><span class="sxs-lookup"><span data-stu-id="56f61-114">The allowed values vary depending on  this `<VersionOverrides>` element's **xsi:type** value and the **xsi:type** value of the parent `<OfficeApp>` element.</span></span> <span data-ttu-id="56f61-115">以下の [名前空間の値](#namespace-values) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="56f61-115">See [Namespace values](#namespace-values) below.</span></span>|
|  <span data-ttu-id="56f61-116">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="56f61-116">**xsi:type**</span></span>  |  <span data-ttu-id="56f61-117">はい</span><span class="sxs-lookup"><span data-stu-id="56f61-117">Yes</span></span>  | <span data-ttu-id="56f61-p103">スキーマのバージョン。現時点では、`VersionOverridesV1_0` および `VersionOverridesV1_1` のみが有効な値になります。</span><span class="sxs-lookup"><span data-stu-id="56f61-p103">The schema version. At this time, the only valid values are `VersionOverridesV1_0` and `VersionOverridesV1_1`.</span></span> |

### <a name="namespace-values"></a><span data-ttu-id="56f61-120">名前空間の値</span><span class="sxs-lookup"><span data-stu-id="56f61-120">Namespace values</span></span>

<span data-ttu-id="56f61-121">次に、親要素の**xsi: type**値に応じて、 **xmlns**値に必要な値を示し `<OfficeApp>` ます。</span><span class="sxs-lookup"><span data-stu-id="56f61-121">The following lists the required value of the **xmlns** value depending on the **xsi:type** value of the parent `<OfficeApp>` element.</span></span>

- <span data-ttu-id="56f61-122">**Task区画アプリ** は、バージョン1.0 の versionoverrides のみをサポートし、 **xmlns** はにする必要があり `http://schemas.microsoft.com/office/taskpaneappversionoverrides` ます。</span><span class="sxs-lookup"><span data-stu-id="56f61-122">**TaskPaneApp** supports only version 1.0 of VersionOverrides, and the **xmlns** should be `http://schemas.microsoft.com/office/taskpaneappversionoverrides`.</span></span>
- <span data-ttu-id="56f61-123">**Contentapp** はバージョン1.0 の versionoverrides のみをサポートし、 **xmlns** はである必要があり `http://schemas.microsoft.com/office/contentappversionoverrides` ます。</span><span class="sxs-lookup"><span data-stu-id="56f61-123">**ContentApp** supports only version 1.0 of VersionOverrides, and the **xmlns** should be `http://schemas.microsoft.com/office/contentappversionoverrides`.</span></span>
- <span data-ttu-id="56f61-124">**Mailapp** はバージョン1.0 および1.1 の versionoverrides をサポートしているため、 **xmlns** の値は次の `<VersionOverrides>` 要素の **xsi: type** 値に応じて異なります。</span><span class="sxs-lookup"><span data-stu-id="56f61-124">**MailApp** supports versions 1.0 and 1.1 of VersionOverrides, so the value of **xmlns** varies depending on this `<VersionOverrides>` element's **xsi:type** value:</span></span>
    - <span data-ttu-id="56f61-125">**Xsi: type**がの場合 `VersionOverridesV1_0` 、 **xmlns**はでなければなりません `http://schemas.microsoft.com/office/mailappversionoverrides` 。</span><span class="sxs-lookup"><span data-stu-id="56f61-125">When **xsi:type** is `VersionOverridesV1_0`, **xmlns** must be `http://schemas.microsoft.com/office/mailappversionoverrides`.</span></span>
    - <span data-ttu-id="56f61-126">**Xsi: type**がの場合 `VersionOverridesV1_1` 、 **xmlns**はでなければなりません `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` 。</span><span class="sxs-lookup"><span data-stu-id="56f61-126">When **xsi:type** is `VersionOverridesV1_1`, **xmlns** must be `http://schemas.microsoft.com/office/mailappversionoverrides/1.1`.</span></span>

> [!NOTE]
> <span data-ttu-id="56f61-127">現在、Outlook 2016 以降では、VersionOverrides v1.1 スキーマと種類をサポートしてい `VersionOverridesV1_1` ます。</span><span class="sxs-lookup"><span data-stu-id="56f61-127">Currently only Outlook 2016 or later supports the VersionOverrides v1.1 schema and the `VersionOverridesV1_1` type.</span></span>

## <a name="child-elements"></a><span data-ttu-id="56f61-128">子要素</span><span class="sxs-lookup"><span data-stu-id="56f61-128">Child elements</span></span>

|  <span data-ttu-id="56f61-129">要素</span><span class="sxs-lookup"><span data-stu-id="56f61-129">Element</span></span> |  <span data-ttu-id="56f61-130">必須</span><span class="sxs-lookup"><span data-stu-id="56f61-130">Required</span></span>  |  <span data-ttu-id="56f61-131">説明</span><span class="sxs-lookup"><span data-stu-id="56f61-131">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="56f61-132">**説明**</span><span class="sxs-lookup"><span data-stu-id="56f61-132">**Description**</span></span>    |  <span data-ttu-id="56f61-133">いいえ</span><span class="sxs-lookup"><span data-stu-id="56f61-133">No</span></span>   |  <span data-ttu-id="56f61-p104">アドインについての説明。これは、マニフェスト内の任意の親部分の `Description` 要素を上書きします。説明のテキストは、**Resources** 要素の [LongString](resources.md) 要素の子要素に含まれています。`resid` 要素の \*\*\*\* の属性は、テキストを含む `id` 要素の `String` 属性の値に設定されています。</span><span class="sxs-lookup"><span data-stu-id="56f61-p104">Describes the add-in. This overrides the `Description` element in any parent portion of the manifest. The text of the description is contained in a child element of the **LongString** element contained in the [Resources](resources.md) element. The `resid` attribute of the **Description** element is set to the value of the `id` attribute of the `String` element that contains the text.</span></span>|
|  <span data-ttu-id="56f61-138">**Requirements**</span><span class="sxs-lookup"><span data-stu-id="56f61-138">**Requirements**</span></span>  |  <span data-ttu-id="56f61-139">いいえ</span><span class="sxs-lookup"><span data-stu-id="56f61-139">No</span></span>   |  <span data-ttu-id="56f61-p105">アドインに必要な最小の Office.js のセットおよびバージョンを指定します。これは、マニフェストの親部分の `Requirements` 要素を上書きします。</span><span class="sxs-lookup"><span data-stu-id="56f61-p105">Specifies the minimum requirement set and version of Office.js that the add-in requires. This overrides the  `Requirements` element in the parent portion of the manifest.</span></span>|
|  [<span data-ttu-id="56f61-142">Hosts</span><span class="sxs-lookup"><span data-stu-id="56f61-142">Hosts</span></span>](hosts.md)                |  <span data-ttu-id="56f61-143">はい</span><span class="sxs-lookup"><span data-stu-id="56f61-143">Yes</span></span>  |  <span data-ttu-id="56f61-144">Office アプリケーションのコレクションを指定します。</span><span class="sxs-lookup"><span data-stu-id="56f61-144">Specifies a collection of Office applications.</span></span> <span data-ttu-id="56f61-145">子の Hosts 要素は、マニフェストの親部分の Hosts 要素より優先されます。</span><span class="sxs-lookup"><span data-stu-id="56f61-145">The child Hosts element overrides the Hosts element in the parent portion of the manifest.</span></span>  |
|  [<span data-ttu-id="56f61-146">Resources</span><span class="sxs-lookup"><span data-stu-id="56f61-146">Resources</span></span>](resources.md)    |  <span data-ttu-id="56f61-147">はい</span><span class="sxs-lookup"><span data-stu-id="56f61-147">Yes</span></span>  | <span data-ttu-id="56f61-148">マニフェストの他の要素によって参照されるリソースのコレクション (文字列、URL、画像) を定義します。</span><span class="sxs-lookup"><span data-stu-id="56f61-148">Defines a collection of resources (strings, URLs, and images) that other manifest elements reference.</span></span>|
|  [<span data-ttu-id="56f61-149">EquivalentAddins</span><span class="sxs-lookup"><span data-stu-id="56f61-149">EquivalentAddins</span></span>](equivalentaddins.md)    |  <span data-ttu-id="56f61-150">いいえ</span><span class="sxs-lookup"><span data-stu-id="56f61-150">No</span></span>  | <span data-ttu-id="56f61-151">Web アドインと同等のネイティブ (COM/XLL) アドインを指定します。</span><span class="sxs-lookup"><span data-stu-id="56f61-151">Specifies the native (COM/XLL) add-ins that are equivalent to the web add-in.</span></span> <span data-ttu-id="56f61-152">同等のネイティブアドインがインストールされている場合、web アドインはアクティブ化されません。</span><span class="sxs-lookup"><span data-stu-id="56f61-152">The web add-in is not activated if an equivalent native add-in is installed.</span></span>|
|  <span data-ttu-id="56f61-153">**VersionOverrides**</span><span class="sxs-lookup"><span data-stu-id="56f61-153">**VersionOverrides**</span></span>    |  <span data-ttu-id="56f61-154">いいえ</span><span class="sxs-lookup"><span data-stu-id="56f61-154">No</span></span>  | <span data-ttu-id="56f61-p108">より新しいスキーマ バージョンでアドイン コマンドを定義します。詳細については、「[複数のバージョンを実装する](#implementing-multiple-versions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="56f61-p108">Defines add-in commands under a newer schema version. See [Implementing multiple versions](#implementing-multiple-versions) for details.</span></span> |
|  [<span data-ttu-id="56f61-157">WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="56f61-157">WebApplicationInfo</span></span>](webapplicationinfo.md)    |  <span data-ttu-id="56f61-158">いいえ</span><span class="sxs-lookup"><span data-stu-id="56f61-158">No</span></span>  | <span data-ttu-id="56f61-159">Azure Active Directory v2.0 など、セキュリティで保護されたトークン発行者によるアドインの登録に関する詳細を指定します。</span><span class="sxs-lookup"><span data-stu-id="56f61-159">Specifies details about the add-in's registration with secure token issuers, such as Azure Active Directory V2.0.</span></span> |
|  [<span data-ttu-id="56f61-160">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="56f61-160">ExtendedPermissions</span></span>](extendedpermissions.md) |  <span data-ttu-id="56f61-161">いいえ</span><span class="sxs-lookup"><span data-stu-id="56f61-161">No</span></span>  |  <span data-ttu-id="56f61-162">拡張アクセス許可のコレクションを指定します。</span><span class="sxs-lookup"><span data-stu-id="56f61-162">Specifies a collection of extended permissions.</span></span><br><br><span data-ttu-id="56f61-163">**重要**: [Office. appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-) API は現在プレビュー段階のため、この要素を使用するアドインは、 `ExtendedPermissions` appsource に発行することも、一元展開によって展開することもできません。</span><span class="sxs-lookup"><span data-stu-id="56f61-163">**Important**: Because the [Office.Body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-) API is currently in preview, add-ins that use the `ExtendedPermissions` element can't be published to AppSource or deployed via centralized deployment.</span></span> |

### <a name="versionoverrides-example"></a><span data-ttu-id="56f61-164">VersionOverrides の例</span><span class="sxs-lookup"><span data-stu-id="56f61-164">VersionOverrides example</span></span>

<span data-ttu-id="56f61-165">通常、必須では `<VersionOverrides>` ありませんが通常使用される子要素を含む一般的な要素の例を次に示します。</span><span class="sxs-lookup"><span data-stu-id="56f61-165">The following is an example of a typical `<VersionOverrides>` element, including some child elements that are not required but are typically used.</span></span>

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

## <a name="implementing-multiple-versions"></a><span data-ttu-id="56f61-166">複数のバージョンを実装する</span><span class="sxs-lookup"><span data-stu-id="56f61-166">Implementing multiple versions</span></span>

<span data-ttu-id="56f61-p109">1 つのマニフェストで、複数のバージョンの `VersionOverrides` 要素を実装することで、異なるバージョンの VersionOverrides スキーマをサポートできます。これは、新しいスキーマの新機能をオプションでサポートしながら、新機能をサポートしていない古いクライアントもサポートすることで実現できます。</span><span class="sxs-lookup"><span data-stu-id="56f61-p109">A manifest can implement multiple versions of the `VersionOverrides` element which support different versions of the VersionOverrides schema. This can be done to optionally support new features in a newer schema while still supporting older clients that do not support the new features.</span></span>

<span data-ttu-id="56f61-169">複数のバージョンを実装するために、新しいバージョンの `VersionOverrides` 要素は、古いバージョンの `VersionOverrides` 要素の子にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="56f61-169">In order to implement multiple versions, the `VersionOverrides` element for the newer version must be a child of the `VersionOverrides` element for the older version.</span></span> <span data-ttu-id="56f61-170">子の `VersionOverrides` 要素は、どの値も親から継承しません。</span><span class="sxs-lookup"><span data-stu-id="56f61-170">The child `VersionOverrides` element doesn't inherit any values from the parent.</span></span>

<span data-ttu-id="56f61-171">VersionOverrides v1.0 と v1.1 の両方のスキーマを実装するためのマニフェストは、次に示す例のようになります。</span><span class="sxs-lookup"><span data-stu-id="56f61-171">To implement both the VersionOverrides v1.0 and v1.1 schema, the manifest would look similar to the following example:</span></span>

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
