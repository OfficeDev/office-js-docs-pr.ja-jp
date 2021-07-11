---
title: マニフェスト ファイルの VersionOverrides 要素
description: アドイン マニフェスト (XML) ファイルOffice VersionOverrides 要素のリファレンス ドキュメント。
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 787ba8e7d90900cc72d6c5e9370d68ced0faee2f
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348658"
---
# <a name="versionoverrides-element"></a><span data-ttu-id="70bd7-103">VersionOverrides 要素</span><span class="sxs-lookup"><span data-stu-id="70bd7-103">VersionOverrides element</span></span>

<span data-ttu-id="70bd7-p101">アドインによって実装されたアドイン コマンドに関する情報を格納するルート要素です。**VersionOverrides** は、マニフェスト内の [OfficeApp](officeapp.md) 要素の子要素です。この要素は、マニフェスト スキーマ v1.1 以降でサポートされていますが、VersionOverrides v1.0 または v1.1 スキーマで定義されています。</span><span class="sxs-lookup"><span data-stu-id="70bd7-p101">The root element that contains information for the add-in commands implemented by the add-in. **VersionOverrides** is a child element of the [OfficeApp](officeapp.md) element in the manifest. This element is supported in manifest schema v1.1 and later but is defined in the VersionOverrides v1.0 or v1.1 schema.</span></span>

## <a name="attributes"></a><span data-ttu-id="70bd7-107">属性</span><span class="sxs-lookup"><span data-stu-id="70bd7-107">Attributes</span></span>

|  <span data-ttu-id="70bd7-108">属性</span><span class="sxs-lookup"><span data-stu-id="70bd7-108">Attribute</span></span>  |  <span data-ttu-id="70bd7-109">必須</span><span class="sxs-lookup"><span data-stu-id="70bd7-109">Required</span></span>  |  <span data-ttu-id="70bd7-110">説明</span><span class="sxs-lookup"><span data-stu-id="70bd7-110">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="70bd7-111">**xmlns**</span><span class="sxs-lookup"><span data-stu-id="70bd7-111">**xmlns**</span></span>       |  <span data-ttu-id="70bd7-112">はい</span><span class="sxs-lookup"><span data-stu-id="70bd7-112">Yes</span></span>  |  <span data-ttu-id="70bd7-113">VersionOverrides スキーマ名前空間。</span><span class="sxs-lookup"><span data-stu-id="70bd7-113">The VersionOverrides schema namespace.</span></span> <span data-ttu-id="70bd7-114">許可される値は、この要素の `<VersionOverrides>` **xsi:type** 値と親要素の **xsi:type** 値によって異 `<OfficeApp>` なります。</span><span class="sxs-lookup"><span data-stu-id="70bd7-114">The allowed values vary depending on  this `<VersionOverrides>` element's **xsi:type** value and the **xsi:type** value of the parent `<OfficeApp>` element.</span></span> <span data-ttu-id="70bd7-115">以下の [名前空間の値を参照](#namespace-values) してください。</span><span class="sxs-lookup"><span data-stu-id="70bd7-115">See [Namespace values](#namespace-values) below.</span></span>|
|  <span data-ttu-id="70bd7-116">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="70bd7-116">**xsi:type**</span></span>  |  <span data-ttu-id="70bd7-117">はい</span><span class="sxs-lookup"><span data-stu-id="70bd7-117">Yes</span></span>  | <span data-ttu-id="70bd7-p103">スキーマのバージョン。現時点では、`VersionOverridesV1_0` および `VersionOverridesV1_1` のみが有効な値になります。</span><span class="sxs-lookup"><span data-stu-id="70bd7-p103">The schema version. At this time, the only valid values are `VersionOverridesV1_0` and `VersionOverridesV1_1`.</span></span> |

### <a name="namespace-values"></a><span data-ttu-id="70bd7-120">名前空間の値</span><span class="sxs-lookup"><span data-stu-id="70bd7-120">Namespace values</span></span>

<span data-ttu-id="70bd7-121">親要素の **xsi:type** 値に応じて **、xmlns** 値の必要な値を次に示 `<OfficeApp>` します。</span><span class="sxs-lookup"><span data-stu-id="70bd7-121">The following lists the required value of the **xmlns** value depending on the **xsi:type** value of the parent `<OfficeApp>` element.</span></span>

- <span data-ttu-id="70bd7-122">**TaskPaneApp は** VersionOverrides のバージョン 1.0 のみをサポートし **、xmlns は** `http://schemas.microsoft.com/office/taskpaneappversionoverrides` .</span><span class="sxs-lookup"><span data-stu-id="70bd7-122">**TaskPaneApp** supports only version 1.0 of VersionOverrides, and the **xmlns** should be `http://schemas.microsoft.com/office/taskpaneappversionoverrides`.</span></span>
- <span data-ttu-id="70bd7-123">**ContentApp** は VersionOverrides のバージョン 1.0 のみをサポートし **、xmlns は** `http://schemas.microsoft.com/office/contentappversionoverrides` .</span><span class="sxs-lookup"><span data-stu-id="70bd7-123">**ContentApp** supports only version 1.0 of VersionOverrides, and the **xmlns** should be `http://schemas.microsoft.com/office/contentappversionoverrides`.</span></span>
- <span data-ttu-id="70bd7-124">**MailApp** は VersionOverrides のバージョン 1.0 と 1.1 をサポートしています。 **したがって、xmlns** の値は、この要素の `<VersionOverrides>` **xsi:type** 値によって異なります。</span><span class="sxs-lookup"><span data-stu-id="70bd7-124">**MailApp** supports versions 1.0 and 1.1 of VersionOverrides, so the value of **xmlns** varies depending on this `<VersionOverrides>` element's **xsi:type** value:</span></span>
    - <span data-ttu-id="70bd7-125">**xsi:type がである** 場合 `VersionOverridesV1_0` は **、xmlns を** 指定する必要があります `http://schemas.microsoft.com/office/mailappversionoverrides` 。</span><span class="sxs-lookup"><span data-stu-id="70bd7-125">When **xsi:type** is `VersionOverridesV1_0`, **xmlns** must be `http://schemas.microsoft.com/office/mailappversionoverrides`.</span></span>
    - <span data-ttu-id="70bd7-126">**xsi:type がである** 場合 `VersionOverridesV1_1` は **、xmlns を** 指定する必要があります `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` 。</span><span class="sxs-lookup"><span data-stu-id="70bd7-126">When **xsi:type** is `VersionOverridesV1_1`, **xmlns** must be `http://schemas.microsoft.com/office/mailappversionoverrides/1.1`.</span></span>

> [!NOTE]
> <span data-ttu-id="70bd7-127">現在のところ、Outlook 2016以降は VersionOverrides v1.1 スキーマと型をサポート `VersionOverridesV1_1` しています。</span><span class="sxs-lookup"><span data-stu-id="70bd7-127">Currently only Outlook 2016 or later supports the VersionOverrides v1.1 schema and the `VersionOverridesV1_1` type.</span></span>

## <a name="child-elements"></a><span data-ttu-id="70bd7-128">子要素</span><span class="sxs-lookup"><span data-stu-id="70bd7-128">Child elements</span></span>

|  <span data-ttu-id="70bd7-129">要素</span><span class="sxs-lookup"><span data-stu-id="70bd7-129">Element</span></span> |  <span data-ttu-id="70bd7-130">必須</span><span class="sxs-lookup"><span data-stu-id="70bd7-130">Required</span></span>  |  <span data-ttu-id="70bd7-131">説明</span><span class="sxs-lookup"><span data-stu-id="70bd7-131">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="70bd7-132">**説明**</span><span class="sxs-lookup"><span data-stu-id="70bd7-132">**Description**</span></span>    |  <span data-ttu-id="70bd7-133">いいえ</span><span class="sxs-lookup"><span data-stu-id="70bd7-133">No</span></span>   |  <span data-ttu-id="70bd7-134">アドインについての説明。</span><span class="sxs-lookup"><span data-stu-id="70bd7-134">Describes the add-in.</span></span> <span data-ttu-id="70bd7-135">これは、マニフェスト内の任意の親部分の `Description` 要素を上書きします。</span><span class="sxs-lookup"><span data-stu-id="70bd7-135">This overrides the `Description` element in any parent portion of the manifest.</span></span> <span data-ttu-id="70bd7-136">説明のテキストは、**Resources** 要素の [LongString](resources.md) 要素の子要素に含まれています。</span><span class="sxs-lookup"><span data-stu-id="70bd7-136">The text of the description is contained in a child element of the **LongString** element contained in the [Resources](resources.md) element.</span></span> <span data-ttu-id="70bd7-137">`resid`Description 要素の **属性** は 32 文字以内で、テキストを含む要素の属性の値 `id` `String` に設定されます。</span><span class="sxs-lookup"><span data-stu-id="70bd7-137">The `resid` attribute of the **Description** element can be no more than 32 characters and is set to the value of the `id` attribute of the `String` element that contains the text.</span></span>|
|  <span data-ttu-id="70bd7-138">**Requirements**</span><span class="sxs-lookup"><span data-stu-id="70bd7-138">**Requirements**</span></span>  |  <span data-ttu-id="70bd7-139">いいえ</span><span class="sxs-lookup"><span data-stu-id="70bd7-139">No</span></span>   |  <span data-ttu-id="70bd7-p105">アドインに必要な最小の Office.js のセットおよびバージョンを指定します。これは、マニフェストの親部分の `Requirements` 要素を上書きします。</span><span class="sxs-lookup"><span data-stu-id="70bd7-p105">Specifies the minimum requirement set and version of Office.js that the add-in requires. This overrides the  `Requirements` element in the parent portion of the manifest.</span></span>|
|  [<span data-ttu-id="70bd7-142">Hosts</span><span class="sxs-lookup"><span data-stu-id="70bd7-142">Hosts</span></span>](hosts.md)                |  <span data-ttu-id="70bd7-143">はい</span><span class="sxs-lookup"><span data-stu-id="70bd7-143">Yes</span></span>  |  <span data-ttu-id="70bd7-144">アプリケーションのコレクションをOfficeします。</span><span class="sxs-lookup"><span data-stu-id="70bd7-144">Specifies a collection of Office applications.</span></span> <span data-ttu-id="70bd7-145">子 Hosts 要素は、マニフェストの親部分にある Hosts 要素をオーバーライドします。</span><span class="sxs-lookup"><span data-stu-id="70bd7-145">The child Hosts element overrides the Hosts element in the parent portion of the manifest.</span></span>  |
|  [<span data-ttu-id="70bd7-146">Resources</span><span class="sxs-lookup"><span data-stu-id="70bd7-146">Resources</span></span>](resources.md)    |  <span data-ttu-id="70bd7-147">はい</span><span class="sxs-lookup"><span data-stu-id="70bd7-147">Yes</span></span>  | <span data-ttu-id="70bd7-148">マニフェストの他の要素によって参照されるリソースのコレクション (文字列、URL、画像) を定義します。</span><span class="sxs-lookup"><span data-stu-id="70bd7-148">Defines a collection of resources (strings, URLs, and images) that other manifest elements reference.</span></span>|
|  [<span data-ttu-id="70bd7-149">EquivalentAddins</span><span class="sxs-lookup"><span data-stu-id="70bd7-149">EquivalentAddins</span></span>](equivalentaddins.md)    |  <span data-ttu-id="70bd7-150">いいえ</span><span class="sxs-lookup"><span data-stu-id="70bd7-150">No</span></span>  | <span data-ttu-id="70bd7-151">Web アドインと同等のネイティブ (COM/XLL) アドインを指定します。</span><span class="sxs-lookup"><span data-stu-id="70bd7-151">Specifies the native (COM/XLL) add-ins that are equivalent to the web add-in.</span></span> <span data-ttu-id="70bd7-152">同等のネイティブ アドインがインストールされている場合、Web アドインはアクティブ化されません。</span><span class="sxs-lookup"><span data-stu-id="70bd7-152">The web add-in is not activated if an equivalent native add-in is installed.</span></span>|
|  <span data-ttu-id="70bd7-153">**VersionOverrides**</span><span class="sxs-lookup"><span data-stu-id="70bd7-153">**VersionOverrides**</span></span>    |  <span data-ttu-id="70bd7-154">いいえ</span><span class="sxs-lookup"><span data-stu-id="70bd7-154">No</span></span>  | <span data-ttu-id="70bd7-p108">より新しいスキーマ バージョンでアドイン コマンドを定義します。詳細については、「[複数のバージョンを実装する](#implementing-multiple-versions)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="70bd7-p108">Defines add-in commands under a newer schema version. See [Implementing multiple versions](#implementing-multiple-versions) for details.</span></span> |
|  [<span data-ttu-id="70bd7-157">WebApplicationInfo</span><span class="sxs-lookup"><span data-stu-id="70bd7-157">WebApplicationInfo</span></span>](webapplicationinfo.md)    |  <span data-ttu-id="70bd7-158">いいえ</span><span class="sxs-lookup"><span data-stu-id="70bd7-158">No</span></span>  | <span data-ttu-id="70bd7-159">セキュリティで保護されたトークン発行者とのアドインの登録に関する詳細 (V2.0 などAzure Active Directory指定します。</span><span class="sxs-lookup"><span data-stu-id="70bd7-159">Specifies details about the add-in's registration with secure token issuers, such as Azure Active Directory V2.0.</span></span> |
|  [<span data-ttu-id="70bd7-160">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="70bd7-160">ExtendedPermissions</span></span>](extendedpermissions.md) |  <span data-ttu-id="70bd7-161">いいえ</span><span class="sxs-lookup"><span data-stu-id="70bd7-161">No</span></span>  |  <span data-ttu-id="70bd7-162">拡張アクセス許可のコレクションを指定します。</span><span class="sxs-lookup"><span data-stu-id="70bd7-162">Specifies a collection of extended permissions.</span></span> |

### <a name="versionoverrides-example"></a><span data-ttu-id="70bd7-163">VersionOverrides の例</span><span class="sxs-lookup"><span data-stu-id="70bd7-163">VersionOverrides example</span></span>

<span data-ttu-id="70bd7-164">次に示すのは、一般的な要素の例です。一部の子要素は必須ではなく、通常 `<VersionOverrides>` は使用されます。</span><span class="sxs-lookup"><span data-stu-id="70bd7-164">The following is an example of a typical `<VersionOverrides>` element, including some child elements that are not required but are typically used.</span></span>

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

## <a name="implementing-multiple-versions"></a><span data-ttu-id="70bd7-165">複数のバージョンを実装する</span><span class="sxs-lookup"><span data-stu-id="70bd7-165">Implementing multiple versions</span></span>

<span data-ttu-id="70bd7-p109">1 つのマニフェストで、複数のバージョンの `VersionOverrides` 要素を実装することで、異なるバージョンの VersionOverrides スキーマをサポートできます。これは、新しいスキーマの新機能をオプションでサポートしながら、新機能をサポートしていない古いクライアントもサポートすることで実現できます。</span><span class="sxs-lookup"><span data-stu-id="70bd7-p109">A manifest can implement multiple versions of the `VersionOverrides` element which support different versions of the VersionOverrides schema. This can be done to optionally support new features in a newer schema while still supporting older clients that do not support the new features.</span></span>

<span data-ttu-id="70bd7-168">複数のバージョンを実装するために、新しいバージョンの `VersionOverrides` 要素は、古いバージョンの `VersionOverrides` 要素の子にする必要があります。</span><span class="sxs-lookup"><span data-stu-id="70bd7-168">In order to implement multiple versions, the `VersionOverrides` element for the newer version must be a child of the `VersionOverrides` element for the older version.</span></span> <span data-ttu-id="70bd7-169">子の `VersionOverrides` 要素は、どの値も親から継承しません。</span><span class="sxs-lookup"><span data-stu-id="70bd7-169">The child `VersionOverrides` element doesn't inherit any values from the parent.</span></span>

<span data-ttu-id="70bd7-170">VersionOverrides v1.0 スキーマと v1.1 スキーマの両方を実装するには、マニフェストは次の例のようになります。</span><span class="sxs-lookup"><span data-stu-id="70bd7-170">To implement both the VersionOverrides v1.0 and v1.1 schema, the manifest would look similar to the following example.</span></span>

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
