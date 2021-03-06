---
title: マニフェスト ファイルの Override 要素
description: Override 要素を使用すると、指定した条件に応じて設定の値を指定できます。
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: d2146cc1f44e829bc78076c8093b2ebf791dc722
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505340"
---
# <a name="override-element"></a><span data-ttu-id="5523c-103">Override 要素</span><span class="sxs-lookup"><span data-stu-id="5523c-103">Override element</span></span>

<span data-ttu-id="5523c-104">指定した条件に応じてマニフェスト設定の値を上書きする方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="5523c-104">Provides a way to override the value of a manifest setting depending on a specified condition.</span></span> <span data-ttu-id="5523c-105">条件には次の 2 種類があります。</span><span class="sxs-lookup"><span data-stu-id="5523c-105">There are two kinds of conditions:</span></span>

- <span data-ttu-id="5523c-106">既定Office異なるロケールを指定します。</span><span class="sxs-lookup"><span data-stu-id="5523c-106">An Office locale that is different from the default.</span></span>
- <span data-ttu-id="5523c-107">既定のパターンとは異なる要件セットのサポートのパターン。</span><span class="sxs-lookup"><span data-stu-id="5523c-107">A pattern of requirement set support that is different from the default pattern.</span></span>

<span data-ttu-id="5523c-108">要素には、LocaleTokenOverride と呼ばれるロケールオーバーライド用の要素と `<Override>` **、RequirementTokenOverride** と呼ばれる要件セットのオーバーライド用の 2 種類があります。</span><span class="sxs-lookup"><span data-stu-id="5523c-108">There are two types of `<Override>` elements, one is for locale overrides, called **LocaleTokenOverride**, and the other for requirement set overrides, called **RequirementTokenOverride**.</span></span> <span data-ttu-id="5523c-109">ただし、要素 `type` のパラメーター `<Override>` はありません。</span><span class="sxs-lookup"><span data-stu-id="5523c-109">But there is no `type` parameter for the `<Override>` element.</span></span> <span data-ttu-id="5523c-110">違いは、親要素と親要素の型によって決まります。</span><span class="sxs-lookup"><span data-stu-id="5523c-110">The difference is determined by the parent element and the parent element's type.</span></span> <span data-ttu-id="5523c-111">要素 `<Override>` の内部にある要素は `<Token>` `xsi:type` `RequirementToken` **、RequirementTokenOverride 型である必要があります**。</span><span class="sxs-lookup"><span data-stu-id="5523c-111">An `<Override>` element that is inside of a `<Token>` element whose `xsi:type` is `RequirementToken`, must be of type **RequirementTokenOverride**.</span></span> <span data-ttu-id="5523c-112">他 `<Override>` の親要素内の要素、または型の要素内の要素は `<Override>` `LocaleToken` **、LocaleTokenOverride 型である必要があります**。</span><span class="sxs-lookup"><span data-stu-id="5523c-112">An `<Override>` element inside any other parent element, or inside an `<Override>` element of type `LocaleToken`, must be of type **LocaleTokenOverride**.</span></span> <span data-ttu-id="5523c-113">各種類については、以下の各セクションで説明します。</span><span class="sxs-lookup"><span data-stu-id="5523c-113">Each type is described in separate sections below.</span></span> <span data-ttu-id="5523c-114">要素の子である場合のこの要素の使用の詳細については、「マニフェストの拡張オーバーライドを処理する」 `<Token>` [を参照してください](../../develop/extended-overrides.md)。</span><span class="sxs-lookup"><span data-stu-id="5523c-114">For more information about the use of this element when it is a child of a `<Token>` element, see [Work with extended overrides of the manifest](../../develop/extended-overrides.md).</span></span>

## <a name="override-element-of-type-localetokenoverride"></a><span data-ttu-id="5523c-115">LocaleTokenOverride 型のオーバーライド要素</span><span class="sxs-lookup"><span data-stu-id="5523c-115">Override element of type LocaleTokenOverride</span></span>

<span data-ttu-id="5523c-116">要素 `<Override>` は条件付きを表し、"If .." として読み取り可能です。その後 ..."。ステートメント。</span><span class="sxs-lookup"><span data-stu-id="5523c-116">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="5523c-117">要素が `<Override>` **LocaleTokenOverride** 型の場合、属性は条件であり、その `Locale` `Value` 結果属性になります。</span><span class="sxs-lookup"><span data-stu-id="5523c-117">If the `<Override>` element is of type **LocaleTokenOverride**, then the `Locale` attribute is the condition, and the `Value` attribute is the consequent.</span></span> <span data-ttu-id="5523c-118">たとえば、次の例は、「Officeロケール設定が fr-fr の場合、表示名は 'Lecteur vidéo'です。</span><span class="sxs-lookup"><span data-stu-id="5523c-118">For example, the following is read "If the Office locale setting is fr-fr, then the display name is 'Lecteur vidéo'."</span></span>

```xml
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

<span data-ttu-id="5523c-119">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="5523c-119">**Add-in type:** Content, Task pane, Mail</span></span>

### <a name="syntax"></a><span data-ttu-id="5523c-120">構文</span><span class="sxs-lookup"><span data-stu-id="5523c-120">Syntax</span></span>

```XML
<Override Locale="string" Value="string"></Override>
```

### <a name="contained-in"></a><span data-ttu-id="5523c-121">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="5523c-121">Contained in</span></span>

|<span data-ttu-id="5523c-122">要素</span><span class="sxs-lookup"><span data-stu-id="5523c-122">Element</span></span>|
|:-----|
|[<span data-ttu-id="5523c-123">CitationText</span><span class="sxs-lookup"><span data-stu-id="5523c-123">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="5523c-124">説明</span><span class="sxs-lookup"><span data-stu-id="5523c-124">Description</span></span>](description.md)|
|[<span data-ttu-id="5523c-125">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="5523c-125">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="5523c-126">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="5523c-126">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="5523c-127">DisplayName</span><span class="sxs-lookup"><span data-stu-id="5523c-127">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="5523c-128">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="5523c-128">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="5523c-129">IconUrl</span><span class="sxs-lookup"><span data-stu-id="5523c-129">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="5523c-130">QueryUri</span><span class="sxs-lookup"><span data-stu-id="5523c-130">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="5523c-131">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="5523c-131">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="5523c-132">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="5523c-132">SupportUrl</span></span>](supporturl.md)|
|[<span data-ttu-id="5523c-133">トークン</span><span class="sxs-lookup"><span data-stu-id="5523c-133">Token</span></span>](token.md)|

### <a name="attributes"></a><span data-ttu-id="5523c-134">属性</span><span class="sxs-lookup"><span data-stu-id="5523c-134">Attributes</span></span>

|<span data-ttu-id="5523c-135">属性</span><span class="sxs-lookup"><span data-stu-id="5523c-135">Attribute</span></span>|<span data-ttu-id="5523c-136">型</span><span class="sxs-lookup"><span data-stu-id="5523c-136">Type</span></span>|<span data-ttu-id="5523c-137">必須</span><span class="sxs-lookup"><span data-stu-id="5523c-137">Required</span></span>|<span data-ttu-id="5523c-138">説明</span><span class="sxs-lookup"><span data-stu-id="5523c-138">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="5523c-139">Locale</span><span class="sxs-lookup"><span data-stu-id="5523c-139">Locale</span></span>|<span data-ttu-id="5523c-140">string</span><span class="sxs-lookup"><span data-stu-id="5523c-140">string</span></span>|<span data-ttu-id="5523c-141">必須</span><span class="sxs-lookup"><span data-stu-id="5523c-141">required</span></span>|<span data-ttu-id="5523c-142">`"en-US"` などの BCP 47 言語タグの書式で、この上書きのロケールのカルチャ名を指定します。</span><span class="sxs-lookup"><span data-stu-id="5523c-142">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="5523c-143">Value</span><span class="sxs-lookup"><span data-stu-id="5523c-143">Value</span></span>|<span data-ttu-id="5523c-144">string</span><span class="sxs-lookup"><span data-stu-id="5523c-144">string</span></span>|<span data-ttu-id="5523c-145">必須</span><span class="sxs-lookup"><span data-stu-id="5523c-145">required</span></span>|<span data-ttu-id="5523c-146">指定のロケールに対して表される設定の値を指定します。</span><span class="sxs-lookup"><span data-stu-id="5523c-146">Specifies value of the setting expressed for the specified locale.</span></span>|

### <a name="examples"></a><span data-ttu-id="5523c-147">例</span><span class="sxs-lookup"><span data-stu-id="5523c-147">Examples</span></span>

```xml
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

```xml
<bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/icon_default.png">
    <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp16-icon_default.png" />
</bt:Image>
```

```xml
  <ExtendedOverrides Url="http://contoso.com/addinmetadata/${token.locale}/extended-manifest-overrides.json">
    <Tokens>
      <Token Name="locale" DefaultValue="en-us" xsi:type="LocaleToken">
        <Override Locale="es-*" Value="es-es" />
        <Override Locale="es-mx" Value="es-mx" />
        <Override Locale="fr-*" Value="fr-fr" />
        <Override Locale="ja-jp" Value="ja-jp" />
      </Token>
    <Tokens>
  </ExtendedOverrides>
```

### <a name="see-also"></a><span data-ttu-id="5523c-148">関連項目</span><span class="sxs-lookup"><span data-stu-id="5523c-148">See also</span></span>

- [<span data-ttu-id="5523c-149">Office アドインのローカライズ</span><span class="sxs-lookup"><span data-stu-id="5523c-149">Localization for Office Add-ins</span></span>](../../develop/localization.md)
- [<span data-ttu-id="5523c-150">SharePoint のキーボード ショートカット</span><span class="sxs-lookup"><span data-stu-id="5523c-150">Keyboard shortcuts</span></span>](../../design/keyboard-shortcuts.md)

## <a name="override-element-of-type-requirementtokenoverride"></a><span data-ttu-id="5523c-151">RequirementTokenOverride 型の Override 要素</span><span class="sxs-lookup"><span data-stu-id="5523c-151">Override element of type RequirementTokenOverride</span></span>

<span data-ttu-id="5523c-152">要素 `<Override>` は条件付きを表し、"If .." として読み取り可能です。その後 ..."。ステートメント。</span><span class="sxs-lookup"><span data-stu-id="5523c-152">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="5523c-153">要素が `<Override>` **RequirementTokenOverride** 型の場合、子要素は条件を表し、属性 `<Requirements>` `Value` はその結果です。</span><span class="sxs-lookup"><span data-stu-id="5523c-153">If the `<Override>` element is of type **RequirementTokenOverride**, then the child `<Requirements>` element expresses the condition, and the `Value` attribute is the consequent.</span></span> <span data-ttu-id="5523c-154">たとえば、次の 1 つ目は、「現在のプラットフォームが FeatureOne バージョン 1.7 をサポートしている場合は、(既定の文字列 'upgrade' ではなく) 祖父母の URL のトークンの代わりに文字列 `<Override>` 'oldAddinVersion' を使用します。 `${token.requirements}` `<ExtendedOverrides>`</span><span class="sxs-lookup"><span data-stu-id="5523c-154">For example, the first `<Override>` in the following is read "If the current platform supports FeatureOne version 1.7, then use string 'oldAddinVersion' in place of the `${token.requirements}` token in the URL of the grandparent `<ExtendedOverrides>` (instead of the default string 'upgrade')."</span></span>

```xml
<ExtendedOverrides Url="http://contoso.com/addinmetadata/${token.requirements}/extended-manifest-overrides.json">
    <Tokens>
        <Token Name="requirements" DefaultValue="upgrade" xsi:type="RequirementsToken">
            <Override Value="oldAddinVersion">
                <Requirements>
                    <Sets>
                        <Set Name="FeatureOne" MinVersion="1.7" />
                    </Sets>
                </Requirements>
            </Override>
            <Override Value="currentAddinVersion">
                <Requirements>
                    <Sets>
                        <Set Name="FeatureOne" MinVersion="1.8" />
                    </Sets>
                    <Methods>
                        <Method Name="MethodThree" />
                    </Methods>
                </Requirements>
            </Override>
        </Token>
    </Tokens>
</ExtendedOverrides>
```

<span data-ttu-id="5523c-155">**アドインの種類:** 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="5523c-155">**Add-in type:** Task pane</span></span>

### <a name="syntax"></a><span data-ttu-id="5523c-156">構文</span><span class="sxs-lookup"><span data-stu-id="5523c-156">Syntax</span></span>

```XML
<Override Value="string" />
```

### <a name="contained-in"></a><span data-ttu-id="5523c-157">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="5523c-157">Contained in</span></span>

|<span data-ttu-id="5523c-158">要素</span><span class="sxs-lookup"><span data-stu-id="5523c-158">Element</span></span>|
|:-----|
|[<span data-ttu-id="5523c-159">トークン</span><span class="sxs-lookup"><span data-stu-id="5523c-159">Token</span></span>](token.md)|

### <a name="must-contain"></a><span data-ttu-id="5523c-160">含める必要があるもの</span><span class="sxs-lookup"><span data-stu-id="5523c-160">Must contain</span></span>

|<span data-ttu-id="5523c-161">要素</span><span class="sxs-lookup"><span data-stu-id="5523c-161">Element</span></span>|<span data-ttu-id="5523c-162">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="5523c-162">Content</span></span>|<span data-ttu-id="5523c-163">メール</span><span class="sxs-lookup"><span data-stu-id="5523c-163">Mail</span></span>|<span data-ttu-id="5523c-164">TaskPane</span><span class="sxs-lookup"><span data-stu-id="5523c-164">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="5523c-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="5523c-165">Requirements</span></span>](requirements.md)|||<span data-ttu-id="5523c-166">x</span><span class="sxs-lookup"><span data-stu-id="5523c-166">x</span></span>|

### <a name="attributes"></a><span data-ttu-id="5523c-167">属性</span><span class="sxs-lookup"><span data-stu-id="5523c-167">Attributes</span></span>

|<span data-ttu-id="5523c-168">属性</span><span class="sxs-lookup"><span data-stu-id="5523c-168">Attribute</span></span>|<span data-ttu-id="5523c-169">型</span><span class="sxs-lookup"><span data-stu-id="5523c-169">Type</span></span>|<span data-ttu-id="5523c-170">必須</span><span class="sxs-lookup"><span data-stu-id="5523c-170">Required</span></span>|<span data-ttu-id="5523c-171">説明</span><span class="sxs-lookup"><span data-stu-id="5523c-171">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="5523c-172">値</span><span class="sxs-lookup"><span data-stu-id="5523c-172">Value</span></span>|<span data-ttu-id="5523c-173">string</span><span class="sxs-lookup"><span data-stu-id="5523c-173">string</span></span>|<span data-ttu-id="5523c-174">必須</span><span class="sxs-lookup"><span data-stu-id="5523c-174">required</span></span>|<span data-ttu-id="5523c-175">条件が満たされた場合の祖父母トークンの値。</span><span class="sxs-lookup"><span data-stu-id="5523c-175">Value of the grandparent token when the condition is satisfied.</span></span>|

### <a name="example"></a><span data-ttu-id="5523c-176">例</span><span class="sxs-lookup"><span data-stu-id="5523c-176">Example</span></span>

```xml
<ExtendedOverrides Url="http://contoso.com/addinmetadata/${token.requirements}/extended-manifest-overrides.json">
    <Token Name="requirements" DefaultValue="upgrade" xsi:type="RequirementsToken">
        <Override Value="very-old">
            <Requirements>
                <Sets>
                    <Set Name="FeatureOne" MinVersion="1.5" />
                    <Set Name="FeatureTwo" MinVersion="1.1" />
                </Sets>
            </Requirements>
        </Override>
        <Override Value="old">
            <Requirements>
                <Sets>
                    <Set Name="FeatureOne" MinVersion="1.7" />
                    <Set Name="FeatureTwo" MinVersion="1.2" />
                </Sets>
            </Requirements>
        </Override>
        <Override Value="current">
            <Requirements>
                <Sets>
                    <Set Name="FeatureOne" MinVersion="1.8" />
                    <Set Name="FeatureTwo" MinVersion="1.3" />
                </Sets>
                <Methods>
                    <Method Name="MethodThree" />
                </Methods>
            </Requirements>
        </Override>
    </Token>
</ExtendedOverrides>
```

### <a name="see-also"></a><span data-ttu-id="5523c-177">関連項目</span><span class="sxs-lookup"><span data-stu-id="5523c-177">See also</span></span>

- [<span data-ttu-id="5523c-178">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="5523c-178">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="5523c-179">マニフェストで Requirements 要素を設定する</span><span class="sxs-lookup"><span data-stu-id="5523c-179">Set the Requirements element in the manifest</span></span>](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)
- [<span data-ttu-id="5523c-180">SharePoint のキーボード ショートカット</span><span class="sxs-lookup"><span data-stu-id="5523c-180">Keyboard shortcuts</span></span>](../../design/keyboard-shortcuts.md)
