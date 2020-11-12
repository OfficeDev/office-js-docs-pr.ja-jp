---
title: マニフェスト ファイルの Override 要素
description: Override 要素を使用すると、指定した条件に応じて設定の値を指定できます。
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 2c66503f9f95155a096b1b6fb23332eed8422da6
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996313"
---
# <a name="override-element"></a><span data-ttu-id="da2ee-103">Override 要素</span><span class="sxs-lookup"><span data-stu-id="da2ee-103">Override element</span></span>

<span data-ttu-id="da2ee-104">指定した条件に応じて、マニフェストの設定値を上書きする方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="da2ee-104">Provides a way to override the value of a manifest setting depending on a specified condition.</span></span> <span data-ttu-id="da2ee-105">条件には、次の2種類があります。</span><span class="sxs-lookup"><span data-stu-id="da2ee-105">There are two kinds of conditions:</span></span>

- <span data-ttu-id="da2ee-106">既定とは異なる Office ロケール。</span><span class="sxs-lookup"><span data-stu-id="da2ee-106">An Office locale that is different from the default.</span></span>
- <span data-ttu-id="da2ee-107">既定のパターンとは異なる、要件セットサポートのパターン。</span><span class="sxs-lookup"><span data-stu-id="da2ee-107">A pattern of requirement set support that is different from the default pattern.</span></span>

<span data-ttu-id="da2ee-108">要素には2つの種類があり `<Override>` ます。1つは **LocaleTokenOverride** と呼ばれるロケールの上書き用で、もう1つは要件セットのオーバーライド ( **RequirementTokenOverride** と呼ばれる) です。</span><span class="sxs-lookup"><span data-stu-id="da2ee-108">There are two types of `<Override>` elements, one is for locale overrides, called **LocaleTokenOverride** , and the other for requirement set overrides, called **RequirementTokenOverride**.</span></span> <span data-ttu-id="da2ee-109">ただし `type` 、要素のパラメーターはありません `<Override>` 。</span><span class="sxs-lookup"><span data-stu-id="da2ee-109">But there is no `type` parameter for the `<Override>` element.</span></span> <span data-ttu-id="da2ee-110">相違点は、親要素と親要素の型によって決まります。</span><span class="sxs-lookup"><span data-stu-id="da2ee-110">The difference is determined by the parent element and the parent element's type.</span></span> <span data-ttu-id="da2ee-111">がである要素 `<Override>` の内部にある要素は `<Token>` `xsi:type` `RequirementToken` 、 **RequirementTokenOverride** 型である必要があります。</span><span class="sxs-lookup"><span data-stu-id="da2ee-111">An `<Override>` element that is inside of a `<Token>` element whose `xsi:type` is `RequirementToken`, must be of type **RequirementTokenOverride**.</span></span> <span data-ttu-id="da2ee-112">`<Override>`他の親要素の中、または型の要素内の要素は `<Override>` `LocaleToken` 、 **LocaleTokenOverride** 型でなければなりません。</span><span class="sxs-lookup"><span data-stu-id="da2ee-112">An `<Override>` element inside any other parent element, or inside an `<Override>` element of type `LocaleToken`, must be of type **LocaleTokenOverride**.</span></span> <span data-ttu-id="da2ee-113">それぞれの種類について、以下の個別のセクションで説明します。</span><span class="sxs-lookup"><span data-stu-id="da2ee-113">Each type is described in separate sections below.</span></span>

## <a name="override-element-of-type-localetokenoverride"></a><span data-ttu-id="da2ee-114">LocaleTokenOverride 型の Override 要素</span><span class="sxs-lookup"><span data-stu-id="da2ee-114">Override element of type LocaleTokenOverride</span></span>

<span data-ttu-id="da2ee-115">`<Override>`要素は条件を表し、"If...[...]if.</span><span class="sxs-lookup"><span data-stu-id="da2ee-115">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="da2ee-116">要素の `<Override>` 型が **LocaleTokenOverride** の場合は、 `Locale` 属性は条件です。属性はその後のものです `Value` 。</span><span class="sxs-lookup"><span data-stu-id="da2ee-116">If the `<Override>` element is of type **LocaleTokenOverride** , then the `Locale` attribute is the condition, and the `Value` attribute is the consequent.</span></span> <span data-ttu-id="da2ee-117">たとえば、"Office ロケール設定が fr-fr で、表示名が ' Lecteur vidéo ' の場合は、次の値が読み取られます。</span><span class="sxs-lookup"><span data-stu-id="da2ee-117">For example, the following is read "If the Office locale setting is fr-fr, then the display name is 'Lecteur vidéo'."</span></span>

```xml
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

<span data-ttu-id="da2ee-118">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="da2ee-118">**Add-in type:** Content, Task pane, Mail</span></span>

### <a name="syntax"></a><span data-ttu-id="da2ee-119">構文</span><span class="sxs-lookup"><span data-stu-id="da2ee-119">Syntax</span></span>

```XML
<Override Locale="string" Value="string"></Override>
```

### <a name="contained-in"></a><span data-ttu-id="da2ee-120">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="da2ee-120">Contained in</span></span>

|<span data-ttu-id="da2ee-121">要素</span><span class="sxs-lookup"><span data-stu-id="da2ee-121">Element</span></span>|
|:-----|
|[<span data-ttu-id="da2ee-122">CitationText</span><span class="sxs-lookup"><span data-stu-id="da2ee-122">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="da2ee-123">説明</span><span class="sxs-lookup"><span data-stu-id="da2ee-123">Description</span></span>](description.md)|
|[<span data-ttu-id="da2ee-124">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="da2ee-124">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="da2ee-125">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="da2ee-125">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="da2ee-126">DisplayName</span><span class="sxs-lookup"><span data-stu-id="da2ee-126">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="da2ee-127">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="da2ee-127">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="da2ee-128">IconUrl</span><span class="sxs-lookup"><span data-stu-id="da2ee-128">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="da2ee-129">QueryUri</span><span class="sxs-lookup"><span data-stu-id="da2ee-129">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="da2ee-130">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="da2ee-130">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="da2ee-131">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="da2ee-131">SupportUrl</span></span>](supporturl.md)|
|[<span data-ttu-id="da2ee-132">トークン</span><span class="sxs-lookup"><span data-stu-id="da2ee-132">Token</span></span>](token.md)|

### <a name="attributes"></a><span data-ttu-id="da2ee-133">属性</span><span class="sxs-lookup"><span data-stu-id="da2ee-133">Attributes</span></span>

|<span data-ttu-id="da2ee-134">属性</span><span class="sxs-lookup"><span data-stu-id="da2ee-134">Attribute</span></span>|<span data-ttu-id="da2ee-135">型</span><span class="sxs-lookup"><span data-stu-id="da2ee-135">Type</span></span>|<span data-ttu-id="da2ee-136">必須</span><span class="sxs-lookup"><span data-stu-id="da2ee-136">Required</span></span>|<span data-ttu-id="da2ee-137">説明</span><span class="sxs-lookup"><span data-stu-id="da2ee-137">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="da2ee-138">Locale</span><span class="sxs-lookup"><span data-stu-id="da2ee-138">Locale</span></span>|<span data-ttu-id="da2ee-139">string</span><span class="sxs-lookup"><span data-stu-id="da2ee-139">string</span></span>|<span data-ttu-id="da2ee-140">必須</span><span class="sxs-lookup"><span data-stu-id="da2ee-140">required</span></span>|<span data-ttu-id="da2ee-141">`"en-US"` などの BCP 47 言語タグの書式で、この上書きのロケールのカルチャ名を指定します。</span><span class="sxs-lookup"><span data-stu-id="da2ee-141">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="da2ee-142">Value</span><span class="sxs-lookup"><span data-stu-id="da2ee-142">Value</span></span>|<span data-ttu-id="da2ee-143">string</span><span class="sxs-lookup"><span data-stu-id="da2ee-143">string</span></span>|<span data-ttu-id="da2ee-144">必須</span><span class="sxs-lookup"><span data-stu-id="da2ee-144">required</span></span>|<span data-ttu-id="da2ee-145">指定のロケールに対して表される設定の値を指定します。</span><span class="sxs-lookup"><span data-stu-id="da2ee-145">Specifies value of the setting expressed for the specified locale.</span></span>|

### <a name="examples"></a><span data-ttu-id="da2ee-146">例</span><span class="sxs-lookup"><span data-stu-id="da2ee-146">Examples</span></span>

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

### <a name="see-also"></a><span data-ttu-id="da2ee-147">関連項目</span><span class="sxs-lookup"><span data-stu-id="da2ee-147">See also</span></span>

- [<span data-ttu-id="da2ee-148">Office アドインのローカライズ</span><span class="sxs-lookup"><span data-stu-id="da2ee-148">Localization for Office Add-ins</span></span>](../../develop/localization.md)
- [<span data-ttu-id="da2ee-149">SharePoint のキーボード ショートカット</span><span class="sxs-lookup"><span data-stu-id="da2ee-149">Keyboard shortcuts</span></span>](../../design/keyboard-shortcuts.md)

## <a name="override-element-of-type-requirementtokenoverride"></a><span data-ttu-id="da2ee-150">RequirementTokenOverride 型の Override 要素</span><span class="sxs-lookup"><span data-stu-id="da2ee-150">Override element of type RequirementTokenOverride</span></span>

<span data-ttu-id="da2ee-151">`<Override>`要素は条件を表し、"If...[...]if.</span><span class="sxs-lookup"><span data-stu-id="da2ee-151">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="da2ee-152">要素の `<Override>` 型が **RequirementTokenOverride** の場合、子要素は `<Requirements>` 条件を表し、属性はその後の `Value` ものです。</span><span class="sxs-lookup"><span data-stu-id="da2ee-152">If the `<Override>` element is of type **RequirementTokenOverride** , then the child `<Requirements>` element expresses the condition, and the `Value` attribute is the consequent.</span></span> <span data-ttu-id="da2ee-153">たとえば、 `<Override>` 現在のプラットフォームが FeatureOne version 1.7 をサポートしている場合は、次のように "oldAddinVersion" を使用します。これは、 `${token.requirements}` 既定の文字列 ' upgrade ' ではなく、祖父母の URL に含まれるトークンの代わりに使用され `<ExtendedOverrides>` ます。</span><span class="sxs-lookup"><span data-stu-id="da2ee-153">For example, the first `<Override>` in the following is read "If the current platform supports FeatureOne version 1.7, then use string 'oldAddinVersion' in place of the `${token.requirements}` token in the URL of the grandparent `<ExtendedOverrides>` (instead of the default string 'upgrade')."</span></span>

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

<span data-ttu-id="da2ee-154">**アドインの種類:** 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="da2ee-154">**Add-in type:** Task pane</span></span>

### <a name="syntax"></a><span data-ttu-id="da2ee-155">構文</span><span class="sxs-lookup"><span data-stu-id="da2ee-155">Syntax</span></span>

```XML
<Override Value="string" />
```

### <a name="contained-in"></a><span data-ttu-id="da2ee-156">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="da2ee-156">Contained in</span></span>

|<span data-ttu-id="da2ee-157">要素</span><span class="sxs-lookup"><span data-stu-id="da2ee-157">Element</span></span>|
|:-----|
|[<span data-ttu-id="da2ee-158">トークン</span><span class="sxs-lookup"><span data-stu-id="da2ee-158">Token</span></span>](token.md)|

### <a name="must-contain"></a><span data-ttu-id="da2ee-159">含める必要があるもの</span><span class="sxs-lookup"><span data-stu-id="da2ee-159">Must contain</span></span>

|<span data-ttu-id="da2ee-160">要素</span><span class="sxs-lookup"><span data-stu-id="da2ee-160">Element</span></span>|<span data-ttu-id="da2ee-161">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="da2ee-161">Content</span></span>|<span data-ttu-id="da2ee-162">メール</span><span class="sxs-lookup"><span data-stu-id="da2ee-162">Mail</span></span>|<span data-ttu-id="da2ee-163">TaskPane</span><span class="sxs-lookup"><span data-stu-id="da2ee-163">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="da2ee-164">Requirements</span><span class="sxs-lookup"><span data-stu-id="da2ee-164">Requirements</span></span>](requirements.md)|||<span data-ttu-id="da2ee-165">x</span><span class="sxs-lookup"><span data-stu-id="da2ee-165">x</span></span>|

### <a name="attributes"></a><span data-ttu-id="da2ee-166">属性</span><span class="sxs-lookup"><span data-stu-id="da2ee-166">Attributes</span></span>

|<span data-ttu-id="da2ee-167">属性</span><span class="sxs-lookup"><span data-stu-id="da2ee-167">Attribute</span></span>|<span data-ttu-id="da2ee-168">型</span><span class="sxs-lookup"><span data-stu-id="da2ee-168">Type</span></span>|<span data-ttu-id="da2ee-169">必須</span><span class="sxs-lookup"><span data-stu-id="da2ee-169">Required</span></span>|<span data-ttu-id="da2ee-170">説明</span><span class="sxs-lookup"><span data-stu-id="da2ee-170">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="da2ee-171">値</span><span class="sxs-lookup"><span data-stu-id="da2ee-171">Value</span></span>|<span data-ttu-id="da2ee-172">string</span><span class="sxs-lookup"><span data-stu-id="da2ee-172">string</span></span>|<span data-ttu-id="da2ee-173">必須</span><span class="sxs-lookup"><span data-stu-id="da2ee-173">required</span></span>|<span data-ttu-id="da2ee-174">条件が満たされた場合の祖父母トークンの値。</span><span class="sxs-lookup"><span data-stu-id="da2ee-174">Value of the grandparent token when the condition is satisfied.</span></span>|

### <a name="example"></a><span data-ttu-id="da2ee-175">例</span><span class="sxs-lookup"><span data-stu-id="da2ee-175">Example</span></span>

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

### <a name="see-also"></a><span data-ttu-id="da2ee-176">関連項目</span><span class="sxs-lookup"><span data-stu-id="da2ee-176">See also</span></span>

- [<span data-ttu-id="da2ee-177">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="da2ee-177">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="da2ee-178">マニフェストで Requirements 要素を設定する</span><span class="sxs-lookup"><span data-stu-id="da2ee-178">Set the Requirements element in the manifest</span></span>](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)
- [<span data-ttu-id="da2ee-179">SharePoint のキーボード ショートカット</span><span class="sxs-lookup"><span data-stu-id="da2ee-179">Keyboard shortcuts</span></span>](../../design/keyboard-shortcuts.md)
