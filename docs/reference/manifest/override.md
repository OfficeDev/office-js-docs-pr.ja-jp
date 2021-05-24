---
title: マニフェスト ファイルの Override 要素
description: Override 要素を使用すると、指定した条件に応じて設定の値を指定できます。
ms.date: 05/19/2021
localization_priority: Normal
ms.openlocfilehash: cd270fa19750810238b42c26c2abc35a61c1bac8
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590905"
---
# <a name="override-element"></a><span data-ttu-id="58b59-103">Override 要素</span><span class="sxs-lookup"><span data-stu-id="58b59-103">Override element</span></span>

<span data-ttu-id="58b59-104">指定した条件に応じてマニフェスト設定の値を上書きする方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="58b59-104">Provides a way to override the value of a manifest setting depending on a specified condition.</span></span> <span data-ttu-id="58b59-105">条件には次の 3 種類があります。</span><span class="sxs-lookup"><span data-stu-id="58b59-105">There are three kinds of conditions:</span></span>

- <span data-ttu-id="58b59-106">LocaleTokenOverride と呼ばれる既定のロケールとは異なるOfficeロケールです `LocaleToken` 。 </span><span class="sxs-lookup"><span data-stu-id="58b59-106">An Office locale that is different from the default `LocaleToken`, called **LocaleTokenOverride**.</span></span>
- <span data-ttu-id="58b59-107">RequirementTokenOverride と呼ばれる既定のパターンとは異なる、要件セット `RequirementToken` **のサポートのパターン** です。</span><span class="sxs-lookup"><span data-stu-id="58b59-107">A pattern of requirement set support that is different from the default `RequirementToken` pattern, called **RequirementTokenOverride**.</span></span>
- <span data-ttu-id="58b59-108">ソースは `Runtime` 、RuntimeOverride と呼ばれる既定 **のソースとは異なります**。</span><span class="sxs-lookup"><span data-stu-id="58b59-108">The source is different from the default `Runtime`, called **RuntimeOverride**.</span></span>

<span data-ttu-id="58b59-109">要素 `<Override>` の内部にある要素は `<Runtime>` **、RuntimeOverride 型である必要があります**。</span><span class="sxs-lookup"><span data-stu-id="58b59-109">An `<Override>` element that is inside of a `<Runtime>` element must be of type **RuntimeOverride**.</span></span>

<span data-ttu-id="58b59-110">要素の `overrideType` 属性 `<Override>` はありません。</span><span class="sxs-lookup"><span data-stu-id="58b59-110">There is no `overrideType` attribute for the `<Override>` element.</span></span> <span data-ttu-id="58b59-111">違いは、親要素と親要素の型によって決まります。</span><span class="sxs-lookup"><span data-stu-id="58b59-111">The difference is determined by the parent element and the parent element's type.</span></span> <span data-ttu-id="58b59-112">要素 `<Override>` の内部にある要素は `<Token>` `xsi:type` `RequirementToken` **、RequirementTokenOverride 型である必要があります**。</span><span class="sxs-lookup"><span data-stu-id="58b59-112">An `<Override>` element that is inside of a `<Token>` element whose `xsi:type` is `RequirementToken`, must be of type **RequirementTokenOverride**.</span></span> <span data-ttu-id="58b59-113">他 `<Override>` の親要素内の要素、または型の要素内の要素は `<Override>` `LocaleToken` **、LocaleTokenOverride 型である必要があります**。</span><span class="sxs-lookup"><span data-stu-id="58b59-113">An `<Override>` element inside any other parent element, or inside an `<Override>` element of type `LocaleToken`, must be of type **LocaleTokenOverride**.</span></span> <span data-ttu-id="58b59-114">要素の子である場合のこの要素の使用の詳細については、「マニフェストの拡張オーバーライドを処理する」 `<Token>` [を参照してください](../../develop/extended-overrides.md)。</span><span class="sxs-lookup"><span data-stu-id="58b59-114">For more information about the use of this element when it is a child of a `<Token>` element, see [Work with extended overrides of the manifest](../../develop/extended-overrides.md).</span></span>

<span data-ttu-id="58b59-115">各種類については、この記事で後述する個別のセクションで説明します。</span><span class="sxs-lookup"><span data-stu-id="58b59-115">Each type is described in separate sections later in this article.</span></span>

## <a name="override-element-for-localetoken"></a><span data-ttu-id="58b59-116">Override 要素 `LocaleToken`</span><span class="sxs-lookup"><span data-stu-id="58b59-116">Override element for `LocaleToken`</span></span>

<span data-ttu-id="58b59-117">要素 `<Override>` は条件付きを表し、"If .." として読み取り可能です。その後 ..."。ステートメント。</span><span class="sxs-lookup"><span data-stu-id="58b59-117">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="58b59-118">要素が `<Override>` **LocaleTokenOverride** 型の場合、属性は条件であり、その `Locale` `Value` 結果属性になります。</span><span class="sxs-lookup"><span data-stu-id="58b59-118">If the `<Override>` element is of type **LocaleTokenOverride**, then the `Locale` attribute is the condition, and the `Value` attribute is the consequent.</span></span> <span data-ttu-id="58b59-119">たとえば、次の例では、「ロケールOffice fr-fr の場合、表示名は 'Lecteur vidéo'です。</span><span class="sxs-lookup"><span data-stu-id="58b59-119">For example, the following is read "If the Office locale setting is fr-fr, then the display name is 'Lecteur vidéo'."</span></span>

```xml
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

<span data-ttu-id="58b59-120">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="58b59-120">**Add-in type:** Content, Task pane, Mail</span></span>

### <a name="syntax"></a><span data-ttu-id="58b59-121">構文</span><span class="sxs-lookup"><span data-stu-id="58b59-121">Syntax</span></span>

```XML
<Override Locale="string" Value="string"></Override>
```

### <a name="contained-in"></a><span data-ttu-id="58b59-122">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="58b59-122">Contained in</span></span>

|<span data-ttu-id="58b59-123">要素</span><span class="sxs-lookup"><span data-stu-id="58b59-123">Element</span></span>|
|:-----|
|[<span data-ttu-id="58b59-124">CitationText</span><span class="sxs-lookup"><span data-stu-id="58b59-124">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="58b59-125">説明</span><span class="sxs-lookup"><span data-stu-id="58b59-125">Description</span></span>](description.md)|
|[<span data-ttu-id="58b59-126">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="58b59-126">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="58b59-127">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="58b59-127">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="58b59-128">DisplayName</span><span class="sxs-lookup"><span data-stu-id="58b59-128">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="58b59-129">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="58b59-129">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="58b59-130">IconUrl</span><span class="sxs-lookup"><span data-stu-id="58b59-130">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="58b59-131">QueryUri</span><span class="sxs-lookup"><span data-stu-id="58b59-131">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="58b59-132">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="58b59-132">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="58b59-133">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="58b59-133">SupportUrl</span></span>](supporturl.md)|
|[<span data-ttu-id="58b59-134">トークン</span><span class="sxs-lookup"><span data-stu-id="58b59-134">Token</span></span>](token.md)|

### <a name="attributes"></a><span data-ttu-id="58b59-135">属性</span><span class="sxs-lookup"><span data-stu-id="58b59-135">Attributes</span></span>

|<span data-ttu-id="58b59-136">属性</span><span class="sxs-lookup"><span data-stu-id="58b59-136">Attribute</span></span>|<span data-ttu-id="58b59-137">型</span><span class="sxs-lookup"><span data-stu-id="58b59-137">Type</span></span>|<span data-ttu-id="58b59-138">必須</span><span class="sxs-lookup"><span data-stu-id="58b59-138">Required</span></span>|<span data-ttu-id="58b59-139">説明</span><span class="sxs-lookup"><span data-stu-id="58b59-139">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="58b59-140">Locale</span><span class="sxs-lookup"><span data-stu-id="58b59-140">Locale</span></span>|<span data-ttu-id="58b59-141">string</span><span class="sxs-lookup"><span data-stu-id="58b59-141">string</span></span>|<span data-ttu-id="58b59-142">必須</span><span class="sxs-lookup"><span data-stu-id="58b59-142">required</span></span>|<span data-ttu-id="58b59-143">`"en-US"` などの BCP 47 言語タグの書式で、この上書きのロケールのカルチャ名を指定します。</span><span class="sxs-lookup"><span data-stu-id="58b59-143">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="58b59-144">Value</span><span class="sxs-lookup"><span data-stu-id="58b59-144">Value</span></span>|<span data-ttu-id="58b59-145">string</span><span class="sxs-lookup"><span data-stu-id="58b59-145">string</span></span>|<span data-ttu-id="58b59-146">必須</span><span class="sxs-lookup"><span data-stu-id="58b59-146">required</span></span>|<span data-ttu-id="58b59-147">指定のロケールに対して表される設定の値を指定します。</span><span class="sxs-lookup"><span data-stu-id="58b59-147">Specifies value of the setting expressed for the specified locale.</span></span>|

### <a name="examples"></a><span data-ttu-id="58b59-148">例</span><span class="sxs-lookup"><span data-stu-id="58b59-148">Examples</span></span>

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

### <a name="see-also"></a><span data-ttu-id="58b59-149">関連項目</span><span class="sxs-lookup"><span data-stu-id="58b59-149">See also</span></span>

- [<span data-ttu-id="58b59-150">Office アドインのローカライズ</span><span class="sxs-lookup"><span data-stu-id="58b59-150">Localization for Office Add-ins</span></span>](../../develop/localization.md)
- [<span data-ttu-id="58b59-151">SharePoint のキーボード ショートカット</span><span class="sxs-lookup"><span data-stu-id="58b59-151">Keyboard shortcuts</span></span>](../../design/keyboard-shortcuts.md)

## <a name="override-element-for-requirementtoken"></a><span data-ttu-id="58b59-152">Override 要素 `RequirementToken`</span><span class="sxs-lookup"><span data-stu-id="58b59-152">Override element for `RequirementToken`</span></span>

<span data-ttu-id="58b59-153">要素 `<Override>` は条件付きを表し、"If .." として読み取り可能です。その後 ..."。ステートメント。</span><span class="sxs-lookup"><span data-stu-id="58b59-153">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="58b59-154">要素が `<Override>` **RequirementTokenOverride** 型の場合、子要素は条件を表し、属性 `<Requirements>` `Value` はその結果です。</span><span class="sxs-lookup"><span data-stu-id="58b59-154">If the `<Override>` element is of type **RequirementTokenOverride**, then the child `<Requirements>` element expresses the condition, and the `Value` attribute is the consequent.</span></span> <span data-ttu-id="58b59-155">たとえば、次の 1 つ目は、「現在のプラットフォームが FeatureOne バージョン 1.7 をサポートしている場合は、(既定の文字列 'upgrade' ではなく) 祖父母の URL のトークンの代わりに文字列 `<Override>` 'oldAddinVersion' を使用します。 `${token.requirements}` `<ExtendedOverrides>`</span><span class="sxs-lookup"><span data-stu-id="58b59-155">For example, the first `<Override>` in the following is read "If the current platform supports FeatureOne version 1.7, then use string 'oldAddinVersion' in place of the `${token.requirements}` token in the URL of the grandparent `<ExtendedOverrides>` (instead of the default string 'upgrade')."</span></span>

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

<span data-ttu-id="58b59-156">**アドインの種類:** 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="58b59-156">**Add-in type:** Task pane</span></span>

### <a name="syntax"></a><span data-ttu-id="58b59-157">構文</span><span class="sxs-lookup"><span data-stu-id="58b59-157">Syntax</span></span>

```XML
<Override Value="string" />
```

### <a name="contained-in"></a><span data-ttu-id="58b59-158">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="58b59-158">Contained in</span></span>

|<span data-ttu-id="58b59-159">要素</span><span class="sxs-lookup"><span data-stu-id="58b59-159">Element</span></span>|
|:-----|
|[<span data-ttu-id="58b59-160">トークン</span><span class="sxs-lookup"><span data-stu-id="58b59-160">Token</span></span>](token.md)|

### <a name="must-contain"></a><span data-ttu-id="58b59-161">含める必要があるもの</span><span class="sxs-lookup"><span data-stu-id="58b59-161">Must contain</span></span>

|<span data-ttu-id="58b59-162">要素</span><span class="sxs-lookup"><span data-stu-id="58b59-162">Element</span></span>|<span data-ttu-id="58b59-163">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="58b59-163">Content</span></span>|<span data-ttu-id="58b59-164">メール</span><span class="sxs-lookup"><span data-stu-id="58b59-164">Mail</span></span>|<span data-ttu-id="58b59-165">TaskPane</span><span class="sxs-lookup"><span data-stu-id="58b59-165">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="58b59-166">Requirements</span><span class="sxs-lookup"><span data-stu-id="58b59-166">Requirements</span></span>](requirements.md)|||<span data-ttu-id="58b59-167">x</span><span class="sxs-lookup"><span data-stu-id="58b59-167">x</span></span>|

### <a name="attributes"></a><span data-ttu-id="58b59-168">属性</span><span class="sxs-lookup"><span data-stu-id="58b59-168">Attributes</span></span>

|<span data-ttu-id="58b59-169">属性</span><span class="sxs-lookup"><span data-stu-id="58b59-169">Attribute</span></span>|<span data-ttu-id="58b59-170">型</span><span class="sxs-lookup"><span data-stu-id="58b59-170">Type</span></span>|<span data-ttu-id="58b59-171">必須</span><span class="sxs-lookup"><span data-stu-id="58b59-171">Required</span></span>|<span data-ttu-id="58b59-172">説明</span><span class="sxs-lookup"><span data-stu-id="58b59-172">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="58b59-173">値</span><span class="sxs-lookup"><span data-stu-id="58b59-173">Value</span></span>|<span data-ttu-id="58b59-174">string</span><span class="sxs-lookup"><span data-stu-id="58b59-174">string</span></span>|<span data-ttu-id="58b59-175">必須</span><span class="sxs-lookup"><span data-stu-id="58b59-175">required</span></span>|<span data-ttu-id="58b59-176">条件が満たされた場合の祖父母トークンの値。</span><span class="sxs-lookup"><span data-stu-id="58b59-176">Value of the grandparent token when the condition is satisfied.</span></span>|

### <a name="example"></a><span data-ttu-id="58b59-177">例</span><span class="sxs-lookup"><span data-stu-id="58b59-177">Example</span></span>

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

### <a name="see-also"></a><span data-ttu-id="58b59-178">関連項目</span><span class="sxs-lookup"><span data-stu-id="58b59-178">See also</span></span>

- [<span data-ttu-id="58b59-179">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="58b59-179">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="58b59-180">マニフェストで Requirements 要素を設定する</span><span class="sxs-lookup"><span data-stu-id="58b59-180">Set the Requirements element in the manifest</span></span>](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)
- [<span data-ttu-id="58b59-181">SharePoint のキーボード ショートカット</span><span class="sxs-lookup"><span data-stu-id="58b59-181">Keyboard shortcuts</span></span>](../../design/keyboard-shortcuts.md)

## <a name="override-element-for-runtime"></a><span data-ttu-id="58b59-182">Override 要素 `Runtime`</span><span class="sxs-lookup"><span data-stu-id="58b59-182">Override element for `Runtime`</span></span>

> [!IMPORTANT]
> <span data-ttu-id="58b59-183">この要素のサポートは、イベント ベースのアクティブ化機能を備えたメールボックス要件 [セット 1.10](../../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md) [で導入されました](../../outlook/autolaunch.md)。</span><span class="sxs-lookup"><span data-stu-id="58b59-183">Support for this element was introduced in [Mailbox requirement set 1.10](../../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md) with the [event-based activation feature](../../outlook/autolaunch.md).</span></span> <span data-ttu-id="58b59-184">この要件セットをサポートする [クライアントおよびプラットフォーム](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) を参照してください。</span><span class="sxs-lookup"><span data-stu-id="58b59-184">See [clients and platforms](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

<span data-ttu-id="58b59-185">要素 `<Override>` は条件付きを表し、"If .." として読み取り可能です。その後 ..."。ステートメント。</span><span class="sxs-lookup"><span data-stu-id="58b59-185">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="58b59-186">要素が RuntimeOverride 型の場合、属性は条件であり、属性 `<Override>`  `type` `resid` は結果です。</span><span class="sxs-lookup"><span data-stu-id="58b59-186">If the `<Override>` element is of type **RuntimeOverride**, then the `type` attribute is the condition, and the `resid` attribute is the consequent.</span></span> <span data-ttu-id="58b59-187">たとえば、「型が 'javascript'の場合は `resid` 、'JSRuntime.Url'です」と読み取ります。Outlookデスクトップでは、LaunchEvent 拡張ポイント[ハンドラーに対してこの要素が](../../reference/manifest/extensionpoint.md#launchevent)必要です。</span><span class="sxs-lookup"><span data-stu-id="58b59-187">For example, the following is read "If the type is 'javascript', then the `resid` is 'JSRuntime.Url'." Outlook Desktop requires this element for [LaunchEvent extension point](../../reference/manifest/extensionpoint.md#launchevent) handlers.</span></span>

```xml
<Runtime resid="WebViewRuntime.Url">
  <Override type="javascript" resid="JSRuntime.Url"/>
</Runtime>
```

<span data-ttu-id="58b59-188">**アドインの種類:** メール</span><span class="sxs-lookup"><span data-stu-id="58b59-188">**Add-in type:** Mail</span></span>

### <a name="syntax"></a><span data-ttu-id="58b59-189">構文</span><span class="sxs-lookup"><span data-stu-id="58b59-189">Syntax</span></span>

```XML
<Override type="javascript" resid="JSRuntime.Url"/>
```

### <a name="contained-in"></a><span data-ttu-id="58b59-190">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="58b59-190">Contained in</span></span>

- [<span data-ttu-id="58b59-191">ランタイム</span><span class="sxs-lookup"><span data-stu-id="58b59-191">Runtime</span></span>](runtime.md)

### <a name="attributes"></a><span data-ttu-id="58b59-192">属性</span><span class="sxs-lookup"><span data-stu-id="58b59-192">Attributes</span></span>

|<span data-ttu-id="58b59-193">属性</span><span class="sxs-lookup"><span data-stu-id="58b59-193">Attribute</span></span>|<span data-ttu-id="58b59-194">型</span><span class="sxs-lookup"><span data-stu-id="58b59-194">Type</span></span>|<span data-ttu-id="58b59-195">必須</span><span class="sxs-lookup"><span data-stu-id="58b59-195">Required</span></span>|<span data-ttu-id="58b59-196">説明</span><span class="sxs-lookup"><span data-stu-id="58b59-196">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="58b59-197">**type**</span><span class="sxs-lookup"><span data-stu-id="58b59-197">**type**</span></span>|<span data-ttu-id="58b59-198">string</span><span class="sxs-lookup"><span data-stu-id="58b59-198">string</span></span>|<span data-ttu-id="58b59-199">はい</span><span class="sxs-lookup"><span data-stu-id="58b59-199">Yes</span></span>|<span data-ttu-id="58b59-200">このオーバーライドの言語を指定します。</span><span class="sxs-lookup"><span data-stu-id="58b59-200">Specifies the language for this override.</span></span> <span data-ttu-id="58b59-201">現時点では、 `"javascript"` サポートされている唯一のオプションです。</span><span class="sxs-lookup"><span data-stu-id="58b59-201">At present, `"javascript"` is the only supported option.</span></span>|
|<span data-ttu-id="58b59-202">**resid**</span><span class="sxs-lookup"><span data-stu-id="58b59-202">**resid**</span></span>|<span data-ttu-id="58b59-203">文字列</span><span class="sxs-lookup"><span data-stu-id="58b59-203">string</span></span>|<span data-ttu-id="58b59-204">はい</span><span class="sxs-lookup"><span data-stu-id="58b59-204">Yes</span></span>|<span data-ttu-id="58b59-205">親 [ランタイム](runtime.md) 要素で定義されている既定の HTML の URL の場所を上書きする JavaScript ファイルの URL の場所を指定します `resid` 。</span><span class="sxs-lookup"><span data-stu-id="58b59-205">Specifies the URL location of the JavaScript file that should override the URL location of the default HTML defined in the parent [Runtime](runtime.md) element's `resid`.</span></span> <span data-ttu-id="58b59-206">32 文字以内で、要素内の要素の属性と一致 `resid` `id` `Url` する必要 `Resources` があります。</span><span class="sxs-lookup"><span data-stu-id="58b59-206">The `resid` can be no more than 32 characters and must match an `id` attribute of a `Url` element in the `Resources` element.</span></span>|

### <a name="examples"></a><span data-ttu-id="58b59-207">例</span><span class="sxs-lookup"><span data-stu-id="58b59-207">Examples</span></span>

```xml
<!-- Event-based activation happens in a lightweight runtime.-->
<Runtimes>
  <!-- HTML file including reference to or inline JavaScript event handlers.
  This is used by Outlook on the web. -->
  <Runtime resid="WebViewRuntime.Url">
    <!-- JavaScript file containing event handlers. This is used by Outlook Desktop. -->
    <Override type="javascript" resid="JSRuntime.Url"/>
  </Runtime>
</Runtimes>
```

### <a name="see-also"></a><span data-ttu-id="58b59-208">関連項目</span><span class="sxs-lookup"><span data-stu-id="58b59-208">See also</span></span>

- [<span data-ttu-id="58b59-209">ランタイム</span><span class="sxs-lookup"><span data-stu-id="58b59-209">Runtime</span></span>](runtime.md)
- [<span data-ttu-id="58b59-210">イベント ベースのOutlook用にアドインを構成する</span><span class="sxs-lookup"><span data-stu-id="58b59-210">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
