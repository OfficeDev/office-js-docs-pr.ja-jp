---
title: マニフェスト ファイルの Override 要素
description: Override 要素を使用すると、指定した条件に応じて設定の値を指定できます。
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: 131d72883d050038e2df5b7d8bbca033af9e6ee4
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555158"
---
# <a name="override-element"></a><span data-ttu-id="bbca2-103">Override 要素</span><span class="sxs-lookup"><span data-stu-id="bbca2-103">Override element</span></span>

<span data-ttu-id="bbca2-104">指定した条件に応じて、マニフェスト設定の値をオーバーライドする方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="bbca2-104">Provides a way to override the value of a manifest setting depending on a specified condition.</span></span> <span data-ttu-id="bbca2-105">条件には、次の 3 種類があります。</span><span class="sxs-lookup"><span data-stu-id="bbca2-105">There are three kinds of conditions:</span></span>

- <span data-ttu-id="bbca2-106">既定のロケールとは異なるOffice ロケールです `LocaleToken` 。 </span><span class="sxs-lookup"><span data-stu-id="bbca2-106">An Office locale that is different from the default `LocaleToken`, called **LocaleTokenOverride**.</span></span>
- <span data-ttu-id="bbca2-107">要件セットのサポートのパターンは、既定のパターンとは異なり `RequirementToken` 、 **要件TokenOverride** と呼ばれます。</span><span class="sxs-lookup"><span data-stu-id="bbca2-107">A pattern of requirement set support that is different from the default `RequirementToken` pattern, called **RequirementTokenOverride**.</span></span>
- <span data-ttu-id="bbca2-108">ソースは `Runtime` **、RuntimeOverride** (現在プレビュー中) と呼ばれる既定のとは異なります。</span><span class="sxs-lookup"><span data-stu-id="bbca2-108">The source is different from the default `Runtime`, called **RuntimeOverride** (currently in preview).</span></span>

<span data-ttu-id="bbca2-109">`<Override>`要素の内部にある要素は、 `<Runtime>` 型が **RuntimeOverride** である必要があります。</span><span class="sxs-lookup"><span data-stu-id="bbca2-109">An `<Override>` element that is inside of a `<Runtime>` element must be of type **RuntimeOverride**.</span></span>

<span data-ttu-id="bbca2-110">`overrideType`要素の属性がありません `<Override>` 。</span><span class="sxs-lookup"><span data-stu-id="bbca2-110">There is no `overrideType` attribute for the `<Override>` element.</span></span> <span data-ttu-id="bbca2-111">違いは、親要素と親要素の型によって決まります。</span><span class="sxs-lookup"><span data-stu-id="bbca2-111">The difference is determined by the parent element and the parent element's type.</span></span> <span data-ttu-id="bbca2-112">である `<Override>` 要素の内部にある `<Token>` 要素 `xsi:type` は `RequirementToken` 、型 **が "要件TokenOverride"** である必要があります。</span><span class="sxs-lookup"><span data-stu-id="bbca2-112">An `<Override>` element that is inside of a `<Token>` element whose `xsi:type` is `RequirementToken`, must be of type **RequirementTokenOverride**.</span></span> <span data-ttu-id="bbca2-113">`<Override>`他の親要素の中、または `<Override>` 型の要素内の要素 `LocaleToken` は **、型が LocaleTokenOverride** である必要があります。</span><span class="sxs-lookup"><span data-stu-id="bbca2-113">An `<Override>` element inside any other parent element, or inside an `<Override>` element of type `LocaleToken`, must be of type **LocaleTokenOverride**.</span></span> <span data-ttu-id="bbca2-114">要素の子である場合にこの要素を使用する方法の詳細については `<Token>` 、「 マニフェストの [拡張オーバーライドを処理する](../../develop/extended-overrides.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="bbca2-114">For more information about the use of this element when it is a child of a `<Token>` element, see [Work with extended overrides of the manifest](../../develop/extended-overrides.md).</span></span>

<span data-ttu-id="bbca2-115">各型については、この記事の後半で説明します。</span><span class="sxs-lookup"><span data-stu-id="bbca2-115">Each type is described in separate sections later in this article.</span></span>

## <a name="override-element-for-localetoken"></a><span data-ttu-id="bbca2-116">要素をオーバーライドする `LocaleToken`</span><span class="sxs-lookup"><span data-stu-id="bbca2-116">Override element for `LocaleToken`</span></span>

<span data-ttu-id="bbca2-117">`<Override>`要素は条件付きを表し、"If..その後..陳述。</span><span class="sxs-lookup"><span data-stu-id="bbca2-117">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="bbca2-118">要素の `<Override>` 型が **LocaleTokenOverride** の場合、 `Locale` 属性は条件であり、 `Value` 属性は結果です。</span><span class="sxs-lookup"><span data-stu-id="bbca2-118">If the `<Override>` element is of type **LocaleTokenOverride**, then the `Locale` attribute is the condition, and the `Value` attribute is the consequent.</span></span> <span data-ttu-id="bbca2-119">たとえば、次の例は、「Officeロケール設定が fr-fr の場合、表示名は 'レクトゥール vidéo' です。</span><span class="sxs-lookup"><span data-stu-id="bbca2-119">For example, the following is read "If the Office locale setting is fr-fr, then the display name is 'Lecteur vidéo'."</span></span>

```xml
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

<span data-ttu-id="bbca2-120">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="bbca2-120">**Add-in type:** Content, Task pane, Mail</span></span>

### <a name="syntax"></a><span data-ttu-id="bbca2-121">構文</span><span class="sxs-lookup"><span data-stu-id="bbca2-121">Syntax</span></span>

```XML
<Override Locale="string" Value="string"></Override>
```

### <a name="contained-in"></a><span data-ttu-id="bbca2-122">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="bbca2-122">Contained in</span></span>

|<span data-ttu-id="bbca2-123">要素</span><span class="sxs-lookup"><span data-stu-id="bbca2-123">Element</span></span>|
|:-----|
|[<span data-ttu-id="bbca2-124">CitationText</span><span class="sxs-lookup"><span data-stu-id="bbca2-124">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="bbca2-125">説明</span><span class="sxs-lookup"><span data-stu-id="bbca2-125">Description</span></span>](description.md)|
|[<span data-ttu-id="bbca2-126">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="bbca2-126">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="bbca2-127">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="bbca2-127">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="bbca2-128">DisplayName</span><span class="sxs-lookup"><span data-stu-id="bbca2-128">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="bbca2-129">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="bbca2-129">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="bbca2-130">IconUrl</span><span class="sxs-lookup"><span data-stu-id="bbca2-130">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="bbca2-131">QueryUri</span><span class="sxs-lookup"><span data-stu-id="bbca2-131">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="bbca2-132">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="bbca2-132">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="bbca2-133">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="bbca2-133">SupportUrl</span></span>](supporturl.md)|
|[<span data-ttu-id="bbca2-134">トークン</span><span class="sxs-lookup"><span data-stu-id="bbca2-134">Token</span></span>](token.md)|

### <a name="attributes"></a><span data-ttu-id="bbca2-135">属性</span><span class="sxs-lookup"><span data-stu-id="bbca2-135">Attributes</span></span>

|<span data-ttu-id="bbca2-136">属性</span><span class="sxs-lookup"><span data-stu-id="bbca2-136">Attribute</span></span>|<span data-ttu-id="bbca2-137">型</span><span class="sxs-lookup"><span data-stu-id="bbca2-137">Type</span></span>|<span data-ttu-id="bbca2-138">必須</span><span class="sxs-lookup"><span data-stu-id="bbca2-138">Required</span></span>|<span data-ttu-id="bbca2-139">説明</span><span class="sxs-lookup"><span data-stu-id="bbca2-139">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="bbca2-140">Locale</span><span class="sxs-lookup"><span data-stu-id="bbca2-140">Locale</span></span>|<span data-ttu-id="bbca2-141">string</span><span class="sxs-lookup"><span data-stu-id="bbca2-141">string</span></span>|<span data-ttu-id="bbca2-142">必須</span><span class="sxs-lookup"><span data-stu-id="bbca2-142">required</span></span>|<span data-ttu-id="bbca2-143">`"en-US"` などの BCP 47 言語タグの書式で、この上書きのロケールのカルチャ名を指定します。</span><span class="sxs-lookup"><span data-stu-id="bbca2-143">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="bbca2-144">Value</span><span class="sxs-lookup"><span data-stu-id="bbca2-144">Value</span></span>|<span data-ttu-id="bbca2-145">string</span><span class="sxs-lookup"><span data-stu-id="bbca2-145">string</span></span>|<span data-ttu-id="bbca2-146">必須</span><span class="sxs-lookup"><span data-stu-id="bbca2-146">required</span></span>|<span data-ttu-id="bbca2-147">指定のロケールに対して表される設定の値を指定します。</span><span class="sxs-lookup"><span data-stu-id="bbca2-147">Specifies value of the setting expressed for the specified locale.</span></span>|

### <a name="examples"></a><span data-ttu-id="bbca2-148">例</span><span class="sxs-lookup"><span data-stu-id="bbca2-148">Examples</span></span>

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

### <a name="see-also"></a><span data-ttu-id="bbca2-149">関連項目</span><span class="sxs-lookup"><span data-stu-id="bbca2-149">See also</span></span>

- [<span data-ttu-id="bbca2-150">Office アドインのローカライズ</span><span class="sxs-lookup"><span data-stu-id="bbca2-150">Localization for Office Add-ins</span></span>](../../develop/localization.md)
- [<span data-ttu-id="bbca2-151">SharePoint のキーボード ショートカット</span><span class="sxs-lookup"><span data-stu-id="bbca2-151">Keyboard shortcuts</span></span>](../../design/keyboard-shortcuts.md)

## <a name="override-element-for-requirementtoken"></a><span data-ttu-id="bbca2-152">要素をオーバーライドする `RequirementToken`</span><span class="sxs-lookup"><span data-stu-id="bbca2-152">Override element for `RequirementToken`</span></span>

<span data-ttu-id="bbca2-153">`<Override>`要素は条件付きを表し、"If..その後..陳述。</span><span class="sxs-lookup"><span data-stu-id="bbca2-153">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="bbca2-154">要素の `<Override>` 型が **"要件TokenOverride"** の場合、子 `<Requirements>` 要素は条件を表し、 `Value` 属性は結果です。</span><span class="sxs-lookup"><span data-stu-id="bbca2-154">If the `<Override>` element is of type **RequirementTokenOverride**, then the child `<Requirements>` element expresses the condition, and the `Value` attribute is the consequent.</span></span> <span data-ttu-id="bbca2-155">たとえば、 `<Override>` 次の最初の例は、「現在のプラットフォームが FeatureOne バージョン 1.7 をサポートしている場合は `${token.requirements}` 、( `<ExtendedOverrides>` 既定の文字列 'upgrade'の代わりに) (既定の文字列 'upgrade') の代わりに、トークンの代わりに文字列 'oldAddinVersion' を使用します。</span><span class="sxs-lookup"><span data-stu-id="bbca2-155">For example, the first `<Override>` in the following is read "If the current platform supports FeatureOne version 1.7, then use string 'oldAddinVersion' in place of the `${token.requirements}` token in the URL of the grandparent `<ExtendedOverrides>` (instead of the default string 'upgrade')."</span></span>

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

<span data-ttu-id="bbca2-156">**アドインの種類:** 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="bbca2-156">**Add-in type:** Task pane</span></span>

### <a name="syntax"></a><span data-ttu-id="bbca2-157">構文</span><span class="sxs-lookup"><span data-stu-id="bbca2-157">Syntax</span></span>

```XML
<Override Value="string" />
```

### <a name="contained-in"></a><span data-ttu-id="bbca2-158">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="bbca2-158">Contained in</span></span>

|<span data-ttu-id="bbca2-159">要素</span><span class="sxs-lookup"><span data-stu-id="bbca2-159">Element</span></span>|
|:-----|
|[<span data-ttu-id="bbca2-160">トークン</span><span class="sxs-lookup"><span data-stu-id="bbca2-160">Token</span></span>](token.md)|

### <a name="must-contain"></a><span data-ttu-id="bbca2-161">含める必要があるもの</span><span class="sxs-lookup"><span data-stu-id="bbca2-161">Must contain</span></span>

|<span data-ttu-id="bbca2-162">要素</span><span class="sxs-lookup"><span data-stu-id="bbca2-162">Element</span></span>|<span data-ttu-id="bbca2-163">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="bbca2-163">Content</span></span>|<span data-ttu-id="bbca2-164">メール</span><span class="sxs-lookup"><span data-stu-id="bbca2-164">Mail</span></span>|<span data-ttu-id="bbca2-165">TaskPane</span><span class="sxs-lookup"><span data-stu-id="bbca2-165">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="bbca2-166">Requirements</span><span class="sxs-lookup"><span data-stu-id="bbca2-166">Requirements</span></span>](requirements.md)|||<span data-ttu-id="bbca2-167">x</span><span class="sxs-lookup"><span data-stu-id="bbca2-167">x</span></span>|

### <a name="attributes"></a><span data-ttu-id="bbca2-168">属性</span><span class="sxs-lookup"><span data-stu-id="bbca2-168">Attributes</span></span>

|<span data-ttu-id="bbca2-169">属性</span><span class="sxs-lookup"><span data-stu-id="bbca2-169">Attribute</span></span>|<span data-ttu-id="bbca2-170">型</span><span class="sxs-lookup"><span data-stu-id="bbca2-170">Type</span></span>|<span data-ttu-id="bbca2-171">必須</span><span class="sxs-lookup"><span data-stu-id="bbca2-171">Required</span></span>|<span data-ttu-id="bbca2-172">説明</span><span class="sxs-lookup"><span data-stu-id="bbca2-172">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="bbca2-173">値</span><span class="sxs-lookup"><span data-stu-id="bbca2-173">Value</span></span>|<span data-ttu-id="bbca2-174">string</span><span class="sxs-lookup"><span data-stu-id="bbca2-174">string</span></span>|<span data-ttu-id="bbca2-175">必須</span><span class="sxs-lookup"><span data-stu-id="bbca2-175">required</span></span>|<span data-ttu-id="bbca2-176">条件が満たされた場合の、祖父母トークンの値。</span><span class="sxs-lookup"><span data-stu-id="bbca2-176">Value of the grandparent token when the condition is satisfied.</span></span>|

### <a name="example"></a><span data-ttu-id="bbca2-177">例</span><span class="sxs-lookup"><span data-stu-id="bbca2-177">Example</span></span>

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

### <a name="see-also"></a><span data-ttu-id="bbca2-178">関連項目</span><span class="sxs-lookup"><span data-stu-id="bbca2-178">See also</span></span>

- [<span data-ttu-id="bbca2-179">Office のバージョンと要件セット</span><span class="sxs-lookup"><span data-stu-id="bbca2-179">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="bbca2-180">マニフェストで Requirements 要素を設定する</span><span class="sxs-lookup"><span data-stu-id="bbca2-180">Set the Requirements element in the manifest</span></span>](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)
- [<span data-ttu-id="bbca2-181">SharePoint のキーボード ショートカット</span><span class="sxs-lookup"><span data-stu-id="bbca2-181">Keyboard shortcuts</span></span>](../../design/keyboard-shortcuts.md)

## <a name="override-element-for-runtime-preview"></a><span data-ttu-id="bbca2-182">要素を上書き `Runtime` (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="bbca2-182">Override element for `Runtime` (preview)</span></span>

> [!IMPORTANT]
> <span data-ttu-id="bbca2-183">この機能は、web 上のOutlookで[プレビュー](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)し、Microsoft 365サブスクリプションでWindowsにのみサポートされます。</span><span class="sxs-lookup"><span data-stu-id="bbca2-183">This feature is only supported for [preview](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) in Outlook on the web and on Windows with a Microsoft 365 subscription.</span></span> <span data-ttu-id="bbca2-184">詳細については、「[イベントベースのアクティブ化用にOutlook アドインを構成する](../../outlook/autolaunch.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="bbca2-184">For more details, see [Configure your Outlook add-in for event-based activation](../../outlook/autolaunch.md).</span></span>
>
> <span data-ttu-id="bbca2-185">プレビュー機能は予告なく変更される場合があるため、運用アドインで使用しないでください。</span><span class="sxs-lookup"><span data-stu-id="bbca2-185">Because preview features are subject to change without notice, they shouldn't be used in production add-ins.</span></span>

<span data-ttu-id="bbca2-186">`<Override>`要素は条件付きを表し、"If..その後..陳述。</span><span class="sxs-lookup"><span data-stu-id="bbca2-186">An `<Override>` element expresses a conditional and can be read as an "If ... then ..." statement.</span></span> <span data-ttu-id="bbca2-187">要素の `<Override>` 型が **RuntimeOverride** の場合、 `type` 属性は条件であり、 `resid` 属性は結果です。</span><span class="sxs-lookup"><span data-stu-id="bbca2-187">If the `<Override>` element is of type **RuntimeOverride**, then the `type` attribute is the condition, and the `resid` attribute is the consequent.</span></span> <span data-ttu-id="bbca2-188">たとえば、次の例は"型が 'javascript' の場合は `resid` 'JSRuntime.Url' です。Outlookデスクトップには[、LaunchEvent 拡張機能ポイント](../../reference/manifest/extensionpoint.md#launchevent-preview)ハンドラーにこの要素が必要です。</span><span class="sxs-lookup"><span data-stu-id="bbca2-188">For example, the following is read "If the type is 'javascript', then the `resid` is 'JSRuntime.Url'." Outlook Desktop requires this element for [LaunchEvent extension point](../../reference/manifest/extensionpoint.md#launchevent-preview) handlers.</span></span>

```xml
<Runtime resid="WebViewRuntime.Url">
  <Override type="javascript" resid="JSRuntime.Url"/>
</Runtime>
```

<span data-ttu-id="bbca2-189">**アドインの種類:** メール</span><span class="sxs-lookup"><span data-stu-id="bbca2-189">**Add-in type:** Mail</span></span>

### <a name="syntax"></a><span data-ttu-id="bbca2-190">構文</span><span class="sxs-lookup"><span data-stu-id="bbca2-190">Syntax</span></span>

```XML
<Override type="javascript" resid="JSRuntime.Url"/>
```

### <a name="contained-in"></a><span data-ttu-id="bbca2-191">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="bbca2-191">Contained in</span></span>

- [<span data-ttu-id="bbca2-192">ランタイム</span><span class="sxs-lookup"><span data-stu-id="bbca2-192">Runtime</span></span>](runtime.md)

### <a name="attributes"></a><span data-ttu-id="bbca2-193">属性</span><span class="sxs-lookup"><span data-stu-id="bbca2-193">Attributes</span></span>

|<span data-ttu-id="bbca2-194">属性</span><span class="sxs-lookup"><span data-stu-id="bbca2-194">Attribute</span></span>|<span data-ttu-id="bbca2-195">型</span><span class="sxs-lookup"><span data-stu-id="bbca2-195">Type</span></span>|<span data-ttu-id="bbca2-196">必須</span><span class="sxs-lookup"><span data-stu-id="bbca2-196">Required</span></span>|<span data-ttu-id="bbca2-197">説明</span><span class="sxs-lookup"><span data-stu-id="bbca2-197">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="bbca2-198">**type**</span><span class="sxs-lookup"><span data-stu-id="bbca2-198">**type**</span></span>|<span data-ttu-id="bbca2-199">string</span><span class="sxs-lookup"><span data-stu-id="bbca2-199">string</span></span>|<span data-ttu-id="bbca2-200">はい</span><span class="sxs-lookup"><span data-stu-id="bbca2-200">Yes</span></span>|<span data-ttu-id="bbca2-201">このオーバーライドの言語を指定します。</span><span class="sxs-lookup"><span data-stu-id="bbca2-201">Specifies the language for this override.</span></span> <span data-ttu-id="bbca2-202">現在、 `"javascript"` サポートされているオプションは唯一です。</span><span class="sxs-lookup"><span data-stu-id="bbca2-202">At present, `"javascript"` is the only supported option.</span></span>|
|<span data-ttu-id="bbca2-203">**resid**</span><span class="sxs-lookup"><span data-stu-id="bbca2-203">**resid**</span></span>|<span data-ttu-id="bbca2-204">文字列</span><span class="sxs-lookup"><span data-stu-id="bbca2-204">string</span></span>|<span data-ttu-id="bbca2-205">はい</span><span class="sxs-lookup"><span data-stu-id="bbca2-205">Yes</span></span>|<span data-ttu-id="bbca2-206">親 [Runtime](runtime.md) 要素で定義された既定の HTML の URL の場所をオーバーライドする JavaScript ファイルの URL の場所を指定 `resid` します。</span><span class="sxs-lookup"><span data-stu-id="bbca2-206">Specifies the URL location of the JavaScript file that should override the URL location of the default HTML defined in the parent [Runtime](runtime.md) element's `resid`.</span></span> <span data-ttu-id="bbca2-207">は `resid` 32 文字以内 `id` で、要素の属性と一致する必要があります `Url` `Resources` 。</span><span class="sxs-lookup"><span data-stu-id="bbca2-207">The `resid` can be no more than 32 characters and must match an `id` attribute of a `Url` element in the `Resources` element.</span></span>|

### <a name="examples"></a><span data-ttu-id="bbca2-208">例</span><span class="sxs-lookup"><span data-stu-id="bbca2-208">Examples</span></span>

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

### <a name="see-also"></a><span data-ttu-id="bbca2-209">関連項目</span><span class="sxs-lookup"><span data-stu-id="bbca2-209">See also</span></span>

- [<span data-ttu-id="bbca2-210">ランタイム</span><span class="sxs-lookup"><span data-stu-id="bbca2-210">Runtime</span></span>](runtime.md)
- [<span data-ttu-id="bbca2-211">イベント ベースのアクティブ化用にOutlook アドインを構成する</span><span class="sxs-lookup"><span data-stu-id="bbca2-211">Configure your Outlook add-in for event-based activation</span></span>](../../outlook/autolaunch.md)
