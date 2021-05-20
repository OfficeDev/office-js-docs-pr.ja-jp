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
# <a name="override-element"></a>Override 要素

指定した条件に応じて、マニフェスト設定の値をオーバーライドする方法を提供します。 条件には、次の 3 種類があります。

- 既定のロケールとは異なるOffice ロケールです `LocaleToken` 。 
- 要件セットのサポートのパターンは、既定のパターンとは異なり `RequirementToken` 、 **要件TokenOverride** と呼ばれます。
- ソースは `Runtime` **、RuntimeOverride** (現在プレビュー中) と呼ばれる既定のとは異なります。

`<Override>`要素の内部にある要素は、 `<Runtime>` 型が **RuntimeOverride** である必要があります。

`overrideType`要素の属性がありません `<Override>` 。 違いは、親要素と親要素の型によって決まります。 である `<Override>` 要素の内部にある `<Token>` 要素 `xsi:type` は `RequirementToken` 、型 **が "要件TokenOverride"** である必要があります。 `<Override>`他の親要素の中、または `<Override>` 型の要素内の要素 `LocaleToken` は **、型が LocaleTokenOverride** である必要があります。 要素の子である場合にこの要素を使用する方法の詳細については `<Token>` 、「 マニフェストの [拡張オーバーライドを処理する](../../develop/extended-overrides.md)」を参照してください。

各型については、この記事の後半で説明します。

## <a name="override-element-for-localetoken"></a>要素をオーバーライドする `LocaleToken`

`<Override>`要素は条件付きを表し、"If..その後..陳述。 要素の `<Override>` 型が **LocaleTokenOverride** の場合、 `Locale` 属性は条件であり、 `Value` 属性は結果です。 たとえば、次の例は、「Officeロケール設定が fr-fr の場合、表示名は 'レクトゥール vidéo' です。

```xml
<DisplayName DefaultValue="Video player">
    <Override Locale="fr-fr" Value="Lecteur vidéo" />
</DisplayName>
```

**アドインの種類:** コンテンツ、作業ウィンドウ、メール

### <a name="syntax"></a>構文

```XML
<Override Locale="string" Value="string"></Override>
```

### <a name="contained-in"></a>含まれる場所

|要素|
|:-----|
|[CitationText](citationtext.md)|
|[説明](description.md)|
|[DictionaryName](dictionaryname.md)|
|[DictionaryHomePage](dictionaryhomepage.md)|
|[DisplayName](displayname.md)|
|[HighResolutionIconUrl](highresolutioniconurl.md)|
|[IconUrl](iconurl.md)|
|[QueryUri](queryuri.md)|
|[SourceLocation](sourcelocation.md)|
|[SupportUrl](supporturl.md)|
|[トークン](token.md)|

### <a name="attributes"></a>属性

|属性|型|必須|説明|
|:-----|:-----|:-----|:-----|
|Locale|string|必須|`"en-US"` などの BCP 47 言語タグの書式で、この上書きのロケールのカルチャ名を指定します。|
|Value|string|必須|指定のロケールに対して表される設定の値を指定します。|

### <a name="examples"></a>例

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

### <a name="see-also"></a>関連項目

- [Office アドインのローカライズ](../../develop/localization.md)
- [SharePoint のキーボード ショートカット](../../design/keyboard-shortcuts.md)

## <a name="override-element-for-requirementtoken"></a>要素をオーバーライドする `RequirementToken`

`<Override>`要素は条件付きを表し、"If..その後..陳述。 要素の `<Override>` 型が **"要件TokenOverride"** の場合、子 `<Requirements>` 要素は条件を表し、 `Value` 属性は結果です。 たとえば、 `<Override>` 次の最初の例は、「現在のプラットフォームが FeatureOne バージョン 1.7 をサポートしている場合は `${token.requirements}` 、( `<ExtendedOverrides>` 既定の文字列 'upgrade'の代わりに) (既定の文字列 'upgrade') の代わりに、トークンの代わりに文字列 'oldAddinVersion' を使用します。

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

**アドインの種類:** 作業ウィンドウ

### <a name="syntax"></a>構文

```XML
<Override Value="string" />
```

### <a name="contained-in"></a>含まれる場所

|要素|
|:-----|
|[トークン](token.md)|

### <a name="must-contain"></a>含める必要があるもの

|要素|コンテンツ|メール|TaskPane|
|:-----|:-----|:-----|:-----|
|[Requirements](requirements.md)|||x|

### <a name="attributes"></a>属性

|属性|型|必須|説明|
|:-----|:-----|:-----|:-----|
|値|string|必須|条件が満たされた場合の、祖父母トークンの値。|

### <a name="example"></a>例

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

### <a name="see-also"></a>関連項目

- [Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)
- [マニフェストで Requirements 要素を設定する](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)
- [SharePoint のキーボード ショートカット](../../design/keyboard-shortcuts.md)

## <a name="override-element-for-runtime-preview"></a>要素を上書き `Runtime` (プレビュー)

> [!IMPORTANT]
> この機能は、web 上のOutlookで[プレビュー](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md)し、Microsoft 365サブスクリプションでWindowsにのみサポートされます。 詳細については、「[イベントベースのアクティブ化用にOutlook アドインを構成する](../../outlook/autolaunch.md)」を参照してください。
>
> プレビュー機能は予告なく変更される場合があるため、運用アドインで使用しないでください。

`<Override>`要素は条件付きを表し、"If..その後..陳述。 要素の `<Override>` 型が **RuntimeOverride** の場合、 `type` 属性は条件であり、 `resid` 属性は結果です。 たとえば、次の例は"型が 'javascript' の場合は `resid` 'JSRuntime.Url' です。Outlookデスクトップには[、LaunchEvent 拡張機能ポイント](../../reference/manifest/extensionpoint.md#launchevent-preview)ハンドラーにこの要素が必要です。

```xml
<Runtime resid="WebViewRuntime.Url">
  <Override type="javascript" resid="JSRuntime.Url"/>
</Runtime>
```

**アドインの種類:** メール

### <a name="syntax"></a>構文

```XML
<Override type="javascript" resid="JSRuntime.Url"/>
```

### <a name="contained-in"></a>含まれる場所

- [ランタイム](runtime.md)

### <a name="attributes"></a>属性

|属性|型|必須|説明|
|:-----|:-----|:-----|:-----|
|**type**|string|はい|このオーバーライドの言語を指定します。 現在、 `"javascript"` サポートされているオプションは唯一です。|
|**resid**|文字列|はい|親 [Runtime](runtime.md) 要素で定義された既定の HTML の URL の場所をオーバーライドする JavaScript ファイルの URL の場所を指定 `resid` します。 は `resid` 32 文字以内 `id` で、要素の属性と一致する必要があります `Url` `Resources` 。|

### <a name="examples"></a>例

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

### <a name="see-also"></a>関連項目

- [ランタイム](runtime.md)
- [イベント ベースのアクティブ化用にOutlook アドインを構成する](../../outlook/autolaunch.md)
