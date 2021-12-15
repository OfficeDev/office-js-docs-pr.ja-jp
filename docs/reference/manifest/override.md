---
title: マニフェスト ファイルの Override 要素
description: Override 要素を使用すると、指定した条件に応じて設定の値を指定できます。
ms.date: 12/13/2021
ms.localizationpriority: medium
ms.openlocfilehash: dda8f6ca5aee1492c51960fc637d96e4d82796cb
ms.sourcegitcommit: e44a8109d9323aea42ace643e11717fb49f40baa
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/15/2021
ms.locfileid: "61513999"
---
# <a name="override-element"></a>Override 要素

指定した条件に応じてマニフェスト設定の値を上書きする方法を提供します。 条件には次の 3 種類があります。

- LocaleTokenOverride と呼ばれる既定のロケールとは異なるOfficeロケールです `LocaleToken` 。 
- RequirementTokenOverride と呼ばれる既定のパターンとは異なる、要件セット `RequirementToken` **のサポートのパターン** です。
- ソースは `Runtime` 、RuntimeOverride と呼ばれる既定 **のソースとは異なります**。

要素 `<Override>` の内部にある要素は `<Runtime>` **、RuntimeOverride 型である必要があります**。

要素の `overrideType` 属性 `<Override>` はありません。 違いは、親要素と親要素の型によって決まります。 要素 `<Override>` の内部にある要素は `<Token>` `xsi:type` `RequirementToken` **、RequirementTokenOverride 型である必要があります**。 他 `<Override>` の親要素内の要素、または型の要素内の要素は `<Override>` `LocaleToken` **、LocaleTokenOverride 型である必要があります**。 要素の子である場合のこの要素の使用の詳細については、「マニフェストの拡張オーバーライドを処理する」 `<Token>` [を参照してください](../../develop/extended-overrides.md)。

各種類については、この記事で後述する個別のセクションで説明します。

## <a name="override-element-for-localetoken"></a>Override 要素 `LocaleToken`

要素 `<Override>` は条件付きを表し、"If .." として読み取り可能です。その後 ..."。ステートメント。 要素が `<Override>` **LocaleTokenOverride** 型の場合、属性は条件であり、その `Locale` `Value` 結果属性になります。 たとえば、次の例では、「ロケールOffice fr-fr の場合、表示名は 'Lecteur vidéo'です。

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

|属性|種類|必須|説明|
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

## <a name="override-element-for-requirementtoken"></a>Override 要素 `RequirementToken`

要素 `<Override>` は条件付きを表し、"If .." として読み取り可能です。その後 ..."。ステートメント。 要素が `<Override>` **RequirementTokenOverride** 型の場合、子要素は条件を表し、属性 `<Requirements>` `Value` はその結果です。 たとえば、次の 1 つ目は、「現在のプラットフォームが FeatureOne バージョン 1.7 をサポートしている場合は、(既定の文字列 'upgrade' ではなく) 祖父母の URL のトークンの代わりに文字列 `<Override>` 'oldAddinVersion' を使用します。 `${token.requirements}` `<ExtendedOverrides>`

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

|属性|種類|必須|説明|
|:-----|:-----|:-----|:-----|
|値|string|必須|条件が満たされた場合の祖父母トークンの値。|

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

## <a name="override-element-for-runtime"></a>Override 要素 `Runtime`

> [!IMPORTANT]
> この要素のサポートは、イベント ベースのアクティブ化機能を備えたメールボックス要件 [セット 1.10](../../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md) [で導入されました](../../outlook/autolaunch.md)。 この要件セットをサポートする [クライアントおよびプラットフォーム](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) を参照してください。

要素 `<Override>` は条件付きを表し、"If .." として読み取り可能です。その後 ..."。ステートメント。 要素が RuntimeOverride 型の場合、属性は条件であり、属性 `<Override>`  `type` `resid` は結果です。 たとえば、「型が 'javascript'の場合は `resid` 、'JSRuntime.Url'です」と読み取ります。Outlookデスクトップでは、LaunchEvent 拡張ポイント ハンドラー[に対してこの要素が](../../reference/manifest/extensionpoint.md#launchevent)必要です。

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

- [Runtime](runtime.md)

### <a name="attributes"></a>属性

|属性|種類|必須|説明|
|:-----|:-----|:-----|:-----|
|**type**|string|はい|このオーバーライドの言語を指定します。 現時点では、 `"javascript"` サポートされている唯一のオプションです。|
|**resid**|文字列|はい|親 [ランタイム](runtime.md) 要素で定義されている既定の HTML の URL の場所を上書きする JavaScript ファイルの URL の場所を指定します `resid` 。 32 文字以内で、要素内の要素の属性と一致 `resid` `id` `Url` する必要 `Resources` があります。|

### <a name="examples"></a>例

```xml
<!-- Event-based activation happens in a lightweight runtime.-->
<Runtimes>
  <!-- HTML file including reference to or inline JavaScript event handlers.
  This is used by Outlook on the web and Outlook on the new Mac UI preview. -->
  <Runtime resid="WebViewRuntime.Url">
    <!-- JavaScript file containing event handlers. This is used by Outlook Desktop. -->
    <Override type="javascript" resid="JSRuntime.Url"/>
  </Runtime>
</Runtimes>
```

### <a name="see-also"></a>関連項目

- [Runtime](runtime.md)
- [イベント ベースのOutlook用にアドインを構成する](../../outlook/autolaunch.md)
