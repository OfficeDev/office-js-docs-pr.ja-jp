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
# <a name="override-element"></a>Override 要素

指定した条件に応じて、マニフェストの設定値を上書きする方法を提供します。 条件には、次の2種類があります。

- 既定とは異なる Office ロケール。
- 既定のパターンとは異なる、要件セットサポートのパターン。

要素には2つの種類があり `<Override>` ます。1つは **LocaleTokenOverride** と呼ばれるロケールの上書き用で、もう1つは要件セットのオーバーライド ( **RequirementTokenOverride** と呼ばれる) です。 ただし `type` 、要素のパラメーターはありません `<Override>` 。 相違点は、親要素と親要素の型によって決まります。 がである要素 `<Override>` の内部にある要素は `<Token>` `xsi:type` `RequirementToken` 、 **RequirementTokenOverride** 型である必要があります。 `<Override>`他の親要素の中、または型の要素内の要素は `<Override>` `LocaleToken` 、 **LocaleTokenOverride** 型でなければなりません。 それぞれの種類について、以下の個別のセクションで説明します。

## <a name="override-element-of-type-localetokenoverride"></a>LocaleTokenOverride 型の Override 要素

`<Override>`要素は条件を表し、"If...[...]if. 要素の `<Override>` 型が **LocaleTokenOverride** の場合は、 `Locale` 属性は条件です。属性はその後のものです `Value` 。 たとえば、"Office ロケール設定が fr-fr で、表示名が ' Lecteur vidéo ' の場合は、次の値が読み取られます。

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

## <a name="override-element-of-type-requirementtokenoverride"></a>RequirementTokenOverride 型の Override 要素

`<Override>`要素は条件を表し、"If...[...]if. 要素の `<Override>` 型が **RequirementTokenOverride** の場合、子要素は `<Requirements>` 条件を表し、属性はその後の `Value` ものです。 たとえば、 `<Override>` 現在のプラットフォームが FeatureOne version 1.7 をサポートしている場合は、次のように "oldAddinVersion" を使用します。これは、 `${token.requirements}` 既定の文字列 ' upgrade ' ではなく、祖父母の URL に含まれるトークンの代わりに使用され `<ExtendedOverrides>` ます。

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
