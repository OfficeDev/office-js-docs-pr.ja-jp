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
# <a name="override-element"></a>Override 要素

指定した条件に応じてマニフェスト設定の値を上書きする方法を提供します。 条件には次の 2 種類があります。

- 既定Office異なるロケールを指定します。
- 既定のパターンとは異なる要件セットのサポートのパターン。

要素には、LocaleTokenOverride と呼ばれるロケールオーバーライド用の要素と `<Override>` **、RequirementTokenOverride** と呼ばれる要件セットのオーバーライド用の 2 種類があります。 ただし、要素 `type` のパラメーター `<Override>` はありません。 違いは、親要素と親要素の型によって決まります。 要素 `<Override>` の内部にある要素は `<Token>` `xsi:type` `RequirementToken` **、RequirementTokenOverride 型である必要があります**。 他 `<Override>` の親要素内の要素、または型の要素内の要素は `<Override>` `LocaleToken` **、LocaleTokenOverride 型である必要があります**。 各種類については、以下の各セクションで説明します。 要素の子である場合のこの要素の使用の詳細については、「マニフェストの拡張オーバーライドを処理する」 `<Token>` [を参照してください](../../develop/extended-overrides.md)。

## <a name="override-element-of-type-localetokenoverride"></a>LocaleTokenOverride 型のオーバーライド要素

要素 `<Override>` は条件付きを表し、"If .." として読み取り可能です。その後 ..."。ステートメント。 要素が `<Override>` **LocaleTokenOverride** 型の場合、属性は条件であり、その `Locale` `Value` 結果属性になります。 たとえば、次の例は、「Officeロケール設定が fr-fr の場合、表示名は 'Lecteur vidéo'です。

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
