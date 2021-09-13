---
title: マニフェスト ファイルの ExtendedOverrides 要素
description: マニフェストの JSON 形式の拡張子の URL を指定します。
ms.date: 02/23/2021
ms.localizationpriority: medium
ms.openlocfilehash: 062d62284d786e81f230b1d094c46a73208a6c2c
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59153034"
---
# <a name="extendedoverrides-element"></a>ExtendedOverrides 要素

マニフェストを拡張する JSON 形式のファイルの完全な URL を指定します。 この要素とその子孫要素の使用の詳細については、「マニフェストの拡張オーバーライドを処理する」 [を参照してください](../../develop/extended-overrides.md)。

**アドインの種類:** 作業ウィンドウ

## <a name="syntax"></a>構文

```XML
<ExtendedOverrides Url="string" [ResourcesUrl="string"] ></ExtendedOverrides>
```

## <a name="contained-in"></a>含まれる場所

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>含めることができるもの

|要素|コンテンツ|メール|TaskPane|
|:-----|:-----|:-----|:-----|
|[トークン](tokens.md)|||x|

## <a name="attributes"></a>属性

|属性|説明|
|:-----|:-----|
|URL (必須)| 拡張の完全な URL は JSON ファイルを上書きします。 将来、この値は、Tokens 要素で定義されたトークンを使用する URL テンプレート [である可能性](tokens.md) があります。 「 [例」を参照してください](#examples)。|
|ResourcesUrl (オプション) | 属性で指定されたファイルの、ローカライズされた文字列などの補足リソースを提供するファイルの完全な `Url` URL。 これは、Tokens 要素で定義されたトークンを使用する URL テンプレート [である可能性](tokens.md) があります。|

## <a name="examples"></a>例

```XML
<OfficeApp ...>
  <!-- other elements omitted -->
  <ExtendedOverrides Url="http://contoso.com/addinmetadata/extended-manifest-overrides.json"
                     ResourceUrl="https://contoso.com/addin/my-resources.json">
  </ExtendedOverrides>
</OfficeApp>
```

将来、この値は、Tokens 要素で定義されたトークンを使用する URL テンプレート [である可能性](tokens.md) があります。 次に例を示します。

```XML
<OfficeApp ...>
  <!-- other elements omitted -->
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
</OfficeApp>
```
