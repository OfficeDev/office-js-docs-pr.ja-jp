---
title: マニフェスト ファイル内の Token 要素
description: マニフェストの URL テンプレートで使用できるトークンまたはワイルドカードを指定します。
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 867bb5bc801b85b63c7815debfaf59c5cee3a8157dc866ba7082803ee1d7fe2a
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57095944"
---
# <a name="token-element"></a>Token 要素

個々の URL トークンを定義します。 この要素の使用の詳細については、「マニフェストの拡張オーバーライドを処理する [」を参照してください](../../develop/extended-overrides.md)。

**アドインの種類:** 作業ウィンドウ

## <a name="syntax"></a>構文

```XML
<Token Name="string" DefaultValue="string" xsi:type=["LocaleToken" | "RequirementsToken"] ></Token>
```

## <a name="contained-in"></a>含まれる場所

[トークン](tokens.md)

## <a name="can-contain"></a>含めることができるもの

|要素|コンテンツ|メール|TaskPane|
|:-----|:-----|:-----|:-----|
|[Override](override.md)|||x|

## <a name="attributes"></a>属性

|属性|説明|
|:-----|:-----|
|DefaultValue|子要素に条件が一致する場合、このトークン `<Override>` の既定値。|
|名前|トークン名。 この名前はユーザー定義です。 トークンの種類は type 属性によって決まります。|
|xsi:type|トークンの種類を定義します。 この属性は、次のいずれかの値に  `"RequirementsToken"` 設定する必要があります  `"LocaleToken"` 。|

## <a name="example"></a>例

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