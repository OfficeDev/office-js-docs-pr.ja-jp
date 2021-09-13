---
title: マニフェスト ファイル内の Token 要素
description: マニフェストの URL テンプレートで使用できるトークンまたはワイルドカードを指定します。
ms.date: 11/06/2020
ms.localizationpriority: medium
ms.openlocfilehash: 69f626f5f6f57dd155756812bcd56267a1da3ffa
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151206"
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