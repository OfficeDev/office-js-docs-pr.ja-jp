---
title: マニフェスト ファイル内の Tokens 要素
description: マニフェストの URL テンプレートで使用できるトークンまたはワイルドカードを指定します。
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 5d42abab46ecc6e7ab465144f061d26da52c0eb3e2623acd8a8a2912ecc13312
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57095788"
---
# <a name="tokens-element"></a>Tokens 要素

テンプレート URL で使用できるトークンを定義します。 この要素の使用の詳細については、「マニフェストの拡張オーバーライドを処理する [」を参照してください](../../develop/extended-overrides.md)。

**アドインの種類:** 作業ウィンドウ

## <a name="syntax"></a>構文

```XML
<Tokens></Tokens>
```

## <a name="contained-in"></a>含まれる場所

[ExtendedOverrides](extendedoverrides.md)

## <a name="must-contain"></a>含める必要があるもの

|要素|コンテンツ|メール|TaskPane|
|:-----|:-----|:-----|:-----|
|[トークン](token.md)|||x|

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