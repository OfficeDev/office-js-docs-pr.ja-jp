---
title: マニフェストファイルの ExtendedOverrides 要素
description: マニフェストの JSON 形式の拡張機能の Url を指定します。
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 76491af34d1caf0ec266826df97a5363e336b85d
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996709"
---
# <a name="extendedoverrides-element"></a>ExtendedOverrides 要素

マニフェストを拡張する JSON 形式のファイルの完全な Url を指定します。

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
|Url (必須)| 拡張オーバーライド JSON ファイルの完全な URL。 これは、 [token](tokens.md) 要素によって定義されたトークンを使用する URL テンプレートである場合があります。|
|ResourcesUrl (省略可能) | 属性で指定されているファイルについて、ローカライズされた文字列などの補足情報を提供するファイルの完全な URL `Url` 。 これは、 [token](tokens.md) 要素によって定義されたトークンを使用する URL テンプレートである場合があります。|

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
