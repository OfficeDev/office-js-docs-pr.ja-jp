---
title: マニフェストファイルの Token 要素
description: マニフェスト内の URL テンプレートで使用できるトークンまたはワイルドカードを指定します。
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 5e26af44c566ab09ac81c8194e1ae7d85aaac327
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996703"
---
# <a name="token-element"></a>Token 要素

個別の URL トークンを定義します。

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
|DefaultValue|いずれかの子要素に一致する条件がない場合は、このトークンの既定値 `<Override>` 。|
|名前|トークン名。 この名前は、ユーザー定義です。 トークンの種類は、type 属性によって決まります。|
|xsi:type|トークンの種類を定義します。 この属性は  `"RequirementsToken"` 、、またはのいずれかに設定する必要があり  `"LocaleToken"` ます。|

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