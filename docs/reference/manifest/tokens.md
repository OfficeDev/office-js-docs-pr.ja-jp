---
title: マニフェストファイルの Tokens 要素
description: マニフェスト内の URL テンプレートで使用できるトークンまたはワイルドカードを指定します。
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: a50de7c2c3e8ebeb9425c1677a94bbcc62281d3b
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996697"
---
# <a name="tokens-element"></a><span data-ttu-id="f2266-103">Tokens 要素</span><span class="sxs-lookup"><span data-stu-id="f2266-103">Tokens element</span></span>

<span data-ttu-id="f2266-104">テンプレート Url で使用できるトークンを定義します。</span><span class="sxs-lookup"><span data-stu-id="f2266-104">Defines tokens that could be used in template URLs.</span></span>

<span data-ttu-id="f2266-105">**アドインの種類:** 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="f2266-105">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="f2266-106">構文</span><span class="sxs-lookup"><span data-stu-id="f2266-106">Syntax</span></span>

```XML
<Tokens></Tokens>
```

## <a name="contained-in"></a><span data-ttu-id="f2266-107">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="f2266-107">Contained in</span></span>

[<span data-ttu-id="f2266-108">ExtendedOverrides</span><span class="sxs-lookup"><span data-stu-id="f2266-108">ExtendedOverrides</span></span>](extendedoverrides.md)

## <a name="must-contain"></a><span data-ttu-id="f2266-109">含める必要があるもの</span><span class="sxs-lookup"><span data-stu-id="f2266-109">Must contain</span></span>

|<span data-ttu-id="f2266-110">要素</span><span class="sxs-lookup"><span data-stu-id="f2266-110">Element</span></span>|<span data-ttu-id="f2266-111">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="f2266-111">Content</span></span>|<span data-ttu-id="f2266-112">メール</span><span class="sxs-lookup"><span data-stu-id="f2266-112">Mail</span></span>|<span data-ttu-id="f2266-113">TaskPane</span><span class="sxs-lookup"><span data-stu-id="f2266-113">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="f2266-114">トークン</span><span class="sxs-lookup"><span data-stu-id="f2266-114">Token</span></span>](token.md)|||<span data-ttu-id="f2266-115">x</span><span class="sxs-lookup"><span data-stu-id="f2266-115">x</span></span>|

## <a name="example"></a><span data-ttu-id="f2266-116">例</span><span class="sxs-lookup"><span data-stu-id="f2266-116">Example</span></span>

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