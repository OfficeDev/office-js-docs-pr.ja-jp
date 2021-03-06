---
title: マニフェスト ファイル内の Tokens 要素
description: マニフェストの URL テンプレートで使用できるトークンまたはワイルドカードを指定します。
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 8680b985068c44e93f601a2b24e2f28899eb483d
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505326"
---
# <a name="tokens-element"></a><span data-ttu-id="fc68d-103">Tokens 要素</span><span class="sxs-lookup"><span data-stu-id="fc68d-103">Tokens element</span></span>

<span data-ttu-id="fc68d-104">テンプレート URL で使用できるトークンを定義します。</span><span class="sxs-lookup"><span data-stu-id="fc68d-104">Defines tokens that could be used in template URLs.</span></span> <span data-ttu-id="fc68d-105">この要素の使用の詳細については、「マニフェストの拡張オーバーライドを処理する [」を参照してください](../../develop/extended-overrides.md)。</span><span class="sxs-lookup"><span data-stu-id="fc68d-105">For more information about the use of this element, see [Work with extended overrides of the manifest](../../develop/extended-overrides.md).</span></span>

<span data-ttu-id="fc68d-106">**アドインの種類:** 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="fc68d-106">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="fc68d-107">構文</span><span class="sxs-lookup"><span data-stu-id="fc68d-107">Syntax</span></span>

```XML
<Tokens></Tokens>
```

## <a name="contained-in"></a><span data-ttu-id="fc68d-108">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="fc68d-108">Contained in</span></span>

[<span data-ttu-id="fc68d-109">ExtendedOverrides</span><span class="sxs-lookup"><span data-stu-id="fc68d-109">ExtendedOverrides</span></span>](extendedoverrides.md)

## <a name="must-contain"></a><span data-ttu-id="fc68d-110">含める必要があるもの</span><span class="sxs-lookup"><span data-stu-id="fc68d-110">Must contain</span></span>

|<span data-ttu-id="fc68d-111">要素</span><span class="sxs-lookup"><span data-stu-id="fc68d-111">Element</span></span>|<span data-ttu-id="fc68d-112">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="fc68d-112">Content</span></span>|<span data-ttu-id="fc68d-113">メール</span><span class="sxs-lookup"><span data-stu-id="fc68d-113">Mail</span></span>|<span data-ttu-id="fc68d-114">TaskPane</span><span class="sxs-lookup"><span data-stu-id="fc68d-114">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="fc68d-115">トークン</span><span class="sxs-lookup"><span data-stu-id="fc68d-115">Token</span></span>](token.md)|||<span data-ttu-id="fc68d-116">x</span><span class="sxs-lookup"><span data-stu-id="fc68d-116">x</span></span>|

## <a name="example"></a><span data-ttu-id="fc68d-117">例</span><span class="sxs-lookup"><span data-stu-id="fc68d-117">Example</span></span>

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