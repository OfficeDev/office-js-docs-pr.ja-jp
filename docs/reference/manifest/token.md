---
title: マニフェスト ファイル内の Token 要素
description: マニフェストの URL テンプレートで使用できるトークンまたはワイルドカードを指定します。
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: 48078f8211a8fd3f0e3f9d7c3f3aabd1d31b0a6d
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505368"
---
# <a name="token-element"></a><span data-ttu-id="00ebc-103">Token 要素</span><span class="sxs-lookup"><span data-stu-id="00ebc-103">Token element</span></span>

<span data-ttu-id="00ebc-104">個々の URL トークンを定義します。</span><span class="sxs-lookup"><span data-stu-id="00ebc-104">Defines an individual URL token.</span></span> <span data-ttu-id="00ebc-105">この要素の使用の詳細については、「マニフェストの拡張オーバーライドを処理する [」を参照してください](../../develop/extended-overrides.md)。</span><span class="sxs-lookup"><span data-stu-id="00ebc-105">For more information about the use of this element, see [Work with extended overrides of the manifest](../../develop/extended-overrides.md).</span></span>

<span data-ttu-id="00ebc-106">**アドインの種類:** 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="00ebc-106">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="00ebc-107">構文</span><span class="sxs-lookup"><span data-stu-id="00ebc-107">Syntax</span></span>

```XML
<Token Name="string" DefaultValue="string" xsi:type=["LocaleToken" | "RequirementsToken"] ></Token>
```

## <a name="contained-in"></a><span data-ttu-id="00ebc-108">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="00ebc-108">Contained in</span></span>

[<span data-ttu-id="00ebc-109">トークン</span><span class="sxs-lookup"><span data-stu-id="00ebc-109">Tokens</span></span>](tokens.md)

## <a name="can-contain"></a><span data-ttu-id="00ebc-110">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="00ebc-110">Can contain</span></span>

|<span data-ttu-id="00ebc-111">要素</span><span class="sxs-lookup"><span data-stu-id="00ebc-111">Element</span></span>|<span data-ttu-id="00ebc-112">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="00ebc-112">Content</span></span>|<span data-ttu-id="00ebc-113">メール</span><span class="sxs-lookup"><span data-stu-id="00ebc-113">Mail</span></span>|<span data-ttu-id="00ebc-114">TaskPane</span><span class="sxs-lookup"><span data-stu-id="00ebc-114">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="00ebc-115">Override</span><span class="sxs-lookup"><span data-stu-id="00ebc-115">Override</span></span>](override.md)|||<span data-ttu-id="00ebc-116">x</span><span class="sxs-lookup"><span data-stu-id="00ebc-116">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="00ebc-117">属性</span><span class="sxs-lookup"><span data-stu-id="00ebc-117">Attributes</span></span>

|<span data-ttu-id="00ebc-118">属性</span><span class="sxs-lookup"><span data-stu-id="00ebc-118">Attribute</span></span>|<span data-ttu-id="00ebc-119">説明</span><span class="sxs-lookup"><span data-stu-id="00ebc-119">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="00ebc-120">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="00ebc-120">DefaultValue</span></span>|<span data-ttu-id="00ebc-121">子要素に条件が一致する場合、このトークン `<Override>` の既定値。</span><span class="sxs-lookup"><span data-stu-id="00ebc-121">Default value for this token if no condition in any child `<Override>` element matches.</span></span>|
|<span data-ttu-id="00ebc-122">名前</span><span class="sxs-lookup"><span data-stu-id="00ebc-122">Name</span></span>|<span data-ttu-id="00ebc-123">トークン名。</span><span class="sxs-lookup"><span data-stu-id="00ebc-123">Token name.</span></span> <span data-ttu-id="00ebc-124">この名前はユーザー定義です。</span><span class="sxs-lookup"><span data-stu-id="00ebc-124">This name is user-defined.</span></span> <span data-ttu-id="00ebc-125">トークンの種類は type 属性によって決まります。</span><span class="sxs-lookup"><span data-stu-id="00ebc-125">The type of the token is determined by the type attribute.</span></span>|
|<span data-ttu-id="00ebc-126">xsi:type</span><span class="sxs-lookup"><span data-stu-id="00ebc-126">xsi:type</span></span>|<span data-ttu-id="00ebc-127">トークンの種類を定義します。</span><span class="sxs-lookup"><span data-stu-id="00ebc-127">Defines the kind of Token.</span></span> <span data-ttu-id="00ebc-128">この属性は、次のいずれかの値に  `"RequirementsToken"` 設定する必要があります  `"LocaleToken"` 。</span><span class="sxs-lookup"><span data-stu-id="00ebc-128">This attribute should be set to one of:  `"RequirementsToken"`,  or  `"LocaleToken"`.</span></span>|

## <a name="example"></a><span data-ttu-id="00ebc-129">例</span><span class="sxs-lookup"><span data-stu-id="00ebc-129">Example</span></span>

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