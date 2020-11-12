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
# <a name="token-element"></a><span data-ttu-id="8b63f-103">Token 要素</span><span class="sxs-lookup"><span data-stu-id="8b63f-103">Token element</span></span>

<span data-ttu-id="8b63f-104">個別の URL トークンを定義します。</span><span class="sxs-lookup"><span data-stu-id="8b63f-104">Defines an individual URL token.</span></span>

<span data-ttu-id="8b63f-105">**アドインの種類:** 作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="8b63f-105">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="8b63f-106">構文</span><span class="sxs-lookup"><span data-stu-id="8b63f-106">Syntax</span></span>

```XML
<Token Name="string" DefaultValue="string" xsi:type=["LocaleToken" | "RequirementsToken"] ></Token>
```

## <a name="contained-in"></a><span data-ttu-id="8b63f-107">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="8b63f-107">Contained in</span></span>

[<span data-ttu-id="8b63f-108">トークン</span><span class="sxs-lookup"><span data-stu-id="8b63f-108">Tokens</span></span>](tokens.md)

## <a name="can-contain"></a><span data-ttu-id="8b63f-109">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="8b63f-109">Can contain</span></span>

|<span data-ttu-id="8b63f-110">要素</span><span class="sxs-lookup"><span data-stu-id="8b63f-110">Element</span></span>|<span data-ttu-id="8b63f-111">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="8b63f-111">Content</span></span>|<span data-ttu-id="8b63f-112">メール</span><span class="sxs-lookup"><span data-stu-id="8b63f-112">Mail</span></span>|<span data-ttu-id="8b63f-113">TaskPane</span><span class="sxs-lookup"><span data-stu-id="8b63f-113">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="8b63f-114">Override</span><span class="sxs-lookup"><span data-stu-id="8b63f-114">Override</span></span>](override.md)|||<span data-ttu-id="8b63f-115">x</span><span class="sxs-lookup"><span data-stu-id="8b63f-115">x</span></span>|

## <a name="attributes"></a><span data-ttu-id="8b63f-116">属性</span><span class="sxs-lookup"><span data-stu-id="8b63f-116">Attributes</span></span>

|<span data-ttu-id="8b63f-117">属性</span><span class="sxs-lookup"><span data-stu-id="8b63f-117">Attribute</span></span>|<span data-ttu-id="8b63f-118">説明</span><span class="sxs-lookup"><span data-stu-id="8b63f-118">Description</span></span>|
|:-----|:-----|
|<span data-ttu-id="8b63f-119">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="8b63f-119">DefaultValue</span></span>|<span data-ttu-id="8b63f-120">いずれかの子要素に一致する条件がない場合は、このトークンの既定値 `<Override>` 。</span><span class="sxs-lookup"><span data-stu-id="8b63f-120">Default value for this token if no condition in any child `<Override>` element matches.</span></span>|
|<span data-ttu-id="8b63f-121">名前</span><span class="sxs-lookup"><span data-stu-id="8b63f-121">Name</span></span>|<span data-ttu-id="8b63f-122">トークン名。</span><span class="sxs-lookup"><span data-stu-id="8b63f-122">Token name.</span></span> <span data-ttu-id="8b63f-123">この名前は、ユーザー定義です。</span><span class="sxs-lookup"><span data-stu-id="8b63f-123">This name is user-defined.</span></span> <span data-ttu-id="8b63f-124">トークンの種類は、type 属性によって決まります。</span><span class="sxs-lookup"><span data-stu-id="8b63f-124">The type of the token is determined by the type attribute.</span></span>|
|<span data-ttu-id="8b63f-125">xsi:type</span><span class="sxs-lookup"><span data-stu-id="8b63f-125">xsi:type</span></span>|<span data-ttu-id="8b63f-126">トークンの種類を定義します。</span><span class="sxs-lookup"><span data-stu-id="8b63f-126">Defines the kind of Token.</span></span> <span data-ttu-id="8b63f-127">この属性は  `"RequirementsToken"` 、、またはのいずれかに設定する必要があり  `"LocaleToken"` ます。</span><span class="sxs-lookup"><span data-stu-id="8b63f-127">This attribute should be set to one of:  `"RequirementsToken"`,  or  `"LocaleToken"`.</span></span>|

## <a name="example"></a><span data-ttu-id="8b63f-128">例</span><span class="sxs-lookup"><span data-stu-id="8b63f-128">Example</span></span>

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