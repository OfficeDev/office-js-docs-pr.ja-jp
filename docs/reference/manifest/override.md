---
title: マニフェスト ファイルの Override 要素
description: Override 要素を使用すると、追加のロケールの設定値を指定できます。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 139a4089a36d8a8adfa71d4a0947b02f5b163b52
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641453"
---
# <a name="override-element"></a><span data-ttu-id="a86f2-103">Override 要素</span><span class="sxs-lookup"><span data-stu-id="a86f2-103">Override element</span></span>

<span data-ttu-id="a86f2-104">追加ロケールの設定の値を指定する方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="a86f2-104">Provides a way to specify the value of a setting for an additional locale.</span></span>

<span data-ttu-id="a86f2-105">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="a86f2-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="a86f2-106">構文</span><span class="sxs-lookup"><span data-stu-id="a86f2-106">Syntax</span></span>

```XML
<Override Locale="string" Value="string" />
```

## <a name="contained-in"></a><span data-ttu-id="a86f2-107">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="a86f2-107">Contained in</span></span>

|<span data-ttu-id="a86f2-108">要素</span><span class="sxs-lookup"><span data-stu-id="a86f2-108">Element</span></span>|
|:-----|
|[<span data-ttu-id="a86f2-109">CitationText</span><span class="sxs-lookup"><span data-stu-id="a86f2-109">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="a86f2-110">説明</span><span class="sxs-lookup"><span data-stu-id="a86f2-110">Description</span></span>](description.md)|
|[<span data-ttu-id="a86f2-111">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="a86f2-111">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="a86f2-112">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="a86f2-112">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="a86f2-113">DisplayName</span><span class="sxs-lookup"><span data-stu-id="a86f2-113">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="a86f2-114">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="a86f2-114">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="a86f2-115">IconUrl</span><span class="sxs-lookup"><span data-stu-id="a86f2-115">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="a86f2-116">QueryUri</span><span class="sxs-lookup"><span data-stu-id="a86f2-116">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="a86f2-117">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="a86f2-117">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="a86f2-118">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="a86f2-118">SupportUrl</span></span>](supporturl.md)|

## <a name="attributes"></a><span data-ttu-id="a86f2-119">属性</span><span class="sxs-lookup"><span data-stu-id="a86f2-119">Attributes</span></span>

|<span data-ttu-id="a86f2-120">属性</span><span class="sxs-lookup"><span data-stu-id="a86f2-120">Attribute</span></span>|<span data-ttu-id="a86f2-121">型</span><span class="sxs-lookup"><span data-stu-id="a86f2-121">Type</span></span>|<span data-ttu-id="a86f2-122">必須</span><span class="sxs-lookup"><span data-stu-id="a86f2-122">Required</span></span>|<span data-ttu-id="a86f2-123">説明</span><span class="sxs-lookup"><span data-stu-id="a86f2-123">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="a86f2-124">Locale</span><span class="sxs-lookup"><span data-stu-id="a86f2-124">Locale</span></span>|<span data-ttu-id="a86f2-125">string</span><span class="sxs-lookup"><span data-stu-id="a86f2-125">string</span></span>|<span data-ttu-id="a86f2-126">必須</span><span class="sxs-lookup"><span data-stu-id="a86f2-126">required</span></span>|<span data-ttu-id="a86f2-127">`"en-US"` などの BCP 47 言語タグの書式で、この上書きのロケールのカルチャ名を指定します。</span><span class="sxs-lookup"><span data-stu-id="a86f2-127">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="a86f2-128">Value</span><span class="sxs-lookup"><span data-stu-id="a86f2-128">Value</span></span>|<span data-ttu-id="a86f2-129">string</span><span class="sxs-lookup"><span data-stu-id="a86f2-129">string</span></span>|<span data-ttu-id="a86f2-130">必須</span><span class="sxs-lookup"><span data-stu-id="a86f2-130">required</span></span>|<span data-ttu-id="a86f2-131">指定のロケールに対して表される設定の値を指定します。</span><span class="sxs-lookup"><span data-stu-id="a86f2-131">Specifies value of the setting expressed for the specified locale.</span></span>|

## <a name="see-also"></a><span data-ttu-id="a86f2-132">関連項目</span><span class="sxs-lookup"><span data-stu-id="a86f2-132">See also</span></span>

- [<span data-ttu-id="a86f2-133">Office アドインのローカライズ</span><span class="sxs-lookup"><span data-stu-id="a86f2-133">Localization for Office Add-ins</span></span>](../../develop/localization.md)
