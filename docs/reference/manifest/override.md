---
title: マニフェスト ファイルの Override 要素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 020ae490dacbb9b8c493dc022c23d0ebf311a1b9
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450451"
---
# <a name="override-element"></a><span data-ttu-id="e74ec-102">Override 要素</span><span class="sxs-lookup"><span data-stu-id="e74ec-102">Override element</span></span>

<span data-ttu-id="e74ec-103">追加ロケールの設定の値を指定する方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="e74ec-103">Provides a way to specify the value of a setting for an additional locale.</span></span>

<span data-ttu-id="e74ec-104">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="e74ec-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="e74ec-105">構文</span><span class="sxs-lookup"><span data-stu-id="e74ec-105">Syntax</span></span>

```XML
<Override Locale="string" Value="string" />
```

## <a name="contained-in"></a><span data-ttu-id="e74ec-106">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="e74ec-106">Contained in</span></span>

|<span data-ttu-id="e74ec-107">**要素**</span><span class="sxs-lookup"><span data-stu-id="e74ec-107">**Element**</span></span>|
|:-----|
|[<span data-ttu-id="e74ec-108">CitationText</span><span class="sxs-lookup"><span data-stu-id="e74ec-108">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="e74ec-109">説明</span><span class="sxs-lookup"><span data-stu-id="e74ec-109">Description</span></span>](description.md)|
|[<span data-ttu-id="e74ec-110">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="e74ec-110">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="e74ec-111">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="e74ec-111">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="e74ec-112">DisplayName</span><span class="sxs-lookup"><span data-stu-id="e74ec-112">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="e74ec-113">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="e74ec-113">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="e74ec-114">IconUrl</span><span class="sxs-lookup"><span data-stu-id="e74ec-114">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="e74ec-115">QueryUri</span><span class="sxs-lookup"><span data-stu-id="e74ec-115">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="e74ec-116">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="e74ec-116">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="e74ec-117">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="e74ec-117">SupportUrl</span></span>](supporturl.md)|

## <a name="attributes"></a><span data-ttu-id="e74ec-118">属性</span><span class="sxs-lookup"><span data-stu-id="e74ec-118">Attributes</span></span>

|<span data-ttu-id="e74ec-119">**属性**</span><span class="sxs-lookup"><span data-stu-id="e74ec-119">**Attribute**</span></span>|<span data-ttu-id="e74ec-120">**型**</span><span class="sxs-lookup"><span data-stu-id="e74ec-120">**Type**</span></span>|<span data-ttu-id="e74ec-121">**必須**</span><span class="sxs-lookup"><span data-stu-id="e74ec-121">**Required**</span></span>|<span data-ttu-id="e74ec-122">**説明**</span><span class="sxs-lookup"><span data-stu-id="e74ec-122">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="e74ec-123">Locale</span><span class="sxs-lookup"><span data-stu-id="e74ec-123">Locale</span></span>|<span data-ttu-id="e74ec-124">string</span><span class="sxs-lookup"><span data-stu-id="e74ec-124">string</span></span>|<span data-ttu-id="e74ec-125">必須</span><span class="sxs-lookup"><span data-stu-id="e74ec-125">required</span></span>|<span data-ttu-id="e74ec-126">`"en-US"` などの BCP 47 言語タグの書式で、この上書きのロケールのカルチャ名を指定します。</span><span class="sxs-lookup"><span data-stu-id="e74ec-126">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="e74ec-127">Value</span><span class="sxs-lookup"><span data-stu-id="e74ec-127">Value</span></span>|<span data-ttu-id="e74ec-128">string</span><span class="sxs-lookup"><span data-stu-id="e74ec-128">string</span></span>|<span data-ttu-id="e74ec-129">必須</span><span class="sxs-lookup"><span data-stu-id="e74ec-129">required</span></span>|<span data-ttu-id="e74ec-130">指定のロケールに対して表される設定の値を指定します。</span><span class="sxs-lookup"><span data-stu-id="e74ec-130">Specifies value of the setting expressed for the specified locale.</span></span>|

## <a name="see-also"></a><span data-ttu-id="e74ec-131">関連項目</span><span class="sxs-lookup"><span data-stu-id="e74ec-131">See also</span></span>

- [<span data-ttu-id="e74ec-132">Office アドインのローカライズ</span><span class="sxs-lookup"><span data-stu-id="e74ec-132">Localization for Office Add-ins</span></span>](/office/dev/add-ins/develop/localization)
    
