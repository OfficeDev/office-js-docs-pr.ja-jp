---
title: マニフェスト ファイルの Override 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: d1d2400312f12116b1ac5f4010135541e783dcc7
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432866"
---
# <a name="override-element"></a><span data-ttu-id="f3e16-102">Override 要素</span><span class="sxs-lookup"><span data-stu-id="f3e16-102">Override element</span></span>

<span data-ttu-id="f3e16-103">追加ロケールの設定の値を指定する方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="f3e16-103">Provides a way to specify the value of a setting for an additional locale.</span></span>

<span data-ttu-id="f3e16-104">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="f3e16-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="f3e16-105">構文</span><span class="sxs-lookup"><span data-stu-id="f3e16-105">Syntax</span></span>

```XML
<Override Locale="string" Value="string" />
```

## <a name="contained-in"></a><span data-ttu-id="f3e16-106">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="f3e16-106">Contained in</span></span>

|<span data-ttu-id="f3e16-107">**要素**</span><span class="sxs-lookup"><span data-stu-id="f3e16-107">**Element**</span></span>|
|:-----|
|[<span data-ttu-id="f3e16-108">CitationText</span><span class="sxs-lookup"><span data-stu-id="f3e16-108">CitationText</span></span>](citationtext.md)|
|[<span data-ttu-id="f3e16-109">Description</span><span class="sxs-lookup"><span data-stu-id="f3e16-109">Description</span></span>](description.md)|
|[<span data-ttu-id="f3e16-110">DictionaryName</span><span class="sxs-lookup"><span data-stu-id="f3e16-110">DictionaryName</span></span>](dictionaryname.md)|
|[<span data-ttu-id="f3e16-111">DictionaryHomePage</span><span class="sxs-lookup"><span data-stu-id="f3e16-111">DictionaryHomePage</span></span>](dictionaryhomepage.md)|
|[<span data-ttu-id="f3e16-112">DisplayName</span><span class="sxs-lookup"><span data-stu-id="f3e16-112">DisplayName</span></span>](displayname.md)|
|[<span data-ttu-id="f3e16-113">HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="f3e16-113">HighResolutionIconUrl</span></span>](highresolutioniconurl.md)|
|[<span data-ttu-id="f3e16-114">IconUrl</span><span class="sxs-lookup"><span data-stu-id="f3e16-114">IconUrl</span></span>](iconurl.md)|
|[<span data-ttu-id="f3e16-115">QueryUri</span><span class="sxs-lookup"><span data-stu-id="f3e16-115">QueryUri</span></span>](queryuri.md)|
|[<span data-ttu-id="f3e16-116">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="f3e16-116">SourceLocation</span></span>](sourcelocation.md)|
|[<span data-ttu-id="f3e16-117">SupportUrl</span><span class="sxs-lookup"><span data-stu-id="f3e16-117">SupportUrl</span></span>](supporturl.md)|

## <a name="attributes"></a><span data-ttu-id="f3e16-118">属性</span><span class="sxs-lookup"><span data-stu-id="f3e16-118">Attributes</span></span>

|<span data-ttu-id="f3e16-119">**属性**</span><span class="sxs-lookup"><span data-stu-id="f3e16-119">**Attribute**</span></span>|<span data-ttu-id="f3e16-120">**型**</span><span class="sxs-lookup"><span data-stu-id="f3e16-120">**Type**</span></span>|<span data-ttu-id="f3e16-121">**必須**</span><span class="sxs-lookup"><span data-stu-id="f3e16-121">**Required**</span></span>|<span data-ttu-id="f3e16-122">**説明**</span><span class="sxs-lookup"><span data-stu-id="f3e16-122">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="f3e16-123">Locale</span><span class="sxs-lookup"><span data-stu-id="f3e16-123">Locale</span></span>|<span data-ttu-id="f3e16-124">string</span><span class="sxs-lookup"><span data-stu-id="f3e16-124">string</span></span>|<span data-ttu-id="f3e16-125">必須</span><span class="sxs-lookup"><span data-stu-id="f3e16-125">required</span></span>|<span data-ttu-id="f3e16-126">`"en-US"` などの BCP 47 言語タグの書式で、この上書きのロケールのカルチャ名を指定します。</span><span class="sxs-lookup"><span data-stu-id="f3e16-126">Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.</span></span>|
|<span data-ttu-id="f3e16-127">Value</span><span class="sxs-lookup"><span data-stu-id="f3e16-127">Value</span></span>|<span data-ttu-id="f3e16-128">string</span><span class="sxs-lookup"><span data-stu-id="f3e16-128">string</span></span>|<span data-ttu-id="f3e16-129">必須</span><span class="sxs-lookup"><span data-stu-id="f3e16-129">required</span></span>|<span data-ttu-id="f3e16-130">指定のロケールに対して表される設定の値を指定します。</span><span class="sxs-lookup"><span data-stu-id="f3e16-130">Specifies value of the setting expressed for the specified locale.</span></span>|

## <a name="see-also"></a><span data-ttu-id="f3e16-131">関連項目</span><span class="sxs-lookup"><span data-stu-id="f3e16-131">See also</span></span>

- [<span data-ttu-id="f3e16-132">Office アドインのローカライズ</span><span class="sxs-lookup"><span data-stu-id="f3e16-132">Localization for Office Add-ins</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/localization)
    
