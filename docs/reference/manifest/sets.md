---
title: マニフェスト ファイルの Sets 要素
description: Sets 要素は、office アドインをアクティブにするために必要な最低限の Office JavaScript API のセットを指定します。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 8c1c97bfc2934ecf3cc20b472b29a03805603729
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608734"
---
# <a name="sets-element"></a><span data-ttu-id="bafb3-103">Sets 要素</span><span class="sxs-lookup"><span data-stu-id="bafb3-103">Sets element</span></span>

<span data-ttu-id="bafb3-104">Office アドインをアクティブにするために必要な最低限の Office JavaScript API のサブセットを指定します。</span><span class="sxs-lookup"><span data-stu-id="bafb3-104">Specifies the minimum subset of the Office JavaScript API that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="bafb3-105">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="bafb3-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="bafb3-106">構文</span><span class="sxs-lookup"><span data-stu-id="bafb3-106">Syntax</span></span>

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a><span data-ttu-id="bafb3-107">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="bafb3-107">Contained in</span></span>

[<span data-ttu-id="bafb3-108">Requirements</span><span class="sxs-lookup"><span data-stu-id="bafb3-108">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="bafb3-109">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="bafb3-109">Can contain</span></span>

[<span data-ttu-id="bafb3-110">Set</span><span class="sxs-lookup"><span data-stu-id="bafb3-110">Set</span></span>](set.md)

## <a name="attributes"></a><span data-ttu-id="bafb3-111">属性</span><span class="sxs-lookup"><span data-stu-id="bafb3-111">Attributes</span></span>

|<span data-ttu-id="bafb3-112">**属性**</span><span class="sxs-lookup"><span data-stu-id="bafb3-112">**Attribute**</span></span>|<span data-ttu-id="bafb3-113">**型**</span><span class="sxs-lookup"><span data-stu-id="bafb3-113">**Type**</span></span>|<span data-ttu-id="bafb3-114">**必須**</span><span class="sxs-lookup"><span data-stu-id="bafb3-114">**Required**</span></span>|<span data-ttu-id="bafb3-115">**説明**</span><span class="sxs-lookup"><span data-stu-id="bafb3-115">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="bafb3-116">DefaultMinVersion</span><span class="sxs-lookup"><span data-stu-id="bafb3-116">DefaultMinVersion</span></span>|<span data-ttu-id="bafb3-117">文字列</span><span class="sxs-lookup"><span data-stu-id="bafb3-117">string</span></span>|<span data-ttu-id="bafb3-118">省略可能</span><span class="sxs-lookup"><span data-stu-id="bafb3-118">optional</span></span>|<span data-ttu-id="bafb3-119">すべての子[セット](set.md)要素の既定の**MinVersion**属性値を指定します。</span><span class="sxs-lookup"><span data-stu-id="bafb3-119">Specifies the default **MinVersion** attribute value for all child [Set](set.md) elements.</span></span> <span data-ttu-id="bafb3-120">既定値は "1.1" です。</span><span class="sxs-lookup"><span data-stu-id="bafb3-120">The default value is "1.1".</span></span>|

## <a name="remarks"></a><span data-ttu-id="bafb3-121">解説</span><span class="sxs-lookup"><span data-stu-id="bafb3-121">Remarks</span></span>

<span data-ttu-id="bafb3-122">利用できる要件セットの詳細については、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="bafb3-122">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="bafb3-123">**Set**要素の**MinVersion**属性と**Sets**要素の**defaultminversion**属性の詳細については、「マニフェストの[要件要素を設定する](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="bafb3-123">For more information about the **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span></span>

