---
title: マニフェスト ファイルの Requirements 要素
description: 要件要素は、Office アドインをアクティブにするために必要な最小要件セットとメソッドを指定します。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: a3f41a763ec820a6c766e6a32b26e55ad34996f7
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720450"
---
# <a name="requirements-element"></a><span data-ttu-id="4a004-103">Requirements 要素</span><span class="sxs-lookup"><span data-stu-id="4a004-103">Requirements element</span></span>

<span data-ttu-id="4a004-104">Office アドインをアクティブにするために必要な Office JavaScript API の要件の最小セット ([要件セット](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets)またはメソッド、あるいはその両方) を指定します。</span><span class="sxs-lookup"><span data-stu-id="4a004-104">Specifies the minimum set of Office JavaScript API requirements ([requirement sets](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="4a004-105">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="4a004-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="4a004-106">構文</span><span class="sxs-lookup"><span data-stu-id="4a004-106">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="4a004-107">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="4a004-107">Contained in</span></span>

[<span data-ttu-id="4a004-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="4a004-108">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="4a004-109">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="4a004-109">Can contain</span></span>

|<span data-ttu-id="4a004-110">**Element**</span><span class="sxs-lookup"><span data-stu-id="4a004-110">**Element**</span></span>|<span data-ttu-id="4a004-111">**コンテンツ**</span><span class="sxs-lookup"><span data-stu-id="4a004-111">**Content**</span></span>|<span data-ttu-id="4a004-112">**メール**</span><span class="sxs-lookup"><span data-stu-id="4a004-112">**Mail**</span></span>|<span data-ttu-id="4a004-113">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="4a004-113">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="4a004-114">Sets</span><span class="sxs-lookup"><span data-stu-id="4a004-114">Sets</span></span>](sets.md)|<span data-ttu-id="4a004-115">x</span><span class="sxs-lookup"><span data-stu-id="4a004-115">x</span></span>|<span data-ttu-id="4a004-116">x</span><span class="sxs-lookup"><span data-stu-id="4a004-116">x</span></span>|<span data-ttu-id="4a004-117">x</span><span class="sxs-lookup"><span data-stu-id="4a004-117">x</span></span>|
|[<span data-ttu-id="4a004-118">メソッド</span><span class="sxs-lookup"><span data-stu-id="4a004-118">Methods</span></span>](methods.md)|<span data-ttu-id="4a004-119">x</span><span class="sxs-lookup"><span data-stu-id="4a004-119">x</span></span>||<span data-ttu-id="4a004-120">x</span><span class="sxs-lookup"><span data-stu-id="4a004-120">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="4a004-121">解説</span><span class="sxs-lookup"><span data-stu-id="4a004-121">Remarks</span></span>

<span data-ttu-id="4a004-122">利用できる要件セットの詳細については、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="4a004-122">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>
