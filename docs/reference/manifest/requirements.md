---
title: マニフェスト ファイルの Requirements 要素
description: 要件要素は、Office アドインをアクティブにするために必要な最小要件セットとメソッドを指定します。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 586f05ec68257462cb64a96abf2a34eb31861a5c
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611716"
---
# <a name="requirements-element"></a><span data-ttu-id="d904a-103">Requirements 要素</span><span class="sxs-lookup"><span data-stu-id="d904a-103">Requirements element</span></span>

<span data-ttu-id="d904a-104">Office アドインをアクティブにするために必要な Office JavaScript API の要件の最小セット ([要件セット](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets)またはメソッド、あるいはその両方) を指定します。</span><span class="sxs-lookup"><span data-stu-id="d904a-104">Specifies the minimum set of Office JavaScript API requirements ([requirement sets](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="d904a-105">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="d904a-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="d904a-106">構文</span><span class="sxs-lookup"><span data-stu-id="d904a-106">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="d904a-107">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="d904a-107">Contained in</span></span>

[<span data-ttu-id="d904a-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="d904a-108">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="d904a-109">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="d904a-109">Can contain</span></span>

|<span data-ttu-id="d904a-110">**Element**</span><span class="sxs-lookup"><span data-stu-id="d904a-110">**Element**</span></span>|<span data-ttu-id="d904a-111">**コンテンツ**</span><span class="sxs-lookup"><span data-stu-id="d904a-111">**Content**</span></span>|<span data-ttu-id="d904a-112">**メール**</span><span class="sxs-lookup"><span data-stu-id="d904a-112">**Mail**</span></span>|<span data-ttu-id="d904a-113">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="d904a-113">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="d904a-114">Sets</span><span class="sxs-lookup"><span data-stu-id="d904a-114">Sets</span></span>](sets.md)|<span data-ttu-id="d904a-115">x</span><span class="sxs-lookup"><span data-stu-id="d904a-115">x</span></span>|<span data-ttu-id="d904a-116">x</span><span class="sxs-lookup"><span data-stu-id="d904a-116">x</span></span>|<span data-ttu-id="d904a-117">x</span><span class="sxs-lookup"><span data-stu-id="d904a-117">x</span></span>|
|[<span data-ttu-id="d904a-118">メソッド</span><span class="sxs-lookup"><span data-stu-id="d904a-118">Methods</span></span>](methods.md)|<span data-ttu-id="d904a-119">x</span><span class="sxs-lookup"><span data-stu-id="d904a-119">x</span></span>||<span data-ttu-id="d904a-120">x</span><span class="sxs-lookup"><span data-stu-id="d904a-120">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="d904a-121">解説</span><span class="sxs-lookup"><span data-stu-id="d904a-121">Remarks</span></span>

<span data-ttu-id="d904a-122">利用できる要件セットの詳細については、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="d904a-122">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>
