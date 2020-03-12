---
title: マニフェスト ファイルの Requirements 要素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 43c66118b9129c4c8ae395254ea82ef1cbcbaab1
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596460"
---
# <a name="requirements-element"></a><span data-ttu-id="0539c-102">Requirements 要素</span><span class="sxs-lookup"><span data-stu-id="0539c-102">Requirements element</span></span>

<span data-ttu-id="0539c-103">Office アドインをアクティブにするために必要な Office JavaScript API の要件の最小セット ([要件セット](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets)またはメソッド、あるいはその両方) を指定します。</span><span class="sxs-lookup"><span data-stu-id="0539c-103">Specifies the minimum set of Office JavaScript API requirements ([requirement sets](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="0539c-104">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="0539c-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="0539c-105">構文</span><span class="sxs-lookup"><span data-stu-id="0539c-105">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="0539c-106">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="0539c-106">Contained in</span></span>

[<span data-ttu-id="0539c-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="0539c-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="0539c-108">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="0539c-108">Can contain</span></span>

|<span data-ttu-id="0539c-109">**Element**</span><span class="sxs-lookup"><span data-stu-id="0539c-109">**Element**</span></span>|<span data-ttu-id="0539c-110">**コンテンツ**</span><span class="sxs-lookup"><span data-stu-id="0539c-110">**Content**</span></span>|<span data-ttu-id="0539c-111">**メール**</span><span class="sxs-lookup"><span data-stu-id="0539c-111">**Mail**</span></span>|<span data-ttu-id="0539c-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="0539c-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="0539c-113">Sets</span><span class="sxs-lookup"><span data-stu-id="0539c-113">Sets</span></span>](sets.md)|<span data-ttu-id="0539c-114">x</span><span class="sxs-lookup"><span data-stu-id="0539c-114">x</span></span>|<span data-ttu-id="0539c-115">x</span><span class="sxs-lookup"><span data-stu-id="0539c-115">x</span></span>|<span data-ttu-id="0539c-116">x</span><span class="sxs-lookup"><span data-stu-id="0539c-116">x</span></span>|
|[<span data-ttu-id="0539c-117">メソッド</span><span class="sxs-lookup"><span data-stu-id="0539c-117">Methods</span></span>](methods.md)|<span data-ttu-id="0539c-118">x</span><span class="sxs-lookup"><span data-stu-id="0539c-118">x</span></span>||<span data-ttu-id="0539c-119">x</span><span class="sxs-lookup"><span data-stu-id="0539c-119">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="0539c-120">解説</span><span class="sxs-lookup"><span data-stu-id="0539c-120">Remarks</span></span>

<span data-ttu-id="0539c-121">利用できる要件セットの詳細については、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0539c-121">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>
