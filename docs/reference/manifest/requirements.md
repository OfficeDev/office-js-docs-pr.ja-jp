---
title: マニフェスト ファイルの Requirements 要素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 3c4cb81ebd6a38ea311e8fcacfa6d5fcd3b26f68
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325249"
---
# <a name="requirements-element"></a><span data-ttu-id="0e6d1-102">Requirements 要素</span><span class="sxs-lookup"><span data-stu-id="0e6d1-102">Requirements element</span></span>

<span data-ttu-id="0e6d1-103">Office アドインをアクティブにするために必要な Office JavaScript API の要件の最小セット ([要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets)またはメソッド、あるいはその両方) を指定します。</span><span class="sxs-lookup"><span data-stu-id="0e6d1-103">Specifies the minimum set of Office JavaScript API requirements ([requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="0e6d1-104">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="0e6d1-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="0e6d1-105">構文</span><span class="sxs-lookup"><span data-stu-id="0e6d1-105">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="0e6d1-106">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="0e6d1-106">Contained in</span></span>

[<span data-ttu-id="0e6d1-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="0e6d1-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="0e6d1-108">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="0e6d1-108">Can contain</span></span>

|<span data-ttu-id="0e6d1-109">**Element**</span><span class="sxs-lookup"><span data-stu-id="0e6d1-109">**Element**</span></span>|<span data-ttu-id="0e6d1-110">**コンテンツ**</span><span class="sxs-lookup"><span data-stu-id="0e6d1-110">**Content**</span></span>|<span data-ttu-id="0e6d1-111">**メール**</span><span class="sxs-lookup"><span data-stu-id="0e6d1-111">**Mail**</span></span>|<span data-ttu-id="0e6d1-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="0e6d1-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="0e6d1-113">Sets</span><span class="sxs-lookup"><span data-stu-id="0e6d1-113">Sets</span></span>](sets.md)|<span data-ttu-id="0e6d1-114">x</span><span class="sxs-lookup"><span data-stu-id="0e6d1-114">x</span></span>|<span data-ttu-id="0e6d1-115">x</span><span class="sxs-lookup"><span data-stu-id="0e6d1-115">x</span></span>|<span data-ttu-id="0e6d1-116">x</span><span class="sxs-lookup"><span data-stu-id="0e6d1-116">x</span></span>|
|[<span data-ttu-id="0e6d1-117">メソッド</span><span class="sxs-lookup"><span data-stu-id="0e6d1-117">Methods</span></span>](methods.md)|<span data-ttu-id="0e6d1-118">x</span><span class="sxs-lookup"><span data-stu-id="0e6d1-118">x</span></span>||<span data-ttu-id="0e6d1-119">x</span><span class="sxs-lookup"><span data-stu-id="0e6d1-119">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="0e6d1-120">解説</span><span class="sxs-lookup"><span data-stu-id="0e6d1-120">Remarks</span></span>

<span data-ttu-id="0e6d1-121">利用できる要件セットの詳細については、「[Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0e6d1-121">For more information about requirement sets, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

