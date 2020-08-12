---
title: マニフェスト ファイルの Requirements 要素
description: 要件要素は、Office アドインをアクティブにするために必要な最小要件セットとメソッドを指定します。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: c6a9a7b5923401fc2551f239b2c6cbc0d1e90755
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641320"
---
# <a name="requirements-element"></a><span data-ttu-id="b5e90-103">Requirements 要素</span><span class="sxs-lookup"><span data-stu-id="b5e90-103">Requirements element</span></span>

<span data-ttu-id="b5e90-104">Office アドインをアクティブにするために必要な Office JavaScript API の要件の最小セット ([要件セット](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets)またはメソッド、あるいはその両方) を指定します。</span><span class="sxs-lookup"><span data-stu-id="b5e90-104">Specifies the minimum set of Office JavaScript API requirements ([requirement sets](../../develop/office-versions-and-requirement-sets.md#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="b5e90-105">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="b5e90-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="b5e90-106">構文</span><span class="sxs-lookup"><span data-stu-id="b5e90-106">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="b5e90-107">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="b5e90-107">Contained in</span></span>

[<span data-ttu-id="b5e90-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="b5e90-108">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="b5e90-109">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="b5e90-109">Can contain</span></span>

|<span data-ttu-id="b5e90-110">要素</span><span class="sxs-lookup"><span data-stu-id="b5e90-110">Element</span></span>|<span data-ttu-id="b5e90-111">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="b5e90-111">Content</span></span>|<span data-ttu-id="b5e90-112">メール</span><span class="sxs-lookup"><span data-stu-id="b5e90-112">Mail</span></span>|<span data-ttu-id="b5e90-113">TaskPane</span><span class="sxs-lookup"><span data-stu-id="b5e90-113">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="b5e90-114">Sets</span><span class="sxs-lookup"><span data-stu-id="b5e90-114">Sets</span></span>](sets.md)|<span data-ttu-id="b5e90-115">x</span><span class="sxs-lookup"><span data-stu-id="b5e90-115">x</span></span>|<span data-ttu-id="b5e90-116">x</span><span class="sxs-lookup"><span data-stu-id="b5e90-116">x</span></span>|<span data-ttu-id="b5e90-117">x</span><span class="sxs-lookup"><span data-stu-id="b5e90-117">x</span></span>|
|[<span data-ttu-id="b5e90-118">メソッド</span><span class="sxs-lookup"><span data-stu-id="b5e90-118">Methods</span></span>](methods.md)|<span data-ttu-id="b5e90-119">x</span><span class="sxs-lookup"><span data-stu-id="b5e90-119">x</span></span>||<span data-ttu-id="b5e90-120">x</span><span class="sxs-lookup"><span data-stu-id="b5e90-120">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="b5e90-121">解説</span><span class="sxs-lookup"><span data-stu-id="b5e90-121">Remarks</span></span>

<span data-ttu-id="b5e90-122">利用できる要件セットの詳細については、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b5e90-122">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>
