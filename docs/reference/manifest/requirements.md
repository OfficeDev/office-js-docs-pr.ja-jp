---
title: マニフェスト ファイルの Requirements 要素
description: 要件要素は、Office アドインをアクティブにするために必要な最小要件セットとメソッドを指定します。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 319ddc59901c524ed1cee580a81cff749ad570db
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292273"
---
# <a name="requirements-element"></a><span data-ttu-id="3df88-103">Requirements 要素</span><span class="sxs-lookup"><span data-stu-id="3df88-103">Requirements element</span></span>

<span data-ttu-id="3df88-104">Office アドインをアクティブにするために必要な Office JavaScript API の要件の最小セット ([要件セット](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets) またはメソッド、あるいはその両方) を指定します。</span><span class="sxs-lookup"><span data-stu-id="3df88-104">Specifies the minimum set of Office JavaScript API requirements ([requirement sets](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="3df88-105">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="3df88-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="3df88-106">構文</span><span class="sxs-lookup"><span data-stu-id="3df88-106">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="3df88-107">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="3df88-107">Contained in</span></span>

[<span data-ttu-id="3df88-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="3df88-108">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="3df88-109">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="3df88-109">Can contain</span></span>

|<span data-ttu-id="3df88-110">要素</span><span class="sxs-lookup"><span data-stu-id="3df88-110">Element</span></span>|<span data-ttu-id="3df88-111">コンテンツ</span><span class="sxs-lookup"><span data-stu-id="3df88-111">Content</span></span>|<span data-ttu-id="3df88-112">メール</span><span class="sxs-lookup"><span data-stu-id="3df88-112">Mail</span></span>|<span data-ttu-id="3df88-113">TaskPane</span><span class="sxs-lookup"><span data-stu-id="3df88-113">TaskPane</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="3df88-114">Sets</span><span class="sxs-lookup"><span data-stu-id="3df88-114">Sets</span></span>](sets.md)|<span data-ttu-id="3df88-115">x</span><span class="sxs-lookup"><span data-stu-id="3df88-115">x</span></span>|<span data-ttu-id="3df88-116">x</span><span class="sxs-lookup"><span data-stu-id="3df88-116">x</span></span>|<span data-ttu-id="3df88-117">x</span><span class="sxs-lookup"><span data-stu-id="3df88-117">x</span></span>|
|[<span data-ttu-id="3df88-118">メソッド</span><span class="sxs-lookup"><span data-stu-id="3df88-118">Methods</span></span>](methods.md)|<span data-ttu-id="3df88-119">x</span><span class="sxs-lookup"><span data-stu-id="3df88-119">x</span></span>||<span data-ttu-id="3df88-120">x</span><span class="sxs-lookup"><span data-stu-id="3df88-120">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="3df88-121">解説</span><span class="sxs-lookup"><span data-stu-id="3df88-121">Remarks</span></span>

<span data-ttu-id="3df88-122">利用できる要件セットの詳細については、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3df88-122">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>
