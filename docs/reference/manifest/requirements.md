---
title: マニフェスト ファイルの Requirements 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 2544e9b01b2d4d3ddc0a0c6238b4a5b0e6c4f832
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432705"
---
# <a name="requirements-element"></a><span data-ttu-id="06930-102">Requirements 要素</span><span class="sxs-lookup"><span data-stu-id="06930-102">Requirements element</span></span>

<span data-ttu-id="06930-103">Office アドインをアクティブにするために必要な JavaScript API for Office の最小要件セット ([要件セット](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets)またはメソッド、あるいはその両方) を指定します。</span><span class="sxs-lookup"><span data-stu-id="06930-103">Specifies the minimum set of JavaScript API for Office requirements ([requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="06930-104">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="06930-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="06930-105">構文</span><span class="sxs-lookup"><span data-stu-id="06930-105">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="06930-106">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="06930-106">Contained in</span></span>

[<span data-ttu-id="06930-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="06930-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="06930-108">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="06930-108">Can contain</span></span>

|<span data-ttu-id="06930-109">**要素**</span><span class="sxs-lookup"><span data-stu-id="06930-109">**Element**</span></span>|<span data-ttu-id="06930-110">**Content**</span><span class="sxs-lookup"><span data-stu-id="06930-110">**Content**</span></span>|<span data-ttu-id="06930-111">**Mail**</span><span class="sxs-lookup"><span data-stu-id="06930-111">**Mail**</span></span>|<span data-ttu-id="06930-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="06930-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="06930-113">Sets</span><span class="sxs-lookup"><span data-stu-id="06930-113">Sets</span></span>](sets.md)|<span data-ttu-id="06930-114">x</span><span class="sxs-lookup"><span data-stu-id="06930-114">x</span></span>|<span data-ttu-id="06930-115">x</span><span class="sxs-lookup"><span data-stu-id="06930-115">x</span></span>|<span data-ttu-id="06930-116">x</span><span class="sxs-lookup"><span data-stu-id="06930-116">x</span></span>|
|[<span data-ttu-id="06930-117">Methods</span><span class="sxs-lookup"><span data-stu-id="06930-117">Methods</span></span>](methods.md)|<span data-ttu-id="06930-118">x</span><span class="sxs-lookup"><span data-stu-id="06930-118">x</span></span>||<span data-ttu-id="06930-119">x</span><span class="sxs-lookup"><span data-stu-id="06930-119">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="06930-120">解説</span><span class="sxs-lookup"><span data-stu-id="06930-120">Remarks</span></span>

<span data-ttu-id="06930-121">利用できる要件セットの詳細については、「[Office のバージョンと要件セット](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="06930-121">For more information about available requirement sets, see [Office add-in requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

