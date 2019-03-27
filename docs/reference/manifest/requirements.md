---
title: マニフェスト ファイルの Requirements 要素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 364ab7c943895e1acecedba7970e54da331a2e6f
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870367"
---
# <a name="requirements-element"></a><span data-ttu-id="721d8-102">Requirements 要素</span><span class="sxs-lookup"><span data-stu-id="721d8-102">Requirements element</span></span>

<span data-ttu-id="721d8-103">Office アドインをアクティブにするために必要な JavaScript API for Office の最小要件セット ([要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets)またはメソッド、あるいはその両方) を指定します。</span><span class="sxs-lookup"><span data-stu-id="721d8-103">Specifies the minimum set of JavaScript API for Office requirements ([requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets#specify-office-hosts-and-requirement-sets) and/or methods) that your Office Add-in needs to activate.</span></span>

<span data-ttu-id="721d8-104">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="721d8-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="721d8-105">構文</span><span class="sxs-lookup"><span data-stu-id="721d8-105">Syntax</span></span>

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a><span data-ttu-id="721d8-106">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="721d8-106">Contained in</span></span>

[<span data-ttu-id="721d8-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="721d8-107">OfficeApp</span></span>](officeapp.md)

## <a name="can-contain"></a><span data-ttu-id="721d8-108">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="721d8-108">Can contain</span></span>

|<span data-ttu-id="721d8-109">**Element**</span><span class="sxs-lookup"><span data-stu-id="721d8-109">**Element**</span></span>|<span data-ttu-id="721d8-110">**コンテンツ**</span><span class="sxs-lookup"><span data-stu-id="721d8-110">**Content**</span></span>|<span data-ttu-id="721d8-111">**メール**</span><span class="sxs-lookup"><span data-stu-id="721d8-111">**Mail**</span></span>|<span data-ttu-id="721d8-112">**TaskPane**</span><span class="sxs-lookup"><span data-stu-id="721d8-112">**TaskPane**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="721d8-113">Sets</span><span class="sxs-lookup"><span data-stu-id="721d8-113">Sets</span></span>](sets.md)|<span data-ttu-id="721d8-114">x</span><span class="sxs-lookup"><span data-stu-id="721d8-114">x</span></span>|<span data-ttu-id="721d8-115">x</span><span class="sxs-lookup"><span data-stu-id="721d8-115">x</span></span>|<span data-ttu-id="721d8-116">x</span><span class="sxs-lookup"><span data-stu-id="721d8-116">x</span></span>|
|[<span data-ttu-id="721d8-117">メソッド</span><span class="sxs-lookup"><span data-stu-id="721d8-117">Methods</span></span>](methods.md)|<span data-ttu-id="721d8-118">x</span><span class="sxs-lookup"><span data-stu-id="721d8-118">x</span></span>||<span data-ttu-id="721d8-119">x</span><span class="sxs-lookup"><span data-stu-id="721d8-119">x</span></span>|

## <a name="remarks"></a><span data-ttu-id="721d8-120">解説</span><span class="sxs-lookup"><span data-stu-id="721d8-120">Remarks</span></span>

<span data-ttu-id="721d8-121">利用できる要件セットの詳細については、「[Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="721d8-121">For more information about requirement sets, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

