---
title: マニフェスト ファイルの Sets 要素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 80f8a74b64186496ac1579b283b3e2976978328b
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596488"
---
# <a name="sets-element"></a><span data-ttu-id="5bfc5-102">Sets 要素</span><span class="sxs-lookup"><span data-stu-id="5bfc5-102">Sets element</span></span>

<span data-ttu-id="5bfc5-103">Office アドインをアクティブにするために必要な最低限の Office JavaScript API のサブセットを指定します。</span><span class="sxs-lookup"><span data-stu-id="5bfc5-103">Specifies the minimum subset of the Office JavaScript API that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="5bfc5-104">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="5bfc5-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="5bfc5-105">構文</span><span class="sxs-lookup"><span data-stu-id="5bfc5-105">Syntax</span></span>

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a><span data-ttu-id="5bfc5-106">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="5bfc5-106">Contained in</span></span>

[<span data-ttu-id="5bfc5-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="5bfc5-107">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="5bfc5-108">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="5bfc5-108">Can contain</span></span>

[<span data-ttu-id="5bfc5-109">Set</span><span class="sxs-lookup"><span data-stu-id="5bfc5-109">Set</span></span>](set.md)

## <a name="attributes"></a><span data-ttu-id="5bfc5-110">属性</span><span class="sxs-lookup"><span data-stu-id="5bfc5-110">Attributes</span></span>

|<span data-ttu-id="5bfc5-111">**属性**</span><span class="sxs-lookup"><span data-stu-id="5bfc5-111">**Attribute**</span></span>|<span data-ttu-id="5bfc5-112">**型**</span><span class="sxs-lookup"><span data-stu-id="5bfc5-112">**Type**</span></span>|<span data-ttu-id="5bfc5-113">**必須**</span><span class="sxs-lookup"><span data-stu-id="5bfc5-113">**Required**</span></span>|<span data-ttu-id="5bfc5-114">**説明**</span><span class="sxs-lookup"><span data-stu-id="5bfc5-114">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="5bfc5-115">DefaultMinVersion</span><span class="sxs-lookup"><span data-stu-id="5bfc5-115">DefaultMinVersion</span></span>|<span data-ttu-id="5bfc5-116">文字列</span><span class="sxs-lookup"><span data-stu-id="5bfc5-116">string</span></span>|<span data-ttu-id="5bfc5-117">省略可能</span><span class="sxs-lookup"><span data-stu-id="5bfc5-117">optional</span></span>|<span data-ttu-id="5bfc5-118">すべての子[セット](set.md)要素の既定の**MinVersion**属性値を指定します。</span><span class="sxs-lookup"><span data-stu-id="5bfc5-118">Specifies the default **MinVersion** attribute value for all child [Set](set.md) elements.</span></span> <span data-ttu-id="5bfc5-119">既定値は "1.1" です。</span><span class="sxs-lookup"><span data-stu-id="5bfc5-119">The default value is "1.1".</span></span>|

## <a name="remarks"></a><span data-ttu-id="5bfc5-120">解説</span><span class="sxs-lookup"><span data-stu-id="5bfc5-120">Remarks</span></span>

<span data-ttu-id="5bfc5-121">利用できる要件セットの詳細については、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5bfc5-121">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="5bfc5-122">**Set**要素の**MinVersion**属性と**Sets**要素の**defaultminversion**属性の詳細については、「マニフェストの[要件要素を設定する](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="5bfc5-122">For more information about the **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span></span>

