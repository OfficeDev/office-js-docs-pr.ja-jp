---
title: マニフェスト ファイルの Sets 要素
description: Sets 要素は、office アドインをアクティブにするために必要な最低限の Office JavaScript API のセットを指定します。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: bd8f8311bb06a8e9e98fc408aece6395ab5643b1
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641425"
---
# <a name="sets-element"></a><span data-ttu-id="bf564-103">Sets 要素</span><span class="sxs-lookup"><span data-stu-id="bf564-103">Sets element</span></span>

<span data-ttu-id="bf564-104">Office アドインをアクティブにするために必要な最低限の Office JavaScript API のサブセットを指定します。</span><span class="sxs-lookup"><span data-stu-id="bf564-104">Specifies the minimum subset of the Office JavaScript API that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="bf564-105">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="bf564-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="bf564-106">構文</span><span class="sxs-lookup"><span data-stu-id="bf564-106">Syntax</span></span>

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a><span data-ttu-id="bf564-107">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="bf564-107">Contained in</span></span>

[<span data-ttu-id="bf564-108">Requirements</span><span class="sxs-lookup"><span data-stu-id="bf564-108">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="bf564-109">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="bf564-109">Can contain</span></span>

[<span data-ttu-id="bf564-110">Set</span><span class="sxs-lookup"><span data-stu-id="bf564-110">Set</span></span>](set.md)

## <a name="attributes"></a><span data-ttu-id="bf564-111">属性</span><span class="sxs-lookup"><span data-stu-id="bf564-111">Attributes</span></span>

|<span data-ttu-id="bf564-112">属性</span><span class="sxs-lookup"><span data-stu-id="bf564-112">Attribute</span></span>|<span data-ttu-id="bf564-113">型</span><span class="sxs-lookup"><span data-stu-id="bf564-113">Type</span></span>|<span data-ttu-id="bf564-114">必須</span><span class="sxs-lookup"><span data-stu-id="bf564-114">Required</span></span>|<span data-ttu-id="bf564-115">説明</span><span class="sxs-lookup"><span data-stu-id="bf564-115">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="bf564-116">DefaultMinVersion</span><span class="sxs-lookup"><span data-stu-id="bf564-116">DefaultMinVersion</span></span>|<span data-ttu-id="bf564-117">文字列</span><span class="sxs-lookup"><span data-stu-id="bf564-117">string</span></span>|<span data-ttu-id="bf564-118">省略可能</span><span class="sxs-lookup"><span data-stu-id="bf564-118">optional</span></span>|<span data-ttu-id="bf564-119">すべての子[セット](set.md)要素の既定の**MinVersion**属性値を指定します。</span><span class="sxs-lookup"><span data-stu-id="bf564-119">Specifies the default **MinVersion** attribute value for all child [Set](set.md) elements.</span></span> <span data-ttu-id="bf564-120">既定値は "1.1" です。</span><span class="sxs-lookup"><span data-stu-id="bf564-120">The default value is "1.1".</span></span>|

## <a name="remarks"></a><span data-ttu-id="bf564-121">解説</span><span class="sxs-lookup"><span data-stu-id="bf564-121">Remarks</span></span>

<span data-ttu-id="bf564-122">利用できる要件セットの詳細については、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="bf564-122">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="bf564-123">**Set**要素の**MinVersion**属性と**Sets**要素の**defaultminversion**属性の詳細については、「マニフェストの[要件要素を設定する](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="bf564-123">For more information about the **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span></span>

