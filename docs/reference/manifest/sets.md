---
title: マニフェスト ファイルの Sets 要素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 768f674b4afbd65df88825e871005f182d06f6ce
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325242"
---
# <a name="sets-element"></a><span data-ttu-id="2e30a-102">Sets 要素</span><span class="sxs-lookup"><span data-stu-id="2e30a-102">Sets element</span></span>

<span data-ttu-id="2e30a-103">Office アドインをアクティブにするために必要な最低限の Office JavaScript API のサブセットを指定します。</span><span class="sxs-lookup"><span data-stu-id="2e30a-103">Specifies the minimum subset of the Office JavaScript API that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="2e30a-104">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="2e30a-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="2e30a-105">構文</span><span class="sxs-lookup"><span data-stu-id="2e30a-105">Syntax</span></span>

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a><span data-ttu-id="2e30a-106">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="2e30a-106">Contained in</span></span>

[<span data-ttu-id="2e30a-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="2e30a-107">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="2e30a-108">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="2e30a-108">Can contain</span></span>

[<span data-ttu-id="2e30a-109">Set</span><span class="sxs-lookup"><span data-stu-id="2e30a-109">Set</span></span>](set.md)

## <a name="attributes"></a><span data-ttu-id="2e30a-110">属性</span><span class="sxs-lookup"><span data-stu-id="2e30a-110">Attributes</span></span>

|<span data-ttu-id="2e30a-111">**属性**</span><span class="sxs-lookup"><span data-stu-id="2e30a-111">**Attribute**</span></span>|<span data-ttu-id="2e30a-112">**型**</span><span class="sxs-lookup"><span data-stu-id="2e30a-112">**Type**</span></span>|<span data-ttu-id="2e30a-113">**必須**</span><span class="sxs-lookup"><span data-stu-id="2e30a-113">**Required**</span></span>|<span data-ttu-id="2e30a-114">**説明**</span><span class="sxs-lookup"><span data-stu-id="2e30a-114">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="2e30a-115">DefaultMinVersion</span><span class="sxs-lookup"><span data-stu-id="2e30a-115">DefaultMinVersion</span></span>|<span data-ttu-id="2e30a-116">文字列</span><span class="sxs-lookup"><span data-stu-id="2e30a-116">string</span></span>|<span data-ttu-id="2e30a-117">省略可能</span><span class="sxs-lookup"><span data-stu-id="2e30a-117">optional</span></span>|<span data-ttu-id="2e30a-118">すべての子[セット](set.md)要素の既定の**MinVersion**属性値を指定します。</span><span class="sxs-lookup"><span data-stu-id="2e30a-118">Specifies the default **MinVersion** attribute value for all child [Set](set.md) elements.</span></span> <span data-ttu-id="2e30a-119">既定値は "1.1" です。</span><span class="sxs-lookup"><span data-stu-id="2e30a-119">The default value is "1.1".</span></span>|

## <a name="remarks"></a><span data-ttu-id="2e30a-120">解説</span><span class="sxs-lookup"><span data-stu-id="2e30a-120">Remarks</span></span>

<span data-ttu-id="2e30a-121">利用できる要件セットの詳細については、「[Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2e30a-121">For more information about requirement sets, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="2e30a-122">**Set**要素の**MinVersion**属性と**Sets**要素の**defaultminversion**属性の詳細については、「マニフェストの[要件要素を設定する](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="2e30a-122">For more information about the **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

