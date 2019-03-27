---
title: マニフェスト ファイルの Sets 要素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 13777e54ec6bd2d97fa35609ebe194ed85ffa1b8
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871774"
---
# <a name="sets-element"></a><span data-ttu-id="0361e-102">Sets 要素</span><span class="sxs-lookup"><span data-stu-id="0361e-102">Sets element</span></span>

<span data-ttu-id="0361e-103">Office アドインをアクティブにするために必要な JavaScript API for Office の最小限のサブセットを指定します。</span><span class="sxs-lookup"><span data-stu-id="0361e-103">Specifies the minimum subset of the JavaScript API for Office that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="0361e-104">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="0361e-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="0361e-105">構文</span><span class="sxs-lookup"><span data-stu-id="0361e-105">Syntax</span></span>

```XML
<Sets DefaultMinVersion="n .n ">
   ...
</Sets>
```

## <a name="contained-in"></a><span data-ttu-id="0361e-106">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="0361e-106">Contained in</span></span>

[<span data-ttu-id="0361e-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="0361e-107">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="0361e-108">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="0361e-108">Can contain</span></span>

[<span data-ttu-id="0361e-109">Set</span><span class="sxs-lookup"><span data-stu-id="0361e-109">Set</span></span>](set.md)

## <a name="attributes"></a><span data-ttu-id="0361e-110">属性</span><span class="sxs-lookup"><span data-stu-id="0361e-110">Attributes</span></span>

|<span data-ttu-id="0361e-111">**属性**</span><span class="sxs-lookup"><span data-stu-id="0361e-111">**Attribute**</span></span>|<span data-ttu-id="0361e-112">**型**</span><span class="sxs-lookup"><span data-stu-id="0361e-112">**Type**</span></span>|<span data-ttu-id="0361e-113">**必須**</span><span class="sxs-lookup"><span data-stu-id="0361e-113">**Required**</span></span>|<span data-ttu-id="0361e-114">**説明**</span><span class="sxs-lookup"><span data-stu-id="0361e-114">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="0361e-115">DefaultMinVersion</span><span class="sxs-lookup"><span data-stu-id="0361e-115">DefaultMinVersion</span></span>|<span data-ttu-id="0361e-116">文字列</span><span class="sxs-lookup"><span data-stu-id="0361e-116">string</span></span>|<span data-ttu-id="0361e-117">省略可能</span><span class="sxs-lookup"><span data-stu-id="0361e-117">optional</span></span>|<span data-ttu-id="0361e-p101">すべての子の **Set** 要素に対して、既定の [MinVersion](set.md) 属性値を指定します。既定値は "1.1" です。</span><span class="sxs-lookup"><span data-stu-id="0361e-p101">Specifies the default  **MinVersion** attribute value for all child [Set](set.md) elements. The default value is "1.1".</span></span>|

## <a name="remarks"></a><span data-ttu-id="0361e-120">解説</span><span class="sxs-lookup"><span data-stu-id="0361e-120">Remarks</span></span>

<span data-ttu-id="0361e-121">利用できる要件セットの詳細については、「[Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0361e-121">For more information about requirement sets, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="0361e-122">**Set** 要素の **MinVersion** 属性と **Sets** 要素の **DefaultMinVersion** 属性の詳細については、「[マニフェストで Requirements 要素を設定する](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="0361e-122">For more information about the  **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

