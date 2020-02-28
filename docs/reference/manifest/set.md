---
title: マニフェスト ファイルの Set 要素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: d86b3123ff856e8618f92629308787b543e8228b
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324807"
---
# <a name="set-element"></a><span data-ttu-id="b204f-102">Set 要素</span><span class="sxs-lookup"><span data-stu-id="b204f-102">Set element</span></span>

<span data-ttu-id="b204f-103">Office アドインをアクティブにするために必要な Office JavaScript API の要件セットを指定します。</span><span class="sxs-lookup"><span data-stu-id="b204f-103">Specifies a requirement set from the Office JavaScript API that your Office Add-in requires to activate.</span></span>

<span data-ttu-id="b204f-104">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="b204f-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="b204f-105">構文</span><span class="sxs-lookup"><span data-stu-id="b204f-105">Syntax</span></span>

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a><span data-ttu-id="b204f-106">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="b204f-106">Contained in</span></span>

[<span data-ttu-id="b204f-107">Sets</span><span class="sxs-lookup"><span data-stu-id="b204f-107">Sets</span></span>](sets.md)

## <a name="attributes"></a><span data-ttu-id="b204f-108">属性</span><span class="sxs-lookup"><span data-stu-id="b204f-108">Attributes</span></span>

|<span data-ttu-id="b204f-109">**属性**</span><span class="sxs-lookup"><span data-stu-id="b204f-109">**Attribute**</span></span>|<span data-ttu-id="b204f-110">**型**</span><span class="sxs-lookup"><span data-stu-id="b204f-110">**Type**</span></span>|<span data-ttu-id="b204f-111">**必須**</span><span class="sxs-lookup"><span data-stu-id="b204f-111">**Required**</span></span>|<span data-ttu-id="b204f-112">**説明**</span><span class="sxs-lookup"><span data-stu-id="b204f-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="b204f-113">名前</span><span class="sxs-lookup"><span data-stu-id="b204f-113">Name</span></span>|<span data-ttu-id="b204f-114">string</span><span class="sxs-lookup"><span data-stu-id="b204f-114">string</span></span>|<span data-ttu-id="b204f-115">必須</span><span class="sxs-lookup"><span data-stu-id="b204f-115">required</span></span>|<span data-ttu-id="b204f-116">[要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)の名前。</span><span class="sxs-lookup"><span data-stu-id="b204f-116">The name of a [requirement set](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>|
|<span data-ttu-id="b204f-117">MinVersion</span><span class="sxs-lookup"><span data-stu-id="b204f-117">MinVersion</span></span>|<span data-ttu-id="b204f-118">文字列</span><span class="sxs-lookup"><span data-stu-id="b204f-118">string</span></span>|<span data-ttu-id="b204f-119">省略可能</span><span class="sxs-lookup"><span data-stu-id="b204f-119">optional</span></span>|<span data-ttu-id="b204f-120">アドインで必要な API セットの最小バージョンを指定します。</span><span class="sxs-lookup"><span data-stu-id="b204f-120">Specifies the minimum version of the API set required by your add-in.</span></span> <span data-ttu-id="b204f-121">親[Sets](sets.md)要素で指定されている場合、 **defaultminversion**の値を上書きします。</span><span class="sxs-lookup"><span data-stu-id="b204f-121">Overrides the value of **DefaultMinVersion**, if it is specified in the parent [Sets](sets.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="b204f-122">解説</span><span class="sxs-lookup"><span data-stu-id="b204f-122">Remarks</span></span>

<span data-ttu-id="b204f-123">利用できる要件セットの詳細については、「[Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b204f-123">For more information about requirement sets, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="b204f-124">**Set**要素の**MinVersion**属性と**Sets**要素の**defaultminversion**属性の詳細については、「マニフェストの[要件要素を設定する](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="b204f-124">For more information about the **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="b204f-125">メール アドインの場合、使用可能なのは `"Mailbox"` 要件セットのみです。</span><span class="sxs-lookup"><span data-stu-id="b204f-125">For mail add-ins, there is only one  `"Mailbox"` requirement set available.</span></span> <span data-ttu-id="b204f-126">この要件セットには、Outlook のメール アドインでサポートされている API のサブセット全体が含まれ、メール アドインのマニフェストで `"Mailbox"` 要件セットを指定する必要があります (コンテンツ アドインと作業ウィンドウ アドインの場合とは異なり、オプションではありません)。</span><span class="sxs-lookup"><span data-stu-id="b204f-126">This requirement set contains the entire subset of API supported in mail add-ins for Outlook, and you must specify the `"Mailbox"` requirement set in your mail add-in's manifest (it's not optional as is the case for content and task pane add-ins).</span></span> <span data-ttu-id="b204f-127">Also, you can't declare support for specific methods in mail add-ins.</span><span class="sxs-lookup"><span data-stu-id="b204f-127">Also, you can't declare support for specific methods in mail add-ins.</span></span>
