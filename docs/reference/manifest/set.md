---
title: マニフェスト ファイルの Set 要素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 0f408d698d297eaa6287ff268bdb7fc737a5a24d
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870535"
---
# <a name="set-element"></a><span data-ttu-id="0a0b9-102">Set 要素</span><span class="sxs-lookup"><span data-stu-id="0a0b9-102">Set element</span></span>

<span data-ttu-id="0a0b9-103">Office アドインをアクティブにするのに必要な JavaScript API for Office の要件セットを指定します。</span><span class="sxs-lookup"><span data-stu-id="0a0b9-103">Specifies a requirement set from the JavaScript API for Office that your Office Add-in requires to activate.</span></span>

<span data-ttu-id="0a0b9-104">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="0a0b9-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="0a0b9-105">構文</span><span class="sxs-lookup"><span data-stu-id="0a0b9-105">Syntax</span></span>

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a><span data-ttu-id="0a0b9-106">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="0a0b9-106">Contained in</span></span>

[<span data-ttu-id="0a0b9-107">Sets</span><span class="sxs-lookup"><span data-stu-id="0a0b9-107">Sets</span></span>](sets.md)

## <a name="attributes"></a><span data-ttu-id="0a0b9-108">属性</span><span class="sxs-lookup"><span data-stu-id="0a0b9-108">Attributes</span></span>

|<span data-ttu-id="0a0b9-109">**属性**</span><span class="sxs-lookup"><span data-stu-id="0a0b9-109">**Attribute**</span></span>|<span data-ttu-id="0a0b9-110">**型**</span><span class="sxs-lookup"><span data-stu-id="0a0b9-110">**Type**</span></span>|<span data-ttu-id="0a0b9-111">**必須**</span><span class="sxs-lookup"><span data-stu-id="0a0b9-111">**Required**</span></span>|<span data-ttu-id="0a0b9-112">**説明**</span><span class="sxs-lookup"><span data-stu-id="0a0b9-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="0a0b9-113">名前</span><span class="sxs-lookup"><span data-stu-id="0a0b9-113">Name</span></span>|<span data-ttu-id="0a0b9-114">string</span><span class="sxs-lookup"><span data-stu-id="0a0b9-114">string</span></span>|<span data-ttu-id="0a0b9-115">必須</span><span class="sxs-lookup"><span data-stu-id="0a0b9-115">required</span></span>|<span data-ttu-id="0a0b9-116">[要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)の名前。</span><span class="sxs-lookup"><span data-stu-id="0a0b9-116">The name of a [requirement set](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>|
|<span data-ttu-id="0a0b9-117">MinVersion</span><span class="sxs-lookup"><span data-stu-id="0a0b9-117">MinVersion</span></span>|<span data-ttu-id="0a0b9-118">文字列</span><span class="sxs-lookup"><span data-stu-id="0a0b9-118">string</span></span>|<span data-ttu-id="0a0b9-119">省略可能</span><span class="sxs-lookup"><span data-stu-id="0a0b9-119">optional</span></span>|<span data-ttu-id="0a0b9-p101">アドインに必要な API セットの最小バージョンを指定します。**DefaultMinVersion** の値が親の [Sets](sets.md) 要素に指定されている場合は、その値を上書きします。</span><span class="sxs-lookup"><span data-stu-id="0a0b9-p101">Specifies the minimum version of the API set required by your add-in. Overrides the value of  **DefaultMinVersion**, if it is specified in the parent [Sets](sets.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="0a0b9-122">解説</span><span class="sxs-lookup"><span data-stu-id="0a0b9-122">Remarks</span></span>

<span data-ttu-id="0a0b9-123">利用できる要件セットの詳細については、「[Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0a0b9-123">For more information about requirement sets, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="0a0b9-124">**Set** 要素の **MinVersion** 属性と **Sets** 要素の **DefaultMinVersion** 属性の詳細については、「[マニフェストで Requirements 要素を設定する](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="0a0b9-124">For more information about the  **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="0a0b9-125">メール アドインの場合、使用可能なのは `"Mailbox"` 要件セットのみです。</span><span class="sxs-lookup"><span data-stu-id="0a0b9-125">For mail add-ins, there is only one  `"Mailbox"` requirement set available.</span></span> <span data-ttu-id="0a0b9-126">この要件セットには、Outlook のメール アドインでサポートされている API のサブセット全体が含まれ、メール アドインのマニフェストで `"Mailbox"` 要件セットを指定する必要があります (コンテンツ アドインと作業ウィンドウ アドインの場合とは異なり、オプションではありません)。</span><span class="sxs-lookup"><span data-stu-id="0a0b9-126">This requirement set contains the entire subset of API supported in mail add-ins for Outlook, and you must specify the `"Mailbox"` requirement set in your mail add-in's manifest (it's not optional as is the case for content and task pane add-ins).</span></span> <span data-ttu-id="0a0b9-127">Also, you can't declare support for specific methods in mail add-ins.</span><span class="sxs-lookup"><span data-stu-id="0a0b9-127">Also, you can't declare support for specific methods in mail add-ins.</span></span>
