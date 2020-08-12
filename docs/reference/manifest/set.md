---
title: マニフェスト ファイルの Set 要素
description: Set 要素は、Office アドインをアクティブにするために必要な、office JavaScript API の要件セットを指定します。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 608830e1ebc0d2e2d4c170b48bba00b3a19e87af
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641418"
---
# <a name="set-element"></a><span data-ttu-id="44cca-103">Set 要素</span><span class="sxs-lookup"><span data-stu-id="44cca-103">Set element</span></span>

<span data-ttu-id="44cca-104">Office アドインをアクティブにするために必要な Office JavaScript API の要件セットを指定します。</span><span class="sxs-lookup"><span data-stu-id="44cca-104">Specifies a requirement set from the Office JavaScript API that your Office Add-in requires to activate.</span></span>

<span data-ttu-id="44cca-105">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="44cca-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="44cca-106">構文</span><span class="sxs-lookup"><span data-stu-id="44cca-106">Syntax</span></span>

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a><span data-ttu-id="44cca-107">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="44cca-107">Contained in</span></span>

[<span data-ttu-id="44cca-108">Sets</span><span class="sxs-lookup"><span data-stu-id="44cca-108">Sets</span></span>](sets.md)

## <a name="attributes"></a><span data-ttu-id="44cca-109">属性</span><span class="sxs-lookup"><span data-stu-id="44cca-109">Attributes</span></span>

|<span data-ttu-id="44cca-110">属性</span><span class="sxs-lookup"><span data-stu-id="44cca-110">Attribute</span></span>|<span data-ttu-id="44cca-111">型</span><span class="sxs-lookup"><span data-stu-id="44cca-111">Type</span></span>|<span data-ttu-id="44cca-112">必須</span><span class="sxs-lookup"><span data-stu-id="44cca-112">Required</span></span>|<span data-ttu-id="44cca-113">説明</span><span class="sxs-lookup"><span data-stu-id="44cca-113">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="44cca-114">名前</span><span class="sxs-lookup"><span data-stu-id="44cca-114">Name</span></span>|<span data-ttu-id="44cca-115">string</span><span class="sxs-lookup"><span data-stu-id="44cca-115">string</span></span>|<span data-ttu-id="44cca-116">必須</span><span class="sxs-lookup"><span data-stu-id="44cca-116">required</span></span>|<span data-ttu-id="44cca-117">[要件セット](../../develop/office-versions-and-requirement-sets.md)の名前。</span><span class="sxs-lookup"><span data-stu-id="44cca-117">The name of a [requirement set](../../develop/office-versions-and-requirement-sets.md).</span></span>|
|<span data-ttu-id="44cca-118">MinVersion</span><span class="sxs-lookup"><span data-stu-id="44cca-118">MinVersion</span></span>|<span data-ttu-id="44cca-119">文字列</span><span class="sxs-lookup"><span data-stu-id="44cca-119">string</span></span>|<span data-ttu-id="44cca-120">省略可能</span><span class="sxs-lookup"><span data-stu-id="44cca-120">optional</span></span>|<span data-ttu-id="44cca-121">アドインで必要な API セットの最小バージョンを指定します。</span><span class="sxs-lookup"><span data-stu-id="44cca-121">Specifies the minimum version of the API set required by your add-in.</span></span> <span data-ttu-id="44cca-122">親[Sets](sets.md)要素で指定されている場合、 **defaultminversion**の値を上書きします。</span><span class="sxs-lookup"><span data-stu-id="44cca-122">Overrides the value of **DefaultMinVersion**, if it is specified in the parent [Sets](sets.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="44cca-123">解説</span><span class="sxs-lookup"><span data-stu-id="44cca-123">Remarks</span></span>

<span data-ttu-id="44cca-124">利用できる要件セットの詳細については、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="44cca-124">For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="44cca-125">**Set**要素の**MinVersion**属性と**Sets**要素の**defaultminversion**属性の詳細については、「マニフェストの[要件要素を設定する](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="44cca-125">For more information about the **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](../../develop/specify-office-hosts-and-api-requirements.md#set-the-requirements-element-in-the-manifest).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="44cca-126">メール アドインの場合、使用可能なのは `"Mailbox"` 要件セットのみです。</span><span class="sxs-lookup"><span data-stu-id="44cca-126">For mail add-ins, there is only one  `"Mailbox"` requirement set available.</span></span> <span data-ttu-id="44cca-127">この要件セットには、Outlook のメール アドインでサポートされている API のサブセット全体が含まれ、メール アドインのマニフェストで `"Mailbox"` 要件セットを指定する必要があります (コンテンツ アドインと作業ウィンドウ アドインの場合とは異なり、オプションではありません)。</span><span class="sxs-lookup"><span data-stu-id="44cca-127">This requirement set contains the entire subset of API supported in mail add-ins for Outlook, and you must specify the `"Mailbox"` requirement set in your mail add-in's manifest (it's not optional as is the case for content and task pane add-ins).</span></span> <span data-ttu-id="44cca-128">Also, you can't declare support for specific methods in mail add-ins.</span><span class="sxs-lookup"><span data-stu-id="44cca-128">Also, you can't declare support for specific methods in mail add-ins.</span></span>
