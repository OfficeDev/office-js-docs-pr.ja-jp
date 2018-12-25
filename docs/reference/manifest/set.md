---
title: マニフェスト ファイルの Set 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 0f137f7b08d6f1d0b0d972173c8085713b0f979d
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432768"
---
# <a name="set-element"></a><span data-ttu-id="22b26-102">Set 要素</span><span class="sxs-lookup"><span data-stu-id="22b26-102">Set element</span></span>

<span data-ttu-id="22b26-103">Office アドインをアクティブにするために必要な JavaScript API for Office の要件セットを指定します。</span><span class="sxs-lookup"><span data-stu-id="22b26-103">Specifies a requirement set from the JavaScript API for Office that your Office Add-in requires to activate.</span></span>

<span data-ttu-id="22b26-104">**アドインの種類:** コンテンツ、作業ウィンドウ、メール</span><span class="sxs-lookup"><span data-stu-id="22b26-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="22b26-105">構文</span><span class="sxs-lookup"><span data-stu-id="22b26-105">Syntax</span></span>

```XML
<Set Name="string" MinVersion="n .n">
```

## <a name="contained-in"></a><span data-ttu-id="22b26-106">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="22b26-106">Contained in</span></span>

[<span data-ttu-id="22b26-107">Sets</span><span class="sxs-lookup"><span data-stu-id="22b26-107">Sets</span></span>](sets.md)

## <a name="attributes"></a><span data-ttu-id="22b26-108">属性</span><span class="sxs-lookup"><span data-stu-id="22b26-108">Attributes</span></span>

|<span data-ttu-id="22b26-109">**属性**</span><span class="sxs-lookup"><span data-stu-id="22b26-109">**Attribute**</span></span>|<span data-ttu-id="22b26-110">**型**</span><span class="sxs-lookup"><span data-stu-id="22b26-110">**Type**</span></span>|<span data-ttu-id="22b26-111">**必須**</span><span class="sxs-lookup"><span data-stu-id="22b26-111">**Required**</span></span>|<span data-ttu-id="22b26-112">**説明**</span><span class="sxs-lookup"><span data-stu-id="22b26-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="22b26-113">名前</span><span class="sxs-lookup"><span data-stu-id="22b26-113">Name</span></span>|<span data-ttu-id="22b26-114">string</span><span class="sxs-lookup"><span data-stu-id="22b26-114">string</span></span>|<span data-ttu-id="22b26-115">必須</span><span class="sxs-lookup"><span data-stu-id="22b26-115">required</span></span>|<span data-ttu-id="22b26-116">[要件セット](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)の名前。</span><span class="sxs-lookup"><span data-stu-id="22b26-116">The name of a [requirement set](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>|
|<span data-ttu-id="22b26-117">MinVersion</span><span class="sxs-lookup"><span data-stu-id="22b26-117">MinVersion</span></span>|<span data-ttu-id="22b26-118">文字列</span><span class="sxs-lookup"><span data-stu-id="22b26-118">string</span></span>|<span data-ttu-id="22b26-119">省略可能</span><span class="sxs-lookup"><span data-stu-id="22b26-119">optional</span></span>|<span data-ttu-id="22b26-p101">アドインに必要な API セットの最小バージョンを指定します。**DefaultMinVersion** の値が親の [Sets](sets.md) 要素に指定されている場合は、その値を上書きします。</span><span class="sxs-lookup"><span data-stu-id="22b26-p101">Specifies the minimum version of the API set required by your add-in. Overrides the value of  **DefaultMinVersion**, if it is specified in the parent [Sets](sets.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="22b26-122">解説</span><span class="sxs-lookup"><span data-stu-id="22b26-122">Remarks</span></span>

<span data-ttu-id="22b26-123">利用できる要件セットの詳細については、「[Office のバージョンと要件セット](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="22b26-123">For more information about available requirement sets, see [Office add-in requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="22b26-124">**Set** 要素の **MinVersion** 属性と **Sets** 要素の **DefaultMinVersion** 属性の詳細については、「[マニフェストで Requirements 要素を設定する](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="22b26-124">For more information about the  **MinVersion** attribute of the **Set** element and the **DefaultMinVersion** attribute of the **Sets** element, see [Set the Requirements element in the manifest](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements#set-the-requirements-element-in-the-manifest).</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="22b26-125">メール アドインの場合、使用可能なのは `"Mailbox"` 要件セットのみです。</span><span class="sxs-lookup"><span data-stu-id="22b26-125">For mail add-ins, there is only one  `"Mailbox"` requirement set available.</span></span> <span data-ttu-id="22b26-126">この要件セットには、Outlook のメール アドインでサポートされている API のサブセット全体が含まれ、メール アドインのマニフェストで `"Mailbox"` 要件セットを指定する必要があります (コンテンツ アドインと作業ウィンドウ アドインの場合とは異なり、オプションではありません)。</span><span class="sxs-lookup"><span data-stu-id="22b26-126">Important  For mail add-ins, there is only one   requirement set available. This requirement set contains the entire subset of API supported in mail add-ins for Outlook, and you must specify the `"Mailbox"` requirement set in your mail add-in's manifest (it's not optional as is the case for content and task pane add-ins). Also, you can't declare support for specific methods in mail add-ins.</span></span> <span data-ttu-id="22b26-127">また、メール アドインで特定のメソッドのサポートを宣言することもできません。</span><span class="sxs-lookup"><span data-stu-id="22b26-127">Also, you can't declare support for specific methods in mail add-ins.</span></span>
