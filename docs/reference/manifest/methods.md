---
title: マニフェスト ファイルの Methods 要素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: b2ef9725b76b21af8d41b9e571d2851464aa1fcc
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596887"
---
# <a name="methods-element"></a><span data-ttu-id="f10b4-102">Methods 要素</span><span class="sxs-lookup"><span data-stu-id="f10b4-102">Methods element</span></span>

<span data-ttu-id="f10b4-103">Office アドインをアクティブにするために必要な Office JavaScript API のメソッドの一覧を指定します。</span><span class="sxs-lookup"><span data-stu-id="f10b4-103">Specifies the list of Office JavaScript API methods that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="f10b4-104">**アドインの種類:** コンテンツ、作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="f10b4-104">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="f10b4-105">構文</span><span class="sxs-lookup"><span data-stu-id="f10b4-105">Syntax</span></span>

```XML
<Methods>
   ...
</Methods>
```

## <a name="contained-in"></a><span data-ttu-id="f10b4-106">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="f10b4-106">Contained in</span></span>

[<span data-ttu-id="f10b4-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="f10b4-107">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="f10b4-108">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="f10b4-108">Can contain</span></span>

[<span data-ttu-id="f10b4-109">Method</span><span class="sxs-lookup"><span data-stu-id="f10b4-109">Method</span></span>](method.md)

## <a name="remarks"></a><span data-ttu-id="f10b4-110">注釈</span><span class="sxs-lookup"><span data-stu-id="f10b4-110">Remarks</span></span>

<span data-ttu-id="f10b4-111">**メソッド**と**メソッド**の要素は、メールアドインではサポートされていません。要件セットの詳細については、「 [Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="f10b4-111">The **Methods** and **Method** elements aren't supported in mail add-ins. For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>
