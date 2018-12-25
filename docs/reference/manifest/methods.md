---
title: マニフェスト ファイルの Methods 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 6e280cb49eadef587cd3a91e0664ece3c3d59f50
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432754"
---
# <a name="methods-element"></a><span data-ttu-id="54c10-102">Methods 要素</span><span class="sxs-lookup"><span data-stu-id="54c10-102">Methods element</span></span>

<span data-ttu-id="54c10-103">Office アドインをアクティブにするために必要な JavaScript API for Office のメソッドの一覧を指定します。</span><span class="sxs-lookup"><span data-stu-id="54c10-103">Specifies the list of JavaScript API for Office methods that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="54c10-104">**アドインの種類:** コンテンツ、作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="54c10-104">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="54c10-105">構文</span><span class="sxs-lookup"><span data-stu-id="54c10-105">Syntax</span></span>

```XML
<Methods>
   ...
</Methods>
```

## <a name="contained-in"></a><span data-ttu-id="54c10-106">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="54c10-106">Contained in</span></span>

[<span data-ttu-id="54c10-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="54c10-107">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="54c10-108">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="54c10-108">Can contain</span></span>

[<span data-ttu-id="54c10-109">Method</span><span class="sxs-lookup"><span data-stu-id="54c10-109">Method</span></span>](method.md)

## <a name="remarks"></a><span data-ttu-id="54c10-110">解説</span><span class="sxs-lookup"><span data-stu-id="54c10-110">Remarks</span></span>

<span data-ttu-id="54c10-111">**Methods** と **Method** 要素はメール アドインではサポートされていません。要件セットの詳細については、「[Office のバージョンと要件セット](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)」をご覧ください。</span><span class="sxs-lookup"><span data-stu-id="54c10-111">The  Methods and Method elements aren't supported in mail add-ins. For more information about requirement sets, see Specify Office hosts and API requirements.</span></span>

