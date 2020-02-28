---
title: マニフェスト ファイルの Methods 要素
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 52e12de0fde9fa1ede4687c3f27707d1dc3dce5f
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325256"
---
# <a name="methods-element"></a><span data-ttu-id="aaf55-102">Methods 要素</span><span class="sxs-lookup"><span data-stu-id="aaf55-102">Methods element</span></span>

<span data-ttu-id="aaf55-103">Office アドインをアクティブにするために必要な Office JavaScript API のメソッドの一覧を指定します。</span><span class="sxs-lookup"><span data-stu-id="aaf55-103">Specifies the list of Office JavaScript API methods that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="aaf55-104">**アドインの種類:** コンテンツ、作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="aaf55-104">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="aaf55-105">構文</span><span class="sxs-lookup"><span data-stu-id="aaf55-105">Syntax</span></span>

```XML
<Methods>
   ...
</Methods>
```

## <a name="contained-in"></a><span data-ttu-id="aaf55-106">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="aaf55-106">Contained in</span></span>

[<span data-ttu-id="aaf55-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="aaf55-107">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="aaf55-108">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="aaf55-108">Can contain</span></span>

[<span data-ttu-id="aaf55-109">Method</span><span class="sxs-lookup"><span data-stu-id="aaf55-109">Method</span></span>](method.md)

## <a name="remarks"></a><span data-ttu-id="aaf55-110">注釈</span><span class="sxs-lookup"><span data-stu-id="aaf55-110">Remarks</span></span>

<span data-ttu-id="aaf55-111">**メソッド**と**メソッド**の要素は、メールアドインではサポートされていません。要件セットの詳細については、「 [Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="aaf55-111">The **Methods** and **Method** elements aren't supported in mail add-ins. For more information about requirement sets, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

