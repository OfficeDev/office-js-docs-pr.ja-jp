---
title: マニフェスト ファイルの Methods 要素
description: メソッド要素は、Office アドインをアクティブにするために必要な Office JavaScript API メソッドのリストを指定します。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: d96eed07b6853cb51c24214b6017f14d6de6b83b
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718063"
---
# <a name="methods-element"></a><span data-ttu-id="47a2d-103">Methods 要素</span><span class="sxs-lookup"><span data-stu-id="47a2d-103">Methods element</span></span>

<span data-ttu-id="47a2d-104">Office アドインをアクティブにするために必要な Office JavaScript API のメソッドの一覧を指定します。</span><span class="sxs-lookup"><span data-stu-id="47a2d-104">Specifies the list of Office JavaScript API methods that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="47a2d-105">**アドインの種類:** コンテンツ、作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="47a2d-105">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="47a2d-106">構文</span><span class="sxs-lookup"><span data-stu-id="47a2d-106">Syntax</span></span>

```XML
<Methods>
   ...
</Methods>
```

## <a name="contained-in"></a><span data-ttu-id="47a2d-107">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="47a2d-107">Contained in</span></span>

[<span data-ttu-id="47a2d-108">Requirements</span><span class="sxs-lookup"><span data-stu-id="47a2d-108">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="47a2d-109">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="47a2d-109">Can contain</span></span>

[<span data-ttu-id="47a2d-110">Method</span><span class="sxs-lookup"><span data-stu-id="47a2d-110">Method</span></span>](method.md)

## <a name="remarks"></a><span data-ttu-id="47a2d-111">注釈</span><span class="sxs-lookup"><span data-stu-id="47a2d-111">Remarks</span></span>

<span data-ttu-id="47a2d-112">**メソッド**と**メソッド**の要素は、メールアドインではサポートされていません。要件セットの詳細については、「 [Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="47a2d-112">The **Methods** and **Method** elements aren't supported in mail add-ins. For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>
