---
title: マニフェスト ファイルの Methods 要素
description: メソッド要素は、Office アドインをアクティブにするために必要な Office JavaScript API メソッドのリストを指定します。
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: b270122240314b792ee492336417a4d133bdcc84
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609021"
---
# <a name="methods-element"></a><span data-ttu-id="7e716-103">Methods 要素</span><span class="sxs-lookup"><span data-stu-id="7e716-103">Methods element</span></span>

<span data-ttu-id="7e716-104">Office アドインをアクティブにするために必要な Office JavaScript API のメソッドの一覧を指定します。</span><span class="sxs-lookup"><span data-stu-id="7e716-104">Specifies the list of Office JavaScript API methods that your Office Add-in requires in order to activate.</span></span>

<span data-ttu-id="7e716-105">**アドインの種類:** コンテンツ、作業ウィンドウ</span><span class="sxs-lookup"><span data-stu-id="7e716-105">**Add-in type:** Content, Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="7e716-106">構文</span><span class="sxs-lookup"><span data-stu-id="7e716-106">Syntax</span></span>

```XML
<Methods>
   ...
</Methods>
```

## <a name="contained-in"></a><span data-ttu-id="7e716-107">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="7e716-107">Contained in</span></span>

[<span data-ttu-id="7e716-108">Requirements</span><span class="sxs-lookup"><span data-stu-id="7e716-108">Requirements</span></span>](requirements.md)

## <a name="can-contain"></a><span data-ttu-id="7e716-109">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="7e716-109">Can contain</span></span>

[<span data-ttu-id="7e716-110">Method</span><span class="sxs-lookup"><span data-stu-id="7e716-110">Method</span></span>](method.md)

## <a name="remarks"></a><span data-ttu-id="7e716-111">注釈</span><span class="sxs-lookup"><span data-stu-id="7e716-111">Remarks</span></span>

<span data-ttu-id="7e716-112">**メソッド**と**メソッド**の要素は、メールアドインではサポートされていません。要件セットの詳細については、「 [Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="7e716-112">The **Methods** and **Method** elements aren't supported in mail add-ins. For more information about requirement sets, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>
