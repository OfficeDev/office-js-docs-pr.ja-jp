---
title: マニフェストファイルの Type 要素
description: ''
ms.date: 04/22/2019
localization_priority: Normal
ms.openlocfilehash: 28514e25d7877c0452fbf006a31f078cd980d819
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/26/2019
ms.locfileid: "33356905"
---
# <a name="type-element"></a><span data-ttu-id="e52ff-102">Type 要素</span><span class="sxs-lookup"><span data-stu-id="e52ff-102">Type element</span></span>

<span data-ttu-id="e52ff-103">対応するアドインが COM addin または XLL であるかどうかを指定します。</span><span class="sxs-lookup"><span data-stu-id="e52ff-103">Specifies if the equivalent add-in is a COM addin or an XLL.</span></span>

<span data-ttu-id="e52ff-104">**アドインの種類:** 作業ウィンドウ、ユーザー設定関数</span><span class="sxs-lookup"><span data-stu-id="e52ff-104">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="e52ff-105">構文</span><span class="sxs-lookup"><span data-stu-id="e52ff-105">Syntax</span></span>

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a><span data-ttu-id="e52ff-106">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="e52ff-106">Contained in</span></span>

[<span data-ttu-id="e52ff-107">EquivalentAdd</span><span class="sxs-lookup"><span data-stu-id="e52ff-107">EquivalentAdd-in</span></span>](equivalentaddin.md)

## <a name="add-in-type-values"></a><span data-ttu-id="e52ff-108">アドインの種類の値</span><span class="sxs-lookup"><span data-stu-id="e52ff-108">Add-in type values</span></span>

<span data-ttu-id="e52ff-109">`Type`要素には、次のいずれかの値を指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e52ff-109">You must specify one of the following values for the `Type` element.</span></span>

- <span data-ttu-id="e52ff-110">com: 対応するアドインが COM アドインであることを指定します。</span><span class="sxs-lookup"><span data-stu-id="e52ff-110">COM: Specifies the equivalent add-in is a COM add-in.</span></span>
- <span data-ttu-id="e52ff-111">xll: 対応するアドインが Excel XLL であることを指定します。</span><span class="sxs-lookup"><span data-stu-id="e52ff-111">XLL: Specifies the equivalent add-in is an Excel XLL.</span></span>

## <a name="see-also"></a><span data-ttu-id="e52ff-112">関連項目</span><span class="sxs-lookup"><span data-stu-id="e52ff-112">See also</span></span>

- [<span data-ttu-id="e52ff-113">カスタム関数を XLL ユーザー定義関数と互換性があるようにする</span><span class="sxs-lookup"><span data-stu-id="e52ff-113">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="e52ff-114">既存の COM アドインと互換性のある Office アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="e52ff-114">Make your Office Add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)