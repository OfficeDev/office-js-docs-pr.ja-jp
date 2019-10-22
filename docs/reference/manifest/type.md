---
title: マニフェストファイルの Type 要素
description: ''
ms.date: 05/03/2019
localization_priority: Normal
ms.openlocfilehash: 1c053d65c5e3c6ce597c9912ec608e0b36bc623b
ms.sourcegitcommit: b3996b1444e520b44cf752e76eef50908386ca26
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/21/2019
ms.locfileid: "33628229"
---
# <a name="type-element"></a><span data-ttu-id="bf35b-102">Type 要素</span><span class="sxs-lookup"><span data-stu-id="bf35b-102">Type element</span></span>

<span data-ttu-id="bf35b-103">対応するアドインが COM addin または XLL であるかどうかを指定します。</span><span class="sxs-lookup"><span data-stu-id="bf35b-103">Specifies if the equivalent add-in is a COM addin or an XLL.</span></span>

<span data-ttu-id="bf35b-104">**アドインの種類:** 作業ウィンドウ、ユーザー設定関数</span><span class="sxs-lookup"><span data-stu-id="bf35b-104">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="bf35b-105">構文</span><span class="sxs-lookup"><span data-stu-id="bf35b-105">Syntax</span></span>

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a><span data-ttu-id="bf35b-106">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="bf35b-106">Contained in</span></span>

[<span data-ttu-id="bf35b-107">EquivalentAdd</span><span class="sxs-lookup"><span data-stu-id="bf35b-107">EquivalentAdd-in</span></span>](equivalentaddin.md)

## <a name="add-in-type-values"></a><span data-ttu-id="bf35b-108">アドインの種類の値</span><span class="sxs-lookup"><span data-stu-id="bf35b-108">Add-in type values</span></span>

<span data-ttu-id="bf35b-109">`Type`要素には、次のいずれかの値を指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="bf35b-109">You must specify one of the following values for the `Type` element.</span></span>

- <span data-ttu-id="bf35b-110">COM: 対応するアドインが COM アドインであることを指定します。</span><span class="sxs-lookup"><span data-stu-id="bf35b-110">COM: Specifies the equivalent add-in is a COM add-in.</span></span>
- <span data-ttu-id="bf35b-111">XLL: 対応するアドインが Excel XLL であることを指定します。</span><span class="sxs-lookup"><span data-stu-id="bf35b-111">XLL: Specifies the equivalent add-in is an Excel XLL.</span></span>

## <a name="see-also"></a><span data-ttu-id="bf35b-112">関連項目</span><span class="sxs-lookup"><span data-stu-id="bf35b-112">See also</span></span>

- [<span data-ttu-id="bf35b-113">XLL ユーザー定義関数と互換性のある、カスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="bf35b-113">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="bf35b-114">既存の COM アドインと互換性のある Excel アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="bf35b-114">Make your Excel add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)