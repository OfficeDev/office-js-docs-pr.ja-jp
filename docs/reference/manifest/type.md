---
title: マニフェストファイルの Type 要素
description: Type 要素は、対応するアドインが COM アドインまたは XLL であるかどうかを指定します。
ms.date: 03/16/2020
localization_priority: Normal
ms.openlocfilehash: 9eeab172ed4ebf06fc93e42f56f8d33f5e7a92db
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720317"
---
# <a name="type-element"></a><span data-ttu-id="ac5cc-103">Type 要素</span><span class="sxs-lookup"><span data-stu-id="ac5cc-103">Type element</span></span>

<span data-ttu-id="ac5cc-104">対応するアドインが COM アドインまたは XLL であるかどうかを指定します。</span><span class="sxs-lookup"><span data-stu-id="ac5cc-104">Specifies if the equivalent add-in is a COM add-in or an XLL.</span></span>

<span data-ttu-id="ac5cc-105">**アドインの種類:** 作業ウィンドウ、ユーザー設定関数</span><span class="sxs-lookup"><span data-stu-id="ac5cc-105">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="ac5cc-106">構文</span><span class="sxs-lookup"><span data-stu-id="ac5cc-106">Syntax</span></span>

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a><span data-ttu-id="ac5cc-107">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="ac5cc-107">Contained in</span></span>

[<span data-ttu-id="ac5cc-108">EquivalentAdd</span><span class="sxs-lookup"><span data-stu-id="ac5cc-108">EquivalentAdd-in</span></span>](equivalentaddin.md)

## <a name="add-in-type-values"></a><span data-ttu-id="ac5cc-109">アドインの種類の値</span><span class="sxs-lookup"><span data-stu-id="ac5cc-109">Add-in type values</span></span>

<span data-ttu-id="ac5cc-110">`Type`要素には、次のいずれかの値を指定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="ac5cc-110">You must specify one of the following values for the `Type` element.</span></span>

- <span data-ttu-id="ac5cc-111">COM: 対応するアドインが COM アドインであることを指定します。</span><span class="sxs-lookup"><span data-stu-id="ac5cc-111">COM: Specifies the equivalent add-in is a COM add-in.</span></span>
- <span data-ttu-id="ac5cc-112">XLL: 対応するアドインが Excel XLL であることを指定します。</span><span class="sxs-lookup"><span data-stu-id="ac5cc-112">XLL: Specifies the equivalent add-in is an Excel XLL.</span></span>

## <a name="see-also"></a><span data-ttu-id="ac5cc-113">関連項目</span><span class="sxs-lookup"><span data-stu-id="ac5cc-113">See also</span></span>

- [<span data-ttu-id="ac5cc-114">XLL ユーザー定義関数と互換性のある、カスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="ac5cc-114">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="ac5cc-115">既存の COM アドインと互換性のある Excel アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="ac5cc-115">Make your Excel add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)