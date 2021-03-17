---
title: マニフェスト ファイルの Type 要素
description: Type 要素は、同等のアドインが COM アドインか XLL かを指定します。
ms.date: 03/09/2021
localization_priority: Normal
ms.openlocfilehash: 5af3359c232e91b097311bfc06fc9b1c932b0703
ms.sourcegitcommit: c0c61fe84f3c5de88bd7eac29120056bb1224fc8
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2021
ms.locfileid: "50836810"
---
# <a name="type-element"></a><span data-ttu-id="234d4-103">Type 要素</span><span class="sxs-lookup"><span data-stu-id="234d4-103">Type element</span></span>

<span data-ttu-id="234d4-104">同等のアドインが COM アドインか XLL かを指定します。</span><span class="sxs-lookup"><span data-stu-id="234d4-104">Specifies if the equivalent add-in is a COM add-in or an XLL.</span></span>

<span data-ttu-id="234d4-105">**アドインの種類:** 作業ウィンドウ、カスタム関数</span><span class="sxs-lookup"><span data-stu-id="234d4-105">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="234d4-106">構文</span><span class="sxs-lookup"><span data-stu-id="234d4-106">Syntax</span></span>

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a><span data-ttu-id="234d4-107">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="234d4-107">Contained in</span></span>

[<span data-ttu-id="234d4-108">EquivalentAddin</span><span class="sxs-lookup"><span data-stu-id="234d4-108">EquivalentAddin</span></span>](equivalentaddin.md)

## <a name="add-in-type-values"></a><span data-ttu-id="234d4-109">アドインの型の値</span><span class="sxs-lookup"><span data-stu-id="234d4-109">Add-in type values</span></span>

<span data-ttu-id="234d4-110">要素には、次のいずれかの値を指定する必要 `Type` があります。</span><span class="sxs-lookup"><span data-stu-id="234d4-110">You must specify one of the following values for the `Type` element.</span></span>

- <span data-ttu-id="234d4-111">COM: COM アドインと同等のアドインを指定します。</span><span class="sxs-lookup"><span data-stu-id="234d4-111">COM: Specifies the equivalent add-in is a COM add-in.</span></span>
- <span data-ttu-id="234d4-112">XLL: Excel XLL と同等のアドインを指定します。</span><span class="sxs-lookup"><span data-stu-id="234d4-112">XLL: Specifies the equivalent add-in is an Excel XLL.</span></span>

## <a name="see-also"></a><span data-ttu-id="234d4-113">関連項目</span><span class="sxs-lookup"><span data-stu-id="234d4-113">See also</span></span>

- [<span data-ttu-id="234d4-114">XLL ユーザー定義関数と互換性のある、カスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="234d4-114">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="234d4-115">Office アドインを既存の COM アドインと互換できるようにする</span><span class="sxs-lookup"><span data-stu-id="234d4-115">Make your Office Add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)