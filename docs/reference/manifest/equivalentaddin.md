---
title: マニフェストファイルの EquivalentAddin 要素
description: 同等の COM アドインまたは XLL の下位互換性を指定します。
ms.date: 06/19/2019
localization_priority: Normal
ms.openlocfilehash: 425b926901b7325665eeede04263f74e4b854d50
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718287"
---
# <a name="equivalentaddin-element"></a><span data-ttu-id="e5073-103">EquivalentAddin 要素</span><span class="sxs-lookup"><span data-stu-id="e5073-103">EquivalentAddin element</span></span>

<span data-ttu-id="e5073-104">同等の COM アドインまたは XLL の下位互換性を指定します。</span><span class="sxs-lookup"><span data-stu-id="e5073-104">Specifies backwards compatibility for an equivalent COM add-in or XLL.</span></span>

<span data-ttu-id="e5073-105">**アドインの種類:** 作業ウィンドウ、ユーザー設定関数</span><span class="sxs-lookup"><span data-stu-id="e5073-105">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="e5073-106">構文</span><span class="sxs-lookup"><span data-stu-id="e5073-106">Syntax</span></span>

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a><span data-ttu-id="e5073-107">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="e5073-107">Contained in</span></span>

[<span data-ttu-id="e5073-108">EquivalentAdd</span><span class="sxs-lookup"><span data-stu-id="e5073-108">EquivalentAdd-ins</span></span>](equivalentaddins.md)

## <a name="must-contain"></a><span data-ttu-id="e5073-109">含める必要があるもの</span><span class="sxs-lookup"><span data-stu-id="e5073-109">Must contain</span></span>

[<span data-ttu-id="e5073-110">型</span><span class="sxs-lookup"><span data-stu-id="e5073-110">Type</span></span>](type.md)

## <a name="can-contain"></a><span data-ttu-id="e5073-111">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="e5073-111">Can contain</span></span>

<span data-ttu-id="e5073-112">[ProgId](progid.md)
[ファイル名](filename.md)</span><span class="sxs-lookup"><span data-stu-id="e5073-112">[ProgId](progid.md)
[FileName](filename.md)</span></span>

## <a name="remarks"></a><span data-ttu-id="e5073-113">注釈</span><span class="sxs-lookup"><span data-stu-id="e5073-113">Remarks</span></span>

<span data-ttu-id="e5073-114">COM アドインを同等のアドインとして指定するには、と`ProgId` `Type`の両方の要素を指定します。</span><span class="sxs-lookup"><span data-stu-id="e5073-114">To specify a COM add-in as the equivalent add-in, provide both the `ProgId` and `Type` elements.</span></span> <span data-ttu-id="e5073-115">XLL を同等のアドインとして指定するには、と`FileName` `Type`の両方の要素を指定します。</span><span class="sxs-lookup"><span data-stu-id="e5073-115">To specify an XLL as the equivalent add-in, provide both the `FileName` and `Type` elements.</span></span>

## <a name="see-also"></a><span data-ttu-id="e5073-116">関連項目</span><span class="sxs-lookup"><span data-stu-id="e5073-116">See also</span></span>

- [<span data-ttu-id="e5073-117">XLL ユーザー定義関数と互換性のある、カスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="e5073-117">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="e5073-118">既存の COM アドインと互換性のある Excel アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="e5073-118">Make your Excel add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)