---
title: マニフェスト ファイルの EquivalentAddin 要素
description: 同等の COM アドインまたは XLL の下位互換性を指定します。
ms.date: 03/09/2021
localization_priority: Normal
ms.openlocfilehash: 412a3ce7bd12d886b7b88b5b84938e28295aba5d
ms.sourcegitcommit: c0c61fe84f3c5de88bd7eac29120056bb1224fc8
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2021
ms.locfileid: "50836838"
---
# <a name="equivalentaddin-element"></a><span data-ttu-id="de72f-103">EquivalentAddin 要素</span><span class="sxs-lookup"><span data-stu-id="de72f-103">EquivalentAddin element</span></span>

<span data-ttu-id="de72f-104">同等の COM アドインまたは XLL の下位互換性を指定します。</span><span class="sxs-lookup"><span data-stu-id="de72f-104">Specifies backwards compatibility for an equivalent COM add-in or XLL.</span></span>

<span data-ttu-id="de72f-105">**アドインの種類:** 作業ウィンドウ、カスタム関数</span><span class="sxs-lookup"><span data-stu-id="de72f-105">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="de72f-106">構文</span><span class="sxs-lookup"><span data-stu-id="de72f-106">Syntax</span></span>

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a><span data-ttu-id="de72f-107">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="de72f-107">Contained in</span></span>

[<span data-ttu-id="de72f-108">EquivalentAddins</span><span class="sxs-lookup"><span data-stu-id="de72f-108">EquivalentAddins</span></span>](equivalentaddins.md)

## <a name="must-contain"></a><span data-ttu-id="de72f-109">含める必要があるもの</span><span class="sxs-lookup"><span data-stu-id="de72f-109">Must contain</span></span>

[<span data-ttu-id="de72f-110">型</span><span class="sxs-lookup"><span data-stu-id="de72f-110">Type</span></span>](type.md)

## <a name="can-contain"></a><span data-ttu-id="de72f-111">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="de72f-111">Can contain</span></span>

<span data-ttu-id="de72f-112">[ProgId](progid.md) 
[FileName](filename.md)</span><span class="sxs-lookup"><span data-stu-id="de72f-112">[ProgId](progid.md)
[FileName](filename.md)</span></span>

## <a name="remarks"></a><span data-ttu-id="de72f-113">備考</span><span class="sxs-lookup"><span data-stu-id="de72f-113">Remarks</span></span>

<span data-ttu-id="de72f-114">COM アドインを同等のアドインとして指定するには、要素と要素の両方を `ProgId` 指定 `Type` します。</span><span class="sxs-lookup"><span data-stu-id="de72f-114">To specify a COM add-in as the equivalent add-in, provide both the `ProgId` and `Type` elements.</span></span> <span data-ttu-id="de72f-115">XLL を同等のアドインとして指定するには、要素と要素の両方を `FileName` 指定 `Type` します。</span><span class="sxs-lookup"><span data-stu-id="de72f-115">To specify an XLL as the equivalent add-in, provide both the `FileName` and `Type` elements.</span></span>

## <a name="see-also"></a><span data-ttu-id="de72f-116">関連項目</span><span class="sxs-lookup"><span data-stu-id="de72f-116">See also</span></span>

- [<span data-ttu-id="de72f-117">XLL ユーザー定義関数と互換性のある、カスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="de72f-117">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="de72f-118">Office アドインを既存の COM アドインと互換できるようにする</span><span class="sxs-lookup"><span data-stu-id="de72f-118">Make your Office Add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)