---
title: マニフェストファイルの EquivalentAddin 要素
description: ''
ms.date: 04/22/2019
localization_priority: Normal
ms.openlocfilehash: 9cb1bb6d7a9cc3df3f4e39f8180b38d47d0a6882
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/26/2019
ms.locfileid: "33356896"
---
# <a name="equivalentaddin-element"></a><span data-ttu-id="4a774-102">EquivalentAddin 要素</span><span class="sxs-lookup"><span data-stu-id="4a774-102">EquivalentAddin element</span></span>

<span data-ttu-id="4a774-103">同等の COM アドインまたは XLL の下位互換性を指定します。</span><span class="sxs-lookup"><span data-stu-id="4a774-103">Specifies backwards compatibility for an equivalent COM add-in or XLL.</span></span>

<span data-ttu-id="4a774-104">**アドインの種類:** 作業ウィンドウ、ユーザー設定関数</span><span class="sxs-lookup"><span data-stu-id="4a774-104">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="4a774-105">構文</span><span class="sxs-lookup"><span data-stu-id="4a774-105">Syntax</span></span>

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a><span data-ttu-id="4a774-106">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="4a774-106">Contained in</span></span>

[<span data-ttu-id="4a774-107">EquivalentAdd</span><span class="sxs-lookup"><span data-stu-id="4a774-107">EquivalentAdd-ins</span></span>](equivalentaddins.md)

## <a name="must-contain"></a><span data-ttu-id="4a774-108">含める必要があるもの</span><span class="sxs-lookup"><span data-stu-id="4a774-108">Must contain</span></span>

[<span data-ttu-id="4a774-109">Type</span><span class="sxs-lookup"><span data-stu-id="4a774-109">Type</span></span>](type.md)

## <a name="can-contain"></a><span data-ttu-id="4a774-110">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="4a774-110">Can contain</span></span>

<span data-ttu-id="4a774-111">[ProgID](progid.md)
[ファイル名](filename.md)</span><span class="sxs-lookup"><span data-stu-id="4a774-111">[ProgID](progid.md)
[FileName](filename.md)</span></span>

## <a name="remarks"></a><span data-ttu-id="4a774-112">注釈</span><span class="sxs-lookup"><span data-stu-id="4a774-112">Remarks</span></span>

<span data-ttu-id="4a774-113">COM アドインを同等のアドインとして指定するには、と`ProgID` `Type`の両方の要素を指定します。</span><span class="sxs-lookup"><span data-stu-id="4a774-113">To specify a COM add-in as the equivalent add-in, provide both the `ProgID` and `Type` elements.</span></span> <span data-ttu-id="4a774-114">XLL を同等のアドインとして指定するには、と`FileName` `Type`の両方の要素を指定します。</span><span class="sxs-lookup"><span data-stu-id="4a774-114">To specify an XLL as the equivalent add-in, provide both the `FileName` and `Type` elements.</span></span>

## <a name="see-also"></a><span data-ttu-id="4a774-115">関連項目</span><span class="sxs-lookup"><span data-stu-id="4a774-115">See also</span></span>

- [<span data-ttu-id="4a774-116">カスタム関数を XLL ユーザー定義関数と互換性があるようにする</span><span class="sxs-lookup"><span data-stu-id="4a774-116">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="4a774-117">既存の COM アドインと互換性のある Office アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="4a774-117">Make your Office Add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)