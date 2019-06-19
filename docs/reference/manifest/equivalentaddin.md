---
title: マニフェストファイルの EquivalentAddin 要素
description: ''
ms.date: 06/19/2019
localization_priority: Normal
ms.openlocfilehash: 33cfb8b73e050fad7e392e0234962d346e903713
ms.sourcegitcommit: 4bf5159a3821f4277c07d89e88808c4c3a25ff81
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/18/2019
ms.locfileid: "35059924"
---
# <a name="equivalentaddin-element"></a><span data-ttu-id="891dd-102">EquivalentAddin 要素</span><span class="sxs-lookup"><span data-stu-id="891dd-102">EquivalentAddin element</span></span>

<span data-ttu-id="891dd-103">同等の COM アドインまたは XLL の下位互換性を指定します。</span><span class="sxs-lookup"><span data-stu-id="891dd-103">Specifies backwards compatibility for an equivalent COM add-in or XLL.</span></span>

<span data-ttu-id="891dd-104">**アドインの種類:** 作業ウィンドウ、ユーザー設定関数</span><span class="sxs-lookup"><span data-stu-id="891dd-104">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="891dd-105">構文</span><span class="sxs-lookup"><span data-stu-id="891dd-105">Syntax</span></span>

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a><span data-ttu-id="891dd-106">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="891dd-106">Contained in</span></span>

[<span data-ttu-id="891dd-107">EquivalentAdd</span><span class="sxs-lookup"><span data-stu-id="891dd-107">EquivalentAdd-ins</span></span>](equivalentaddins.md)

## <a name="must-contain"></a><span data-ttu-id="891dd-108">含める必要があるもの</span><span class="sxs-lookup"><span data-stu-id="891dd-108">Must contain</span></span>

[<span data-ttu-id="891dd-109">Type</span><span class="sxs-lookup"><span data-stu-id="891dd-109">Type</span></span>](type.md)

## <a name="can-contain"></a><span data-ttu-id="891dd-110">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="891dd-110">Can contain</span></span>

<span data-ttu-id="891dd-111">[ProgId](progid.md)
[ファイル名](filename.md)</span><span class="sxs-lookup"><span data-stu-id="891dd-111">[ProgId](progid.md)
[FileName](filename.md)</span></span>

## <a name="remarks"></a><span data-ttu-id="891dd-112">解説</span><span class="sxs-lookup"><span data-stu-id="891dd-112">Remarks</span></span>

<span data-ttu-id="891dd-113">COM アドインを同等のアドインとして指定するには、と`ProgId` `Type`の両方の要素を指定します。</span><span class="sxs-lookup"><span data-stu-id="891dd-113">To specify a COM add-in as the equivalent add-in, provide both the `ProgId` and `Type` elements.</span></span> <span data-ttu-id="891dd-114">XLL を同等のアドインとして指定するには、と`FileName` `Type`の両方の要素を指定します。</span><span class="sxs-lookup"><span data-stu-id="891dd-114">To specify an XLL as the equivalent add-in, provide both the `FileName` and `Type` elements.</span></span>

## <a name="see-also"></a><span data-ttu-id="891dd-115">関連項目</span><span class="sxs-lookup"><span data-stu-id="891dd-115">See also</span></span>

- [<span data-ttu-id="891dd-116">XLL ユーザー定義関数と互換性のある、カスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="891dd-116">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="891dd-117">既存の COM アドインと互換性のある Excel アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="891dd-117">Make your Excel add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)