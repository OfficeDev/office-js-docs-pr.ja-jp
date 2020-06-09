---
title: マニフェストファイルの EquivalentAddin 要素
description: 同等の COM アドインまたは XLL の下位互換性を指定します。
ms.date: 06/19/2019
localization_priority: Normal
ms.openlocfilehash: e14fe91bf7a5fe321019acf205ddb1753fedd569
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611562"
---
# <a name="equivalentaddin-element"></a><span data-ttu-id="629fe-103">EquivalentAddin 要素</span><span class="sxs-lookup"><span data-stu-id="629fe-103">EquivalentAddin element</span></span>

<span data-ttu-id="629fe-104">同等の COM アドインまたは XLL の下位互換性を指定します。</span><span class="sxs-lookup"><span data-stu-id="629fe-104">Specifies backwards compatibility for an equivalent COM add-in or XLL.</span></span>

<span data-ttu-id="629fe-105">**アドインの種類:** 作業ウィンドウ、ユーザー設定関数</span><span class="sxs-lookup"><span data-stu-id="629fe-105">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="629fe-106">構文</span><span class="sxs-lookup"><span data-stu-id="629fe-106">Syntax</span></span>

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a><span data-ttu-id="629fe-107">含まれる場所</span><span class="sxs-lookup"><span data-stu-id="629fe-107">Contained in</span></span>

[<span data-ttu-id="629fe-108">EquivalentAdd</span><span class="sxs-lookup"><span data-stu-id="629fe-108">EquivalentAdd-ins</span></span>](equivalentaddins.md)

## <a name="must-contain"></a><span data-ttu-id="629fe-109">含める必要があるもの</span><span class="sxs-lookup"><span data-stu-id="629fe-109">Must contain</span></span>

[<span data-ttu-id="629fe-110">種類</span><span class="sxs-lookup"><span data-stu-id="629fe-110">Type</span></span>](type.md)

## <a name="can-contain"></a><span data-ttu-id="629fe-111">含めることができるもの</span><span class="sxs-lookup"><span data-stu-id="629fe-111">Can contain</span></span>

<span data-ttu-id="629fe-112">[ProgId](progid.md) 
[ファイル名](filename.md)</span><span class="sxs-lookup"><span data-stu-id="629fe-112">[ProgId](progid.md)
[FileName](filename.md)</span></span>

## <a name="remarks"></a><span data-ttu-id="629fe-113">注釈</span><span class="sxs-lookup"><span data-stu-id="629fe-113">Remarks</span></span>

<span data-ttu-id="629fe-114">COM アドインを同等のアドインとして指定するには、との両方の要素を指定し `ProgId` `Type` ます。</span><span class="sxs-lookup"><span data-stu-id="629fe-114">To specify a COM add-in as the equivalent add-in, provide both the `ProgId` and `Type` elements.</span></span> <span data-ttu-id="629fe-115">XLL を同等のアドインとして指定するには、との両方の要素を指定し `FileName` `Type` ます。</span><span class="sxs-lookup"><span data-stu-id="629fe-115">To specify an XLL as the equivalent add-in, provide both the `FileName` and `Type` elements.</span></span>

## <a name="see-also"></a><span data-ttu-id="629fe-116">関連項目</span><span class="sxs-lookup"><span data-stu-id="629fe-116">See also</span></span>

- [<span data-ttu-id="629fe-117">XLL ユーザー定義関数と互換性のある、カスタム関数を作成します。</span><span class="sxs-lookup"><span data-stu-id="629fe-117">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="629fe-118">既存の COM アドインと互換性のある Excel アドインを作成する</span><span class="sxs-lookup"><span data-stu-id="629fe-118">Make your Excel add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)