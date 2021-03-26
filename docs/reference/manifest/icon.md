---
title: マニフェスト ファイルの Icon 要素
description: ボタン または メニュー コントロールの Image 要素を定義します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: d2986d6d0325f425d5a66d010fff003c67356862
ms.sourcegitcommit: 5ad32261f80e7ab371aba032d9024ad1275c23f9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2021
ms.locfileid: "51221361"
---
# <a name="icon-element"></a><span data-ttu-id="35df6-103">Icon 要素</span><span class="sxs-lookup"><span data-stu-id="35df6-103">Icon element</span></span>

<span data-ttu-id="35df6-104">[ボタン](control.md#button-control) または [メニュー](control.md#menu-dropdown-button-controls) コントロールの **Image** 要素を定義します。</span><span class="sxs-lookup"><span data-stu-id="35df6-104">Defines **Image** elements for [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) controls.</span></span>

## <a name="attributes"></a><span data-ttu-id="35df6-105">属性</span><span class="sxs-lookup"><span data-stu-id="35df6-105">Attributes</span></span>

|  <span data-ttu-id="35df6-106">属性</span><span class="sxs-lookup"><span data-stu-id="35df6-106">Attribute</span></span>  |  <span data-ttu-id="35df6-107">必須</span><span class="sxs-lookup"><span data-stu-id="35df6-107">Required</span></span>  |  <span data-ttu-id="35df6-108">説明</span><span class="sxs-lookup"><span data-stu-id="35df6-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="35df6-109">**xsi:type**</span><span class="sxs-lookup"><span data-stu-id="35df6-109">**xsi:type**</span></span>  |  <span data-ttu-id="35df6-110">いいえ</span><span class="sxs-lookup"><span data-stu-id="35df6-110">No</span></span>  | <span data-ttu-id="35df6-p101">定義されているアイコンの型。これは、モバイル フォーム ファクターのアイコンにのみ適用されます [MobileFormFactor](mobileformfactor.md) 要素に含まれる **Icon** 要素は、この属性を `bt:MobileIconList` に設定する必要があります。</span><span class="sxs-lookup"><span data-stu-id="35df6-p101">The type of icon being defined. This is only applicable to icons in mobile form factors. **Icon** elements contained within a [MobileFormFactor](mobileformfactor.md) element must have this attribute set to `bt:MobileIconList`.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="35df6-114">子要素</span><span class="sxs-lookup"><span data-stu-id="35df6-114">Child elements</span></span>

|  <span data-ttu-id="35df6-115">要素</span><span class="sxs-lookup"><span data-stu-id="35df6-115">Element</span></span> |  <span data-ttu-id="35df6-116">必須</span><span class="sxs-lookup"><span data-stu-id="35df6-116">Required</span></span>  |  <span data-ttu-id="35df6-117">説明</span><span class="sxs-lookup"><span data-stu-id="35df6-117">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="35df6-118">Image</span><span class="sxs-lookup"><span data-stu-id="35df6-118">Image</span></span>](#image)        | <span data-ttu-id="35df6-119">はい</span><span class="sxs-lookup"><span data-stu-id="35df6-119">Yes</span></span> |   <span data-ttu-id="35df6-120">使用するイメージの resid</span><span class="sxs-lookup"><span data-stu-id="35df6-120">resid of an image to use</span></span>         |

### <a name="image"></a><span data-ttu-id="35df6-121">Image</span><span class="sxs-lookup"><span data-stu-id="35df6-121">Image</span></span>

<span data-ttu-id="35df6-122">ボタンの画像です。</span><span class="sxs-lookup"><span data-stu-id="35df6-122">An image for the button.</span></span> <span data-ttu-id="35df6-123">**resid 属性** は 32 文字以内で **、Resources** 要素の **Images** 要素の Image 要素の **id** 属性の値に [設定する必要](resources.md)があります。</span><span class="sxs-lookup"><span data-stu-id="35df6-123">The **resid** attribute can be no more than 32 characters and must be set to the value of the **id** attribute of an **Image** element in the **Images** element in the [Resources](resources.md) element.</span></span> <span data-ttu-id="35df6-124">**size** 属性は、画像のサイズをピクセル単位で示します。</span><span class="sxs-lookup"><span data-stu-id="35df6-124">The **size** attribute indicates the size in pixels of the image.</span></span> <span data-ttu-id="35df6-125">他に 5 つのサイズ (20、24、40、48、64 ピクセル) がサポートされていますが、3 つの画像のサイズ (16、32、80 ピクセル) を必ず指定します。</span><span class="sxs-lookup"><span data-stu-id="35df6-125">Three image sizes are required (16, 32, and 80 pixels) while five other sizes are supported (20, 24, 40, 48, and 64 pixels).</span></span>

```xml
<Icon>
  <bt:Image size="16" resid="blue-icon-16" />
  <bt:Image size="32" resid="blue-icon-32" />
  <bt:Image size="80" resid="blue-icon-80" />
</Icon>
```

## <a name="additional-requirements-for-mobile-form-factors"></a><span data-ttu-id="35df6-126">モバイル フォーム ファクターの追加要件</span><span class="sxs-lookup"><span data-stu-id="35df6-126">Additional requirements for mobile form factors</span></span>

<span data-ttu-id="35df6-p103">親 **Icon** 要素が、[MobileFormFactor](mobileformfactor.md) 要素の子孫である場合は、必要な最小サイズが若干異なります。マニフェストで、最小サイズを 25、32、および 48 ピクセルに指定する必要があります。指定するサイズは、`1`、`2` または `3` に設定された `scale` 属性で必ずそれぞれ 3 回ずつ表示されます。</span><span class="sxs-lookup"><span data-stu-id="35df6-p103">When the parent **Icon** element is a descendant of a [MobileFormFactor](mobileformfactor.md) element, the minimum required sizes are slightly different. The manifest must minimally provide 25, 32, and 48 pixel sizes. Each size provided must appear three times, with a `scale` attribute set to `1`, `2`, or `3`.</span></span>

```xml
<Icon xsi:type="bt:MobileIconList">
  <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
  <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
  <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
  <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
  <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
  <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
  <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
  <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
  <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
</Icon>
```
