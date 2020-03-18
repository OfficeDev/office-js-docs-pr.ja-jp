---
title: マニフェスト ファイルの Namespace 要素
description: Namespace 要素は、Excel でカスタム関数によって使用される名前空間を定義します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 45fd0caa039fdeb885cba4b739750fbd8b642252
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718057"
---
# <a name="namespace-element"></a><span data-ttu-id="2b9ef-103">Namespace 要素</span><span class="sxs-lookup"><span data-stu-id="2b9ef-103">Namespace element</span></span>

<span data-ttu-id="2b9ef-104">Excel でカスタム関数によって使用される名前空間を定義します。</span><span class="sxs-lookup"><span data-stu-id="2b9ef-104">Defines the namespace used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="2b9ef-105">属性</span><span class="sxs-lookup"><span data-stu-id="2b9ef-105">Attributes</span></span>

|  <span data-ttu-id="2b9ef-106">属性</span><span class="sxs-lookup"><span data-stu-id="2b9ef-106">Attribute</span></span>  |  <span data-ttu-id="2b9ef-107">必須</span><span class="sxs-lookup"><span data-stu-id="2b9ef-107">Required</span></span>  |  <span data-ttu-id="2b9ef-108">説明</span><span class="sxs-lookup"><span data-stu-id="2b9ef-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="2b9ef-109">**resid="namespace"**</span><span class="sxs-lookup"><span data-stu-id="2b9ef-109">**resid="namespace"**</span></span>  |  <span data-ttu-id="2b9ef-110">はい</span><span class="sxs-lookup"><span data-stu-id="2b9ef-110">Yes</span></span>  | <span data-ttu-id="2b9ef-111">[Resources](resources.md) 要素で指定されているカスタム関数の ShortStrings のタイトルと一致する必要があります。</span><span class="sxs-lookup"><span data-stu-id="2b9ef-111">Should match the ShortStrings title for your custom function, specified within the [Resources](resources.md) element.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="2b9ef-112">子要素</span><span class="sxs-lookup"><span data-stu-id="2b9ef-112">Child elements</span></span>

<span data-ttu-id="2b9ef-113">なし</span><span class="sxs-lookup"><span data-stu-id="2b9ef-113">None</span></span>

## <a name="example"></a><span data-ttu-id="2b9ef-114">例</span><span class="sxs-lookup"><span data-stu-id="2b9ef-114">Example</span></span>

```xml
<Namespace resid="namespace" />
```
