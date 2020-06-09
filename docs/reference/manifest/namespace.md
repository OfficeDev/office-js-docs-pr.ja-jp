---
title: マニフェスト ファイルの Namespace 要素
description: Namespace 要素は、Excel でカスタム関数によって使用される名前空間を定義します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: f4b3510c6c137bd303af8a3eaac8ebe66c5f4dc7
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612235"
---
# <a name="namespace-element"></a><span data-ttu-id="5d5c2-103">Namespace 要素</span><span class="sxs-lookup"><span data-stu-id="5d5c2-103">Namespace element</span></span>

<span data-ttu-id="5d5c2-104">Excel でカスタム関数によって使用される名前空間を定義します。</span><span class="sxs-lookup"><span data-stu-id="5d5c2-104">Defines the namespace used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="5d5c2-105">属性</span><span class="sxs-lookup"><span data-stu-id="5d5c2-105">Attributes</span></span>

|  <span data-ttu-id="5d5c2-106">属性</span><span class="sxs-lookup"><span data-stu-id="5d5c2-106">Attribute</span></span>  |  <span data-ttu-id="5d5c2-107">必須</span><span class="sxs-lookup"><span data-stu-id="5d5c2-107">Required</span></span>  |  <span data-ttu-id="5d5c2-108">説明</span><span class="sxs-lookup"><span data-stu-id="5d5c2-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="5d5c2-109">**resid="namespace"**</span><span class="sxs-lookup"><span data-stu-id="5d5c2-109">**resid="namespace"**</span></span>  |  <span data-ttu-id="5d5c2-110">いいえ</span><span class="sxs-lookup"><span data-stu-id="5d5c2-110">No</span></span>  | <span data-ttu-id="5d5c2-111">[Resources](resources.md) 要素で指定されているカスタム関数の ShortStrings のタイトルと一致する必要があります。</span><span class="sxs-lookup"><span data-stu-id="5d5c2-111">Should match the ShortStrings title for your custom function, specified within the [Resources](resources.md) element.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="5d5c2-112">子要素</span><span class="sxs-lookup"><span data-stu-id="5d5c2-112">Child elements</span></span>

<span data-ttu-id="5d5c2-113">なし</span><span class="sxs-lookup"><span data-stu-id="5d5c2-113">None</span></span>

## <a name="example"></a><span data-ttu-id="5d5c2-114">例</span><span class="sxs-lookup"><span data-stu-id="5d5c2-114">Example</span></span>

```xml
<Namespace resid="namespace" />
```
