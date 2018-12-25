---
title: マニフェスト ファイルの Namespace 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 8000ea5774b38dd038888c686a33127a2d5bc482
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432327"
---
# <a name="namespace-element"></a><span data-ttu-id="a1843-102">Namespace 要素</span><span class="sxs-lookup"><span data-stu-id="a1843-102">Namespace element</span></span>

<span data-ttu-id="a1843-103">Excel でカスタム関数によって使用される名前空間を定義します。</span><span class="sxs-lookup"><span data-stu-id="a1843-103">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="a1843-104">属性</span><span class="sxs-lookup"><span data-stu-id="a1843-104">Attributes</span></span>

|  <span data-ttu-id="a1843-105">属性</span><span class="sxs-lookup"><span data-stu-id="a1843-105">Attribute</span></span>  |  <span data-ttu-id="a1843-106">必須</span><span class="sxs-lookup"><span data-stu-id="a1843-106">Required</span></span>  |  <span data-ttu-id="a1843-107">説明</span><span class="sxs-lookup"><span data-stu-id="a1843-107">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="a1843-108">**resid="namespace"**</span><span class="sxs-lookup"><span data-stu-id="a1843-108">**resid="namespace"**</span></span>  |  <span data-ttu-id="a1843-109">はい</span><span class="sxs-lookup"><span data-stu-id="a1843-109">Yes</span></span>  | <span data-ttu-id="a1843-110">[Resources](resources.md) 要素で指定されているカスタム関数の ShortStrings のタイトルと一致する必要があります。</span><span class="sxs-lookup"><span data-stu-id="a1843-110">Should match the ShortStrings title for your custom function, specified within the [Resources](resources.md) element.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="a1843-111">子要素</span><span class="sxs-lookup"><span data-stu-id="a1843-111">Child elements</span></span>

<span data-ttu-id="a1843-112">なし</span><span class="sxs-lookup"><span data-stu-id="a1843-112">None</span></span>

## <a name="example"></a><span data-ttu-id="a1843-113">例</span><span class="sxs-lookup"><span data-stu-id="a1843-113">Example</span></span>

```xml
<Namespace resid="namespace" />
```
