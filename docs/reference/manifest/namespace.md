---
title: マニフェスト ファイルの Namespace 要素
description: Namespace 要素は、Excel でカスタム関数によって使用される名前空間を定義します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: eabd73d3be98271c81723787dd3d1bdb6ee2ebcd
ms.sourcegitcommit: 315a648cce38609c3e1c92bd4a339e268f8a2e1d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2020
ms.locfileid: "42978670"
---
# <a name="namespace-element"></a><span data-ttu-id="20623-103">Namespace 要素</span><span class="sxs-lookup"><span data-stu-id="20623-103">Namespace element</span></span>

<span data-ttu-id="20623-104">Excel でカスタム関数によって使用される名前空間を定義します。</span><span class="sxs-lookup"><span data-stu-id="20623-104">Defines the namespace used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="20623-105">属性</span><span class="sxs-lookup"><span data-stu-id="20623-105">Attributes</span></span>

|  <span data-ttu-id="20623-106">属性</span><span class="sxs-lookup"><span data-stu-id="20623-106">Attribute</span></span>  |  <span data-ttu-id="20623-107">必須</span><span class="sxs-lookup"><span data-stu-id="20623-107">Required</span></span>  |  <span data-ttu-id="20623-108">説明</span><span class="sxs-lookup"><span data-stu-id="20623-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="20623-109">**resid="namespace"**</span><span class="sxs-lookup"><span data-stu-id="20623-109">**resid="namespace"**</span></span>  |  <span data-ttu-id="20623-110">いいえ</span><span class="sxs-lookup"><span data-stu-id="20623-110">No</span></span>  | <span data-ttu-id="20623-111">[Resources](resources.md) 要素で指定されているカスタム関数の ShortStrings のタイトルと一致する必要があります。</span><span class="sxs-lookup"><span data-stu-id="20623-111">Should match the ShortStrings title for your custom function, specified within the [Resources](resources.md) element.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="20623-112">子要素</span><span class="sxs-lookup"><span data-stu-id="20623-112">Child elements</span></span>

<span data-ttu-id="20623-113">なし</span><span class="sxs-lookup"><span data-stu-id="20623-113">None</span></span>

## <a name="example"></a><span data-ttu-id="20623-114">例</span><span class="sxs-lookup"><span data-stu-id="20623-114">Example</span></span>

```xml
<Namespace resid="namespace" />
```
