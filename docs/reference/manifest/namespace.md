---
title: マニフェスト ファイルの Namespace 要素
description: Namespace 要素は、カスタム関数が Excel で使用する名前空間を定義します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 342f5ebcafa861838956f1033f8597cf05e60215
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771261"
---
# <a name="namespace-element"></a><span data-ttu-id="e0e39-103">Namespace 要素</span><span class="sxs-lookup"><span data-stu-id="e0e39-103">Namespace element</span></span>

<span data-ttu-id="e0e39-104">Excel でカスタム関数によって使用される名前空間を定義します。</span><span class="sxs-lookup"><span data-stu-id="e0e39-104">Defines the namespace used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="e0e39-105">属性</span><span class="sxs-lookup"><span data-stu-id="e0e39-105">Attributes</span></span>

|  <span data-ttu-id="e0e39-106">属性</span><span class="sxs-lookup"><span data-stu-id="e0e39-106">Attribute</span></span>  |  <span data-ttu-id="e0e39-107">必須</span><span class="sxs-lookup"><span data-stu-id="e0e39-107">Required</span></span>  |  <span data-ttu-id="e0e39-108">説明</span><span class="sxs-lookup"><span data-stu-id="e0e39-108">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="e0e39-109">**resid="namespace"**</span><span class="sxs-lookup"><span data-stu-id="e0e39-109">**resid="namespace"**</span></span>  |  <span data-ttu-id="e0e39-110">いいえ</span><span class="sxs-lookup"><span data-stu-id="e0e39-110">No</span></span>  | <span data-ttu-id="e0e39-111">[Resources](resources.md) 要素で指定されているカスタム関数の ShortStrings のタイトルと一致する必要があります。</span><span class="sxs-lookup"><span data-stu-id="e0e39-111">Should match the ShortStrings title for your custom function, specified within the [Resources](resources.md) element.</span></span> <span data-ttu-id="e0e39-112">使用できる文字数は 32 文字です。</span><span class="sxs-lookup"><span data-stu-id="e0e39-112">Can be no more than 32 characters.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="e0e39-113">子要素</span><span class="sxs-lookup"><span data-stu-id="e0e39-113">Child elements</span></span>

<span data-ttu-id="e0e39-114">なし</span><span class="sxs-lookup"><span data-stu-id="e0e39-114">None</span></span>

## <a name="example"></a><span data-ttu-id="e0e39-115">例</span><span class="sxs-lookup"><span data-stu-id="e0e39-115">Example</span></span>

```xml
<Namespace resid="namespace" />
```
