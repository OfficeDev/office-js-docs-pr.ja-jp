---
title: マニフェスト ファイルの Namespace 要素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: faf77fe8b6bddc734f1b47eb544ffe7e1e7c4aaa
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452103"
---
# <a name="namespace-element"></a><span data-ttu-id="982e0-102">Namespace 要素</span><span class="sxs-lookup"><span data-stu-id="982e0-102">Namespace element</span></span>

<span data-ttu-id="982e0-103">Excel でカスタム関数によって使用される名前空間を定義します。</span><span class="sxs-lookup"><span data-stu-id="982e0-103">Defines the namespace used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="982e0-104">属性</span><span class="sxs-lookup"><span data-stu-id="982e0-104">Attributes</span></span>

|  <span data-ttu-id="982e0-105">属性</span><span class="sxs-lookup"><span data-stu-id="982e0-105">Attribute</span></span>  |  <span data-ttu-id="982e0-106">必須</span><span class="sxs-lookup"><span data-stu-id="982e0-106">Required</span></span>  |  <span data-ttu-id="982e0-107">説明</span><span class="sxs-lookup"><span data-stu-id="982e0-107">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="982e0-108">**resid="namespace"**</span><span class="sxs-lookup"><span data-stu-id="982e0-108">**resid="namespace"**</span></span>  |  <span data-ttu-id="982e0-109">はい</span><span class="sxs-lookup"><span data-stu-id="982e0-109">Yes</span></span>  | <span data-ttu-id="982e0-110">[Resources](resources.md) 要素で指定されているカスタム関数の ShortStrings のタイトルと一致する必要があります。</span><span class="sxs-lookup"><span data-stu-id="982e0-110">Should match the ShortStrings title for your custom function, specified within the [Resources](resources.md) element.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="982e0-111">子要素</span><span class="sxs-lookup"><span data-stu-id="982e0-111">Child elements</span></span>

<span data-ttu-id="982e0-112">なし</span><span class="sxs-lookup"><span data-stu-id="982e0-112">None</span></span>

## <a name="example"></a><span data-ttu-id="982e0-113">例</span><span class="sxs-lookup"><span data-stu-id="982e0-113">Example</span></span>

```xml
<Namespace resid="namespace" />
```
