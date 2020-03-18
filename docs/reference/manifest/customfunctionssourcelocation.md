---
title: マニフェスト ファイルの SourceLocation 要素
description: Excel でカスタム関数によって使用される Script または Page 要素が必要とするリソースの場所を定義します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 56ebe122853c98a14c52d450bea31fecaefb15d3
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720688"
---
# <a name="sourcelocation-element"></a><span data-ttu-id="f50c3-103">SourceLocation 要素</span><span class="sxs-lookup"><span data-stu-id="f50c3-103">SourceLocation element</span></span>

<span data-ttu-id="f50c3-104">Excel でカスタム関数によって使用される Script または Page 要素が必要とするリソースの場所を定義します。</span><span class="sxs-lookup"><span data-stu-id="f50c3-104">Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="f50c3-105">属性</span><span class="sxs-lookup"><span data-stu-id="f50c3-105">Attributes</span></span>

| <span data-ttu-id="f50c3-106">**属性**</span><span class="sxs-lookup"><span data-stu-id="f50c3-106">**Attribute**</span></span> | <span data-ttu-id="f50c3-107">**必須**</span><span class="sxs-lookup"><span data-stu-id="f50c3-107">**Required**</span></span> | <span data-ttu-id="f50c3-108">**説明**</span><span class="sxs-lookup"><span data-stu-id="f50c3-108">**Description**</span></span>                                                                      |
|---------------|--------------|--------------------------------------------------------------------------------------|
| <span data-ttu-id="f50c3-109">resid</span><span class="sxs-lookup"><span data-stu-id="f50c3-109">resid</span></span>         | <span data-ttu-id="f50c3-110">はい</span><span class="sxs-lookup"><span data-stu-id="f50c3-110">Yes</span></span>          | <span data-ttu-id="f50c3-111">マニフェストの &lt;Resources&gt; セクションで定義される URL リソースの名前。</span><span class="sxs-lookup"><span data-stu-id="f50c3-111">The name of a URL resource defined in the &lt;Resources&gt; section of the manifest.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="f50c3-112">子要素</span><span class="sxs-lookup"><span data-stu-id="f50c3-112">Child elements</span></span>

<span data-ttu-id="f50c3-113">なし</span><span class="sxs-lookup"><span data-stu-id="f50c3-113">None</span></span>

## <a name="example"></a><span data-ttu-id="f50c3-114">例</span><span class="sxs-lookup"><span data-stu-id="f50c3-114">Example</span></span>

```xml
<SourceLocation resid="pageURL"/>
```
