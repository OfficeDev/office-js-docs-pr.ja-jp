---
title: マニフェスト ファイルの SourceLocation 要素
description: Excel でカスタム関数によって使用される Script または Page 要素が必要とするリソースの場所を定義します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 88ae0558577167074a870170833617c4f60730f1
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612313"
---
# <a name="sourcelocation-element"></a><span data-ttu-id="a8c32-103">SourceLocation 要素</span><span class="sxs-lookup"><span data-stu-id="a8c32-103">SourceLocation element</span></span>

<span data-ttu-id="a8c32-104">Excel でカスタム関数によって使用される Script または Page 要素が必要とするリソースの場所を定義します。</span><span class="sxs-lookup"><span data-stu-id="a8c32-104">Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="a8c32-105">属性</span><span class="sxs-lookup"><span data-stu-id="a8c32-105">Attributes</span></span>

| <span data-ttu-id="a8c32-106">**属性**</span><span class="sxs-lookup"><span data-stu-id="a8c32-106">**Attribute**</span></span> | <span data-ttu-id="a8c32-107">**必須**</span><span class="sxs-lookup"><span data-stu-id="a8c32-107">**Required**</span></span> | <span data-ttu-id="a8c32-108">**説明**</span><span class="sxs-lookup"><span data-stu-id="a8c32-108">**Description**</span></span>                                                                      |
|---------------|--------------|--------------------------------------------------------------------------------------|
| <span data-ttu-id="a8c32-109">resid</span><span class="sxs-lookup"><span data-stu-id="a8c32-109">resid</span></span>         | <span data-ttu-id="a8c32-110">はい</span><span class="sxs-lookup"><span data-stu-id="a8c32-110">Yes</span></span>          | <span data-ttu-id="a8c32-111">マニフェストの &lt;Resources&gt; セクションで定義される URL リソースの名前。</span><span class="sxs-lookup"><span data-stu-id="a8c32-111">The name of a URL resource defined in the &lt;Resources&gt; section of the manifest.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="a8c32-112">子要素</span><span class="sxs-lookup"><span data-stu-id="a8c32-112">Child elements</span></span>

<span data-ttu-id="a8c32-113">なし</span><span class="sxs-lookup"><span data-stu-id="a8c32-113">None</span></span>

## <a name="example"></a><span data-ttu-id="a8c32-114">例</span><span class="sxs-lookup"><span data-stu-id="a8c32-114">Example</span></span>

```xml
<SourceLocation resid="pageURL"/>
```
