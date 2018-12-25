---
title: マニフェスト ファイルの SourceLocation 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 878a8184984e31fdbcf46192a2f56507edaf4b37
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432408"
---
# <a name="sourcelocation-element"></a><span data-ttu-id="79870-102">SourceLocation 要素</span><span class="sxs-lookup"><span data-stu-id="79870-102">SourceLocation element</span></span>

<span data-ttu-id="79870-103">Excel でカスタム関数によって使用される Script または Page 要素が必要とするリソースの場所を定義します。</span><span class="sxs-lookup"><span data-stu-id="79870-103">Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="79870-104">属性</span><span class="sxs-lookup"><span data-stu-id="79870-104">Attributes</span></span>

| <span data-ttu-id="79870-105">**属性**</span><span class="sxs-lookup"><span data-stu-id="79870-105">**Attribute**</span></span> | <span data-ttu-id="79870-106">**必須**</span><span class="sxs-lookup"><span data-stu-id="79870-106">**Required**</span></span> | <span data-ttu-id="79870-107">**説明**</span><span class="sxs-lookup"><span data-stu-id="79870-107">**Description**</span></span>                                                                      |
|---------------|--------------|--------------------------------------------------------------------------------------|
| <span data-ttu-id="79870-108">resid</span><span class="sxs-lookup"><span data-stu-id="79870-108">resid</span></span>         | <span data-ttu-id="79870-109">はい</span><span class="sxs-lookup"><span data-stu-id="79870-109">Yes</span></span>          | <span data-ttu-id="79870-110">マニフェストの &lt;Resources&gt; セクションで定義される URL リソースの名前。</span><span class="sxs-lookup"><span data-stu-id="79870-110">The name of a URL resource defined in the &lt;Resources&gt; section of the manifest.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="79870-111">子要素</span><span class="sxs-lookup"><span data-stu-id="79870-111">Child elements</span></span>

<span data-ttu-id="79870-112">なし</span><span class="sxs-lookup"><span data-stu-id="79870-112">None</span></span>

## <a name="example"></a><span data-ttu-id="79870-113">例</span><span class="sxs-lookup"><span data-stu-id="79870-113">Example</span></span>

```xml
<SourceLocation resid="pageURL"/>
```