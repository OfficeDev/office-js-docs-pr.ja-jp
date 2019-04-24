---
title: マニフェスト ファイルの SourceLocation 要素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: b2b78065fc8bde6fc827ddcb21e2bc700ed5bf49
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450689"
---
# <a name="sourcelocation-element"></a><span data-ttu-id="971da-102">SourceLocation 要素</span><span class="sxs-lookup"><span data-stu-id="971da-102">SourceLocation element</span></span>

<span data-ttu-id="971da-103">Excel でカスタム関数によって使用される Script または Page 要素が必要とするリソースの場所を定義します。</span><span class="sxs-lookup"><span data-stu-id="971da-103">Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="971da-104">属性</span><span class="sxs-lookup"><span data-stu-id="971da-104">Attributes</span></span>

| <span data-ttu-id="971da-105">**属性**</span><span class="sxs-lookup"><span data-stu-id="971da-105">**Attribute**</span></span> | <span data-ttu-id="971da-106">**必須**</span><span class="sxs-lookup"><span data-stu-id="971da-106">**Required**</span></span> | <span data-ttu-id="971da-107">**説明**</span><span class="sxs-lookup"><span data-stu-id="971da-107">**Description**</span></span>                                                                      |
|---------------|--------------|--------------------------------------------------------------------------------------|
| <span data-ttu-id="971da-108">resid</span><span class="sxs-lookup"><span data-stu-id="971da-108">resid</span></span>         | <span data-ttu-id="971da-109">はい</span><span class="sxs-lookup"><span data-stu-id="971da-109">Yes</span></span>          | <span data-ttu-id="971da-110">マニフェストの &lt;Resources&gt; セクションで定義される URL リソースの名前。</span><span class="sxs-lookup"><span data-stu-id="971da-110">The name of a URL resource defined in the &lt;Resources&gt; section of the manifest.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="971da-111">子要素</span><span class="sxs-lookup"><span data-stu-id="971da-111">Child elements</span></span>

<span data-ttu-id="971da-112">なし</span><span class="sxs-lookup"><span data-stu-id="971da-112">None</span></span>

## <a name="example"></a><span data-ttu-id="971da-113">例</span><span class="sxs-lookup"><span data-stu-id="971da-113">Example</span></span>

```xml
<SourceLocation resid="pageURL"/>
```
