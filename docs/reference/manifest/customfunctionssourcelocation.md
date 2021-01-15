---
title: マニフェスト ファイル内のカスタム関数の SourceLocation 要素
description: Excel でカスタム関数によって使用される Script または Page 要素が必要とするリソースの場所を定義します。
ms.date: 08/07/2020
localization_priority: Normal
ms.openlocfilehash: 6001673f1954a4af2de66ff7611069c3fb402a13
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771383"
---
# <a name="sourcelocation-element-custom-functions"></a><span data-ttu-id="18e94-103">SourceLocation 要素 (カスタム関数)</span><span class="sxs-lookup"><span data-stu-id="18e94-103">SourceLocation element (custom functions)</span></span>

<span data-ttu-id="18e94-104">Excel でカスタム関数によって使用される Script または Page 要素が必要とするリソースの場所を定義します。</span><span class="sxs-lookup"><span data-stu-id="18e94-104">Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="18e94-105">属性</span><span class="sxs-lookup"><span data-stu-id="18e94-105">Attributes</span></span>

| <span data-ttu-id="18e94-106">属性</span><span class="sxs-lookup"><span data-stu-id="18e94-106">Attribute</span></span> | <span data-ttu-id="18e94-107">必須</span><span class="sxs-lookup"><span data-stu-id="18e94-107">Required</span></span> | <span data-ttu-id="18e94-108">説明</span><span class="sxs-lookup"><span data-stu-id="18e94-108">Description</span></span>                                                                          |
|-----------|----------|--------------------------------------------------------------------------------------|
| <span data-ttu-id="18e94-109">resid</span><span class="sxs-lookup"><span data-stu-id="18e94-109">resid</span></span>     | <span data-ttu-id="18e94-110">はい</span><span class="sxs-lookup"><span data-stu-id="18e94-110">Yes</span></span>      | <span data-ttu-id="18e94-111">マニフェストの &lt;Resources&gt; セクションで定義される URL リソースの名前。</span><span class="sxs-lookup"><span data-stu-id="18e94-111">The name of a URL resource defined in the &lt;Resources&gt; section of the manifest.</span></span> <span data-ttu-id="18e94-112">使用できる文字数は 32 文字です。</span><span class="sxs-lookup"><span data-stu-id="18e94-112">Can be no more than 32 characters.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="18e94-113">子要素</span><span class="sxs-lookup"><span data-stu-id="18e94-113">Child elements</span></span>

<span data-ttu-id="18e94-114">なし</span><span class="sxs-lookup"><span data-stu-id="18e94-114">None</span></span>

## <a name="example"></a><span data-ttu-id="18e94-115">例</span><span class="sxs-lookup"><span data-stu-id="18e94-115">Example</span></span>

```xml
<SourceLocation resid="pageURL"/>
```
