---
title: マニフェストファイル内のカスタム関数の SourceLocation 要素
description: Excel でカスタム関数によって使用される Script または Page 要素が必要とするリソースの場所を定義します。
ms.date: 08/07/2020
localization_priority: Normal
ms.openlocfilehash: 1c509987b0ce7948a63fa8ad51f7cf9c84144c5f
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641383"
---
# <a name="sourcelocation-element-custom-functions"></a><span data-ttu-id="5c184-103">SourceLocation 要素 (カスタム関数)</span><span class="sxs-lookup"><span data-stu-id="5c184-103">SourceLocation element (custom functions)</span></span>

<span data-ttu-id="5c184-104">Excel でカスタム関数によって使用される Script または Page 要素が必要とするリソースの場所を定義します。</span><span class="sxs-lookup"><span data-stu-id="5c184-104">Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="5c184-105">属性</span><span class="sxs-lookup"><span data-stu-id="5c184-105">Attributes</span></span>

| <span data-ttu-id="5c184-106">属性</span><span class="sxs-lookup"><span data-stu-id="5c184-106">Attribute</span></span> | <span data-ttu-id="5c184-107">必須</span><span class="sxs-lookup"><span data-stu-id="5c184-107">Required</span></span> | <span data-ttu-id="5c184-108">説明</span><span class="sxs-lookup"><span data-stu-id="5c184-108">Description</span></span>                                                                          |
|-----------|----------|--------------------------------------------------------------------------------------|
| <span data-ttu-id="5c184-109">resid</span><span class="sxs-lookup"><span data-stu-id="5c184-109">resid</span></span>     | <span data-ttu-id="5c184-110">はい</span><span class="sxs-lookup"><span data-stu-id="5c184-110">Yes</span></span>      | <span data-ttu-id="5c184-111">マニフェストの &lt;Resources&gt; セクションで定義される URL リソースの名前。</span><span class="sxs-lookup"><span data-stu-id="5c184-111">The name of a URL resource defined in the &lt;Resources&gt; section of the manifest.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="5c184-112">子要素</span><span class="sxs-lookup"><span data-stu-id="5c184-112">Child elements</span></span>

<span data-ttu-id="5c184-113">なし</span><span class="sxs-lookup"><span data-stu-id="5c184-113">None</span></span>

## <a name="example"></a><span data-ttu-id="5c184-114">例</span><span class="sxs-lookup"><span data-stu-id="5c184-114">Example</span></span>

```xml
<SourceLocation resid="pageURL"/>
```
