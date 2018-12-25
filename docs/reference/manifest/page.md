---
title: マニフェスト ファイルの Page 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 83bafd24d0b56322ea5f7d51025f2416be019168
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433734"
---
# <a name="page-element"></a><span data-ttu-id="c6e68-102">Page 要素</span><span class="sxs-lookup"><span data-stu-id="c6e68-102">Page element</span></span>

<span data-ttu-id="c6e68-103">Excel でカスタム関数によって使用される HTML ページの設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="c6e68-103">Defines HTML page settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="c6e68-104">属性</span><span class="sxs-lookup"><span data-stu-id="c6e68-104">Attributes</span></span>

<span data-ttu-id="c6e68-105">なし</span><span class="sxs-lookup"><span data-stu-id="c6e68-105">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="c6e68-106">子要素</span><span class="sxs-lookup"><span data-stu-id="c6e68-106">Child elements</span></span>

|  <span data-ttu-id="c6e68-107">要素</span><span class="sxs-lookup"><span data-stu-id="c6e68-107">Element</span></span>  |  <span data-ttu-id="c6e68-108">必須</span><span class="sxs-lookup"><span data-stu-id="c6e68-108">Required</span></span>  |  <span data-ttu-id="c6e68-109">説明</span><span class="sxs-lookup"><span data-stu-id="c6e68-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="c6e68-110">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="c6e68-110">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="c6e68-111">はい</span><span class="sxs-lookup"><span data-stu-id="c6e68-111">Yes</span></span>  | <span data-ttu-id="c6e68-112">カスタム関数によって使用される HTML ファイルのリソース ID を持つ文字列。</span><span class="sxs-lookup"><span data-stu-id="c6e68-112">String with the resource id of the HTML file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="c6e68-113">例</span><span class="sxs-lookup"><span data-stu-id="c6e68-113">Example</span></span>

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
