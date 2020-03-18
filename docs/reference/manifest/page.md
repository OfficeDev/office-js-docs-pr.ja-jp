---
title: マニフェスト ファイルの Page 要素
description: Page 要素は、カスタム関数が Excel で使用する HTML ページ設定を定義します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 0c56b955b79f9052ee2c89a391dd95b2975d69c2
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720485"
---
# <a name="page-element"></a><span data-ttu-id="24603-103">Page 要素</span><span class="sxs-lookup"><span data-stu-id="24603-103">Page element</span></span>

<span data-ttu-id="24603-104">Excel でカスタム関数によって使用される HTML ページの設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="24603-104">Defines HTML page settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="24603-105">属性</span><span class="sxs-lookup"><span data-stu-id="24603-105">Attributes</span></span>

<span data-ttu-id="24603-106">なし</span><span class="sxs-lookup"><span data-stu-id="24603-106">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="24603-107">子要素</span><span class="sxs-lookup"><span data-stu-id="24603-107">Child elements</span></span>

|  <span data-ttu-id="24603-108">要素</span><span class="sxs-lookup"><span data-stu-id="24603-108">Element</span></span>  |  <span data-ttu-id="24603-109">必須</span><span class="sxs-lookup"><span data-stu-id="24603-109">Required</span></span>  |  <span data-ttu-id="24603-110">説明</span><span class="sxs-lookup"><span data-stu-id="24603-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="24603-111">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="24603-111">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="24603-112">はい</span><span class="sxs-lookup"><span data-stu-id="24603-112">Yes</span></span>  | <span data-ttu-id="24603-113">カスタム関数によって使用される HTML ファイルのリソース ID を持つ文字列。</span><span class="sxs-lookup"><span data-stu-id="24603-113">String with the resource id of the HTML file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="24603-114">例</span><span class="sxs-lookup"><span data-stu-id="24603-114">Example</span></span>

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
