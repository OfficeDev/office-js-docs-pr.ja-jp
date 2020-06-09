---
title: マニフェスト ファイルの Page 要素
description: Page 要素は、カスタム関数が Excel で使用する HTML ページ設定を定義します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: aa8a2807cbf2549ded680a22b17f24513ea76b9a
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611499"
---
# <a name="page-element"></a><span data-ttu-id="e43d6-103">Page 要素</span><span class="sxs-lookup"><span data-stu-id="e43d6-103">Page element</span></span>

<span data-ttu-id="e43d6-104">Excel でカスタム関数によって使用される HTML ページの設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="e43d6-104">Defines HTML page settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="e43d6-105">属性</span><span class="sxs-lookup"><span data-stu-id="e43d6-105">Attributes</span></span>

<span data-ttu-id="e43d6-106">なし</span><span class="sxs-lookup"><span data-stu-id="e43d6-106">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="e43d6-107">子要素</span><span class="sxs-lookup"><span data-stu-id="e43d6-107">Child elements</span></span>

|  <span data-ttu-id="e43d6-108">要素</span><span class="sxs-lookup"><span data-stu-id="e43d6-108">Element</span></span>  |  <span data-ttu-id="e43d6-109">必須</span><span class="sxs-lookup"><span data-stu-id="e43d6-109">Required</span></span>  |  <span data-ttu-id="e43d6-110">説明</span><span class="sxs-lookup"><span data-stu-id="e43d6-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="e43d6-111">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="e43d6-111">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="e43d6-112">はい</span><span class="sxs-lookup"><span data-stu-id="e43d6-112">Yes</span></span>  | <span data-ttu-id="e43d6-113">カスタム関数によって使用される HTML ファイルのリソース ID を持つ文字列。</span><span class="sxs-lookup"><span data-stu-id="e43d6-113">String with the resource id of the HTML file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="e43d6-114">例</span><span class="sxs-lookup"><span data-stu-id="e43d6-114">Example</span></span>

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
