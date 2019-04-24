---
title: マニフェスト ファイルの Page 要素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: f85cc3a834f628a7390f3b96faa596145c7d331a
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452075"
---
# <a name="page-element"></a><span data-ttu-id="f45ca-102">Page 要素</span><span class="sxs-lookup"><span data-stu-id="f45ca-102">Page element</span></span>

<span data-ttu-id="f45ca-103">Excel でカスタム関数によって使用される HTML ページの設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="f45ca-103">Defines HTML page settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="f45ca-104">属性</span><span class="sxs-lookup"><span data-stu-id="f45ca-104">Attributes</span></span>

<span data-ttu-id="f45ca-105">なし</span><span class="sxs-lookup"><span data-stu-id="f45ca-105">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="f45ca-106">子要素</span><span class="sxs-lookup"><span data-stu-id="f45ca-106">Child elements</span></span>

|  <span data-ttu-id="f45ca-107">要素</span><span class="sxs-lookup"><span data-stu-id="f45ca-107">Element</span></span>  |  <span data-ttu-id="f45ca-108">必須</span><span class="sxs-lookup"><span data-stu-id="f45ca-108">Required</span></span>  |  <span data-ttu-id="f45ca-109">説明</span><span class="sxs-lookup"><span data-stu-id="f45ca-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="f45ca-110">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="f45ca-110">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="f45ca-111">はい</span><span class="sxs-lookup"><span data-stu-id="f45ca-111">Yes</span></span>  | <span data-ttu-id="f45ca-112">カスタム関数によって使用される HTML ファイルのリソース ID を持つ文字列。</span><span class="sxs-lookup"><span data-stu-id="f45ca-112">String with the resource id of the HTML file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="f45ca-113">例</span><span class="sxs-lookup"><span data-stu-id="f45ca-113">Example</span></span>

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
