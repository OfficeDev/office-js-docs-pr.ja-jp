---
title: マニフェスト ファイルの Metadata 要素
description: Metadata 要素は、Excel でカスタム関数によって使用されるメタデータ設定を定義します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 8ea81818aa96b407ce386ec318495ec5ba773d05
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718070"
---
# <a name="metadata-element"></a><span data-ttu-id="73429-103">MetaData 要素</span><span class="sxs-lookup"><span data-stu-id="73429-103">Metadata element</span></span>

<span data-ttu-id="73429-104">Excel でカスタム関数によって使用されるメタデータの設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="73429-104">Defines the metadata settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="73429-105">属性</span><span class="sxs-lookup"><span data-stu-id="73429-105">Attributes</span></span>

<span data-ttu-id="73429-106">なし</span><span class="sxs-lookup"><span data-stu-id="73429-106">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="73429-107">子要素</span><span class="sxs-lookup"><span data-stu-id="73429-107">Child elements</span></span>

|  <span data-ttu-id="73429-108">要素</span><span class="sxs-lookup"><span data-stu-id="73429-108">Element</span></span>  |  <span data-ttu-id="73429-109">必須</span><span class="sxs-lookup"><span data-stu-id="73429-109">Required</span></span>  |  <span data-ttu-id="73429-110">説明</span><span class="sxs-lookup"><span data-stu-id="73429-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="73429-111">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="73429-111">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="73429-112">はい</span><span class="sxs-lookup"><span data-stu-id="73429-112">Yes</span></span>  | <span data-ttu-id="73429-113">カスタム関数によって使用される JSON ファイルのリソース ID を持つ文字列。</span><span class="sxs-lookup"><span data-stu-id="73429-113">String with the resource id of the JSON file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="73429-114">例</span><span class="sxs-lookup"><span data-stu-id="73429-114">Example</span></span>

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
