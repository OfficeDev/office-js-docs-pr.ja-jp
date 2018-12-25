---
title: マニフェスト ファイルの Metadata 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 79038fc13eba76176be19e484ffa57e64727bf94
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432663"
---
# <a name="metadata-element"></a><span data-ttu-id="7a23a-102">MetaData 要素</span><span class="sxs-lookup"><span data-stu-id="7a23a-102">MetaData element</span></span>

<span data-ttu-id="7a23a-103">Excel でカスタム関数によって使用されるメタデータの設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="7a23a-103">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="7a23a-104">属性</span><span class="sxs-lookup"><span data-stu-id="7a23a-104">Attributes</span></span>

<span data-ttu-id="7a23a-105">なし</span><span class="sxs-lookup"><span data-stu-id="7a23a-105">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="7a23a-106">子要素</span><span class="sxs-lookup"><span data-stu-id="7a23a-106">Child elements</span></span>

|  <span data-ttu-id="7a23a-107">要素</span><span class="sxs-lookup"><span data-stu-id="7a23a-107">Element</span></span>  |  <span data-ttu-id="7a23a-108">必須</span><span class="sxs-lookup"><span data-stu-id="7a23a-108">Required</span></span>  |  <span data-ttu-id="7a23a-109">説明</span><span class="sxs-lookup"><span data-stu-id="7a23a-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="7a23a-110">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="7a23a-110">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="7a23a-111">はい</span><span class="sxs-lookup"><span data-stu-id="7a23a-111">Yes</span></span>  | <span data-ttu-id="7a23a-112">カスタム関数によって使用される JSON ファイルのリソース ID を持つ文字列。</span><span class="sxs-lookup"><span data-stu-id="7a23a-112">String with the resource id of the HTML file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="7a23a-113">例</span><span class="sxs-lookup"><span data-stu-id="7a23a-113">Example</span></span>

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
