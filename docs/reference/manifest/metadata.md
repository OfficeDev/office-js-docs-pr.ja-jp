---
title: マニフェスト ファイルの Metadata 要素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: a3aecb1983905658f3a55fdb8bf0629a8d5ef474
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452047"
---
# <a name="metadata-element"></a><span data-ttu-id="0c32b-102">MetaData 要素</span><span class="sxs-lookup"><span data-stu-id="0c32b-102">Metadata element</span></span>

<span data-ttu-id="0c32b-103">Excel でカスタム関数によって使用されるメタデータの設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="0c32b-103">Defines the metadata settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="0c32b-104">属性</span><span class="sxs-lookup"><span data-stu-id="0c32b-104">Attributes</span></span>

<span data-ttu-id="0c32b-105">なし</span><span class="sxs-lookup"><span data-stu-id="0c32b-105">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="0c32b-106">子要素</span><span class="sxs-lookup"><span data-stu-id="0c32b-106">Child elements</span></span>

|  <span data-ttu-id="0c32b-107">要素</span><span class="sxs-lookup"><span data-stu-id="0c32b-107">Element</span></span>  |  <span data-ttu-id="0c32b-108">必須</span><span class="sxs-lookup"><span data-stu-id="0c32b-108">Required</span></span>  |  <span data-ttu-id="0c32b-109">説明</span><span class="sxs-lookup"><span data-stu-id="0c32b-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="0c32b-110">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="0c32b-110">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="0c32b-111">はい</span><span class="sxs-lookup"><span data-stu-id="0c32b-111">Yes</span></span>  | <span data-ttu-id="0c32b-112">カスタム関数によって使用される JSON ファイルのリソース ID を持つ文字列。</span><span class="sxs-lookup"><span data-stu-id="0c32b-112">String with the resource id of the JSON file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="0c32b-113">例</span><span class="sxs-lookup"><span data-stu-id="0c32b-113">Example</span></span>

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
