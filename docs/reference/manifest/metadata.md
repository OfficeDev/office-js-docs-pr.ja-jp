---
title: マニフェスト ファイルの Metadata 要素
description: Metadata 要素は、Excel でカスタム関数によって使用されるメタデータ設定を定義します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 01be124b5526ce8328e0a20b8ff7d21ba6da96bc
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611765"
---
# <a name="metadata-element"></a><span data-ttu-id="1ef24-103">MetaData 要素</span><span class="sxs-lookup"><span data-stu-id="1ef24-103">Metadata element</span></span>

<span data-ttu-id="1ef24-104">Excel でカスタム関数によって使用されるメタデータの設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="1ef24-104">Defines the metadata settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="1ef24-105">属性</span><span class="sxs-lookup"><span data-stu-id="1ef24-105">Attributes</span></span>

<span data-ttu-id="1ef24-106">なし</span><span class="sxs-lookup"><span data-stu-id="1ef24-106">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="1ef24-107">子要素</span><span class="sxs-lookup"><span data-stu-id="1ef24-107">Child elements</span></span>

|  <span data-ttu-id="1ef24-108">要素</span><span class="sxs-lookup"><span data-stu-id="1ef24-108">Element</span></span>  |  <span data-ttu-id="1ef24-109">必須</span><span class="sxs-lookup"><span data-stu-id="1ef24-109">Required</span></span>  |  <span data-ttu-id="1ef24-110">説明</span><span class="sxs-lookup"><span data-stu-id="1ef24-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="1ef24-111">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="1ef24-111">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="1ef24-112">はい</span><span class="sxs-lookup"><span data-stu-id="1ef24-112">Yes</span></span>  | <span data-ttu-id="1ef24-113">カスタム関数によって使用される JSON ファイルのリソース ID を持つ文字列。</span><span class="sxs-lookup"><span data-stu-id="1ef24-113">String with the resource id of the JSON file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="1ef24-114">例</span><span class="sxs-lookup"><span data-stu-id="1ef24-114">Example</span></span>

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
