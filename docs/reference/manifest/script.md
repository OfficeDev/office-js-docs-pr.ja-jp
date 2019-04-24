---
title: マニフェスト ファイルの Script 要素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 8352ada0eeb6af071d5f20f750dcdeaefe31e918
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450437"
---
# <a name="script-element"></a><span data-ttu-id="60a74-102">Script 要素</span><span class="sxs-lookup"><span data-stu-id="60a74-102">Script element</span></span>

<span data-ttu-id="60a74-103">Excel でカスタム関数によって使用されるスクリプトの設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="60a74-103">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="60a74-104">属性</span><span class="sxs-lookup"><span data-stu-id="60a74-104">Attributes</span></span>

<span data-ttu-id="60a74-105">なし</span><span class="sxs-lookup"><span data-stu-id="60a74-105">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="60a74-106">子要素</span><span class="sxs-lookup"><span data-stu-id="60a74-106">Child elements</span></span>

|<span data-ttu-id="60a74-107">要素</span><span class="sxs-lookup"><span data-stu-id="60a74-107">Elements</span></span>  |  <span data-ttu-id="60a74-108">必須</span><span class="sxs-lookup"><span data-stu-id="60a74-108">Required</span></span>  |  <span data-ttu-id="60a74-109">説明</span><span class="sxs-lookup"><span data-stu-id="60a74-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="60a74-110">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="60a74-110">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="60a74-111">はい</span><span class="sxs-lookup"><span data-stu-id="60a74-111">Yes</span></span>  | <span data-ttu-id="60a74-112">カスタム関数によって使用される JavaScript ファイルのリソース ID を持つ文字列。</span><span class="sxs-lookup"><span data-stu-id="60a74-112">String with resource id of the JavaScript file used by custom functions.</span></span>|

## <a name="example"></a><span data-ttu-id="60a74-113">例</span><span class="sxs-lookup"><span data-stu-id="60a74-113">Example</span></span>

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```
