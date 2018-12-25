---
title: マニフェスト ファイルの Script 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 95e4cbadc35302b4f76108e0ff2a51d31ca89aac
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433139"
---
# <a name="script-element"></a><span data-ttu-id="5e200-102">Script 要素</span><span class="sxs-lookup"><span data-stu-id="5e200-102">Script element</span></span>

<span data-ttu-id="5e200-103">Excel でカスタム関数によって使用されるスクリプトの設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="5e200-103">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="5e200-104">属性</span><span class="sxs-lookup"><span data-stu-id="5e200-104">Attributes</span></span>

<span data-ttu-id="5e200-105">なし</span><span class="sxs-lookup"><span data-stu-id="5e200-105">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="5e200-106">子要素</span><span class="sxs-lookup"><span data-stu-id="5e200-106">Child elements</span></span>

|<span data-ttu-id="5e200-107">要素</span><span class="sxs-lookup"><span data-stu-id="5e200-107">Elements</span></span>  |  <span data-ttu-id="5e200-108">必須</span><span class="sxs-lookup"><span data-stu-id="5e200-108">Required</span></span>  |  <span data-ttu-id="5e200-109">説明</span><span class="sxs-lookup"><span data-stu-id="5e200-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="5e200-110">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="5e200-110">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="5e200-111">はい</span><span class="sxs-lookup"><span data-stu-id="5e200-111">Yes</span></span>  | <span data-ttu-id="5e200-112">カスタム関数によって使用される JavaScript ファイルのリソース ID を持つ文字列。</span><span class="sxs-lookup"><span data-stu-id="5e200-112">String with resource id of the JavaScript file used by custom functions.</span></span>|

## <a name="example"></a><span data-ttu-id="5e200-113">例</span><span class="sxs-lookup"><span data-stu-id="5e200-113">Example</span></span>

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```
