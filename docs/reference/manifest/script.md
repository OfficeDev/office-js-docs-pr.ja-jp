---
title: マニフェスト ファイルの Script 要素
description: Script 要素は、カスタム関数が Excel で使用するスクリプト設定を定義します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 791f49f15673a029b982e40946f8cc90f02ba887
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608091"
---
# <a name="script-element"></a><span data-ttu-id="fc889-103">Script 要素</span><span class="sxs-lookup"><span data-stu-id="fc889-103">Script element</span></span>

<span data-ttu-id="fc889-104">Excel でカスタム関数によって使用されるスクリプトの設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="fc889-104">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="fc889-105">属性</span><span class="sxs-lookup"><span data-stu-id="fc889-105">Attributes</span></span>

<span data-ttu-id="fc889-106">なし</span><span class="sxs-lookup"><span data-stu-id="fc889-106">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="fc889-107">子要素</span><span class="sxs-lookup"><span data-stu-id="fc889-107">Child elements</span></span>

|<span data-ttu-id="fc889-108">要素</span><span class="sxs-lookup"><span data-stu-id="fc889-108">Elements</span></span>  |  <span data-ttu-id="fc889-109">必須</span><span class="sxs-lookup"><span data-stu-id="fc889-109">Required</span></span>  |  <span data-ttu-id="fc889-110">説明</span><span class="sxs-lookup"><span data-stu-id="fc889-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="fc889-111">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="fc889-111">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="fc889-112">はい</span><span class="sxs-lookup"><span data-stu-id="fc889-112">Yes</span></span>  | <span data-ttu-id="fc889-113">カスタム関数によって使用される JavaScript ファイルのリソース ID を持つ文字列。</span><span class="sxs-lookup"><span data-stu-id="fc889-113">String with resource id of the JavaScript file used by custom functions.</span></span>|

## <a name="example"></a><span data-ttu-id="fc889-114">例</span><span class="sxs-lookup"><span data-stu-id="fc889-114">Example</span></span>

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```
