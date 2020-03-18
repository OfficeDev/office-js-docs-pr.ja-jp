---
title: マニフェスト ファイルの Script 要素
description: Script 要素は、カスタム関数が Excel で使用するスクリプト設定を定義します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: f05fc85bd0454c340f4352bb73f299b9e7730224
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720415"
---
# <a name="script-element"></a><span data-ttu-id="4f9fc-103">Script 要素</span><span class="sxs-lookup"><span data-stu-id="4f9fc-103">Script element</span></span>

<span data-ttu-id="4f9fc-104">Excel でカスタム関数によって使用されるスクリプトの設定を定義します。</span><span class="sxs-lookup"><span data-stu-id="4f9fc-104">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="4f9fc-105">属性</span><span class="sxs-lookup"><span data-stu-id="4f9fc-105">Attributes</span></span>

<span data-ttu-id="4f9fc-106">なし</span><span class="sxs-lookup"><span data-stu-id="4f9fc-106">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="4f9fc-107">子要素</span><span class="sxs-lookup"><span data-stu-id="4f9fc-107">Child elements</span></span>

|<span data-ttu-id="4f9fc-108">要素</span><span class="sxs-lookup"><span data-stu-id="4f9fc-108">Elements</span></span>  |  <span data-ttu-id="4f9fc-109">必須</span><span class="sxs-lookup"><span data-stu-id="4f9fc-109">Required</span></span>  |  <span data-ttu-id="4f9fc-110">説明</span><span class="sxs-lookup"><span data-stu-id="4f9fc-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="4f9fc-111">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="4f9fc-111">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="4f9fc-112">はい</span><span class="sxs-lookup"><span data-stu-id="4f9fc-112">Yes</span></span>  | <span data-ttu-id="4f9fc-113">カスタム関数によって使用される JavaScript ファイルのリソース ID を持つ文字列。</span><span class="sxs-lookup"><span data-stu-id="4f9fc-113">String with resource id of the JavaScript file used by custom functions.</span></span>|

## <a name="example"></a><span data-ttu-id="4f9fc-114">例</span><span class="sxs-lookup"><span data-stu-id="4f9fc-114">Example</span></span>

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```
