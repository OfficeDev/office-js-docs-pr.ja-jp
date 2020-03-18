---
title: マニフェスト ファイルの SourceLocation 要素
description: Excel でカスタム関数によって使用される Script または Page 要素が必要とするリソースの場所を定義します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 56ebe122853c98a14c52d450bea31fecaefb15d3
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720688"
---
# <a name="sourcelocation-element"></a>SourceLocation 要素

Excel でカスタム関数によって使用される Script または Page 要素が必要とするリソースの場所を定義します。

## <a name="attributes"></a>属性

| **属性** | **必須** | **説明**                                                                      |
|---------------|--------------|--------------------------------------------------------------------------------------|
| resid         | はい          | マニフェストの &lt;Resources&gt; セクションで定義される URL リソースの名前。 |

## <a name="child-elements"></a>子要素

なし

## <a name="example"></a>例

```xml
<SourceLocation resid="pageURL"/>
```
