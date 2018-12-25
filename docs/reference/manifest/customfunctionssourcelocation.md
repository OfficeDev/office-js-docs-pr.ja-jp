---
title: マニフェスト ファイルの SourceLocation 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 878a8184984e31fdbcf46192a2f56507edaf4b37
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432408"
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