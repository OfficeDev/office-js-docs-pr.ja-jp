---
title: マニフェスト ファイルの SourceLocation 要素
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: b2b78065fc8bde6fc827ddcb21e2bc700ed5bf49
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450689"
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
