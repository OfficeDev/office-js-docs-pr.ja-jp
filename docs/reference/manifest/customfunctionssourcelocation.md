---
title: マニフェストファイル内のカスタム関数の SourceLocation 要素
description: Excel でカスタム関数によって使用される Script または Page 要素が必要とするリソースの場所を定義します。
ms.date: 08/07/2020
localization_priority: Normal
ms.openlocfilehash: 1c509987b0ce7948a63fa8ad51f7cf9c84144c5f
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641383"
---
# <a name="sourcelocation-element-custom-functions"></a>SourceLocation 要素 (カスタム関数)

Excel でカスタム関数によって使用される Script または Page 要素が必要とするリソースの場所を定義します。

## <a name="attributes"></a>属性

| 属性 | 必須 | 説明                                                                          |
|-----------|----------|--------------------------------------------------------------------------------------|
| resid     | はい      | マニフェストの &lt;Resources&gt; セクションで定義される URL リソースの名前。 |

## <a name="child-elements"></a>子要素

なし

## <a name="example"></a>例

```xml
<SourceLocation resid="pageURL"/>
```
