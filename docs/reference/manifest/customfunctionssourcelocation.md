---
title: マニフェスト ファイル内のカスタム関数の SourceLocation 要素
description: Excel でカスタム関数によって使用される Script または Page 要素が必要とするリソースの場所を定義します。
ms.date: 08/07/2020
localization_priority: Normal
ms.openlocfilehash: b18a340d4dd4403b1e5fd2c7d8868a820eef5a241ac3d666926d8f2cb49fcc09
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57098303"
---
# <a name="sourcelocation-element-custom-functions"></a>SourceLocation 要素 (カスタム関数)

Excel でカスタム関数によって使用される Script または Page 要素が必要とするリソースの場所を定義します。

## <a name="attributes"></a>属性

| 属性 | 必須 | 説明                                                                          |
|-----------|----------|--------------------------------------------------------------------------------------|
| resid     | はい      | マニフェストの &lt;Resources&gt; セクションで定義される URL リソースの名前。 32 文字以内で指定できます。 |

## <a name="child-elements"></a>子要素

なし

## <a name="example"></a>例

```xml
<SourceLocation resid="pageURL"/>
```
