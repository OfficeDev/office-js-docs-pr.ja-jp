---
title: マニフェスト ファイル内のカスタム関数の SourceLocation 要素
description: Excel でカスタム関数によって使用される Script または Page 要素が必要とするリソースの場所を定義します。
ms.date: 08/07/2020
ms.localizationpriority: medium
ms.openlocfilehash: 84d5607fbb02c1925137e1a143b7715c7c87c6fa
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151663"
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
