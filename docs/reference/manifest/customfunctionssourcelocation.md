---
title: マニフェスト ファイル内のカスタム関数の SourceLocation 要素
description: Excel でカスタム関数によって使用される Script または Page 要素が必要とするリソースの場所を定義します。
ms.date: 08/07/2020
localization_priority: Normal
ms.openlocfilehash: 6001673f1954a4af2de66ff7611069c3fb402a13
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771383"
---
# <a name="sourcelocation-element-custom-functions"></a>SourceLocation 要素 (カスタム関数)

Excel でカスタム関数によって使用される Script または Page 要素が必要とするリソースの場所を定義します。

## <a name="attributes"></a>属性

| 属性 | 必須 | 説明                                                                          |
|-----------|----------|--------------------------------------------------------------------------------------|
| resid     | はい      | マニフェストの &lt;Resources&gt; セクションで定義される URL リソースの名前。 使用できる文字数は 32 文字です。 |

## <a name="child-elements"></a>子要素

なし

## <a name="example"></a>例

```xml
<SourceLocation resid="pageURL"/>
```
