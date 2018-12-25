---
title: マニフェスト ファイルの Namespace 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 8000ea5774b38dd038888c686a33127a2d5bc482
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432327"
---
# <a name="namespace-element"></a>Namespace 要素

Excel でカスタム関数によって使用される名前空間を定義します。

## <a name="attributes"></a>属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  **resid="namespace"**  |  はい  | [Resources](resources.md) 要素で指定されているカスタム関数の ShortStrings のタイトルと一致する必要があります。 |

## <a name="child-elements"></a>子要素

なし

## <a name="example"></a>例

```xml
<Namespace resid="namespace" />
```
