---
title: マニフェスト ファイルの Namespace 要素
description: Namespace 要素は、Excel でカスタム関数によって使用される名前空間を定義します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: f4b3510c6c137bd303af8a3eaac8ebe66c5f4dc7
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612235"
---
# <a name="namespace-element"></a>Namespace 要素

Excel でカスタム関数によって使用される名前空間を定義します。

## <a name="attributes"></a>属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  **resid="namespace"**  |  いいえ  | [Resources](resources.md) 要素で指定されているカスタム関数の ShortStrings のタイトルと一致する必要があります。 |

## <a name="child-elements"></a>子要素

なし

## <a name="example"></a>例

```xml
<Namespace resid="namespace" />
```
