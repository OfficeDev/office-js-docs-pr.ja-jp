---
title: マニフェスト ファイルの Namespace 要素
description: Namespace 要素は、カスタム関数がカスタム関数で使用する名前空間を定義Excel。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 3f20e744839d5791797642a9019f546922efd710367d5f23446241eebad0e48f
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57089680"
---
# <a name="namespace-element"></a>Namespace 要素

Excel でカスタム関数によって使用される名前空間を定義します。

## <a name="attributes"></a>属性

|  属性  |  必須  |  説明  |
|:-----|:-----|:-----|
|  **resid="namespace"**  |  いいえ  | [Resources](resources.md) 要素で指定されているカスタム関数の ShortStrings のタイトルと一致する必要があります。 32 文字以内で指定できます。 |

## <a name="child-elements"></a>子要素

なし

## <a name="example"></a>例

```xml
<Namespace resid="namespace" />
```
