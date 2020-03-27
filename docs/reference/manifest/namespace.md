---
title: マニフェスト ファイルの Namespace 要素
description: Namespace 要素は、Excel でカスタム関数によって使用される名前空間を定義します。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: eabd73d3be98271c81723787dd3d1bdb6ee2ebcd
ms.sourcegitcommit: 315a648cce38609c3e1c92bd4a339e268f8a2e1d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2020
ms.locfileid: "42978670"
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
