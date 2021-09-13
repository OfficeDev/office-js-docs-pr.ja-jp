---
title: マニフェスト ファイルの Namespace 要素
description: Namespace 要素は、カスタム関数がカスタム関数で使用する名前空間を定義Excel。
ms.date: 10/09/2018
ms.localizationpriority: medium
ms.openlocfilehash: 3a5afed3d55bde7e9735df534215f96ae1ba7bd3
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154915"
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
