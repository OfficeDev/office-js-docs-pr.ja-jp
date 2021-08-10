---
title: マニフェスト ファイルの Metadata 要素
description: Metadata 要素は、カスタム関数がカスタム 関数で使用するメタデータ設定を定義Excel。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: d6b7af8988265baf8fbdea51e1414646f88868ede76ed7194c47db1eb874608d
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57092491"
---
# <a name="metadata-element"></a>MetaData 要素

Excel でカスタム関数によって使用されるメタデータの設定を定義します。

## <a name="attributes"></a>属性

なし

## <a name="child-elements"></a>子要素

|  要素  |  必須  |  説明  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  はい  | カスタム関数によって使用される JSON ファイルのリソース ID を持つ文字列。 |

## <a name="example"></a>例

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
