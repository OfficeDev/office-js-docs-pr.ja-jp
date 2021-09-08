---
title: マニフェスト ファイルの Metadata 要素
description: Metadata 要素は、カスタム関数がカスタム 関数で使用するメタデータ設定を定義Excel。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 01be124b5526ce8328e0a20b8ff7d21ba6da96bc
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937692"
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
