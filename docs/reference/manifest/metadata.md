---
title: マニフェスト ファイルの Metadata 要素
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 79038fc13eba76176be19e484ffa57e64727bf94
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432663"
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
