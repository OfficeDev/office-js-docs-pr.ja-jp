---
title: マニフェスト ファイルの Metadata 要素
description: Metadata 要素は、カスタム関数がカスタム 関数で使用するメタデータ設定を定義Excel。
ms.date: 10/09/2018
ms.localizationpriority: medium
ms.openlocfilehash: 6f58b00bb13bde1e2b1742462716119b8b6d369d
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152866"
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
