---
title: マニフェスト ファイルの Page 要素
description: Page 要素は、カスタム関数で使用する HTML ページ設定を定義Excel。
ms.date: 10/09/2018
ms.localizationpriority: medium
ms.openlocfilehash: 6bde3ba86270874b1d9059b2f1c44952241bf00f
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154861"
---
# <a name="page-element"></a>Page 要素

Excel でカスタム関数によって使用される HTML ページの設定を定義します。

## <a name="attributes"></a>属性

なし

## <a name="child-elements"></a>子要素

|  要素  |  必須  |  説明  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  はい  | カスタム関数によって使用される HTML ファイルのリソース ID を持つ文字列。 |

## <a name="example"></a>例

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
