---
title: マニフェスト ファイルの Page 要素
description: Page 要素は、カスタム関数で使用する HTML ページ設定を定義Excel。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 0d10ed04b73dfd786d50150dd8d01629f826cde17cb5ed1a7633d7319b5d6490
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57090166"
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
