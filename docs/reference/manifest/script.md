---
title: マニフェスト ファイルの Script 要素
description: Script 要素は、カスタム関数がカスタム 関数で使用するスクリプト設定をExcel。
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 51902864081e135faed778de1bc6fdee15d67490de8eabc9febf493cb0c09889
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57095045"
---
# <a name="script-element"></a>Script 要素

Excel でカスタム関数によって使用されるスクリプトの設定を定義します。

## <a name="attributes"></a>属性

なし

## <a name="child-elements"></a>子要素

|要素  |  必須  |  説明  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  はい  | カスタム関数によって使用される JavaScript ファイルのリソース ID を持つ文字列。|

## <a name="example"></a>例

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```
