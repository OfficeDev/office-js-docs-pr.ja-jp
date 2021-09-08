---
title: Excel アドインの空白値と null 値
description: オブジェクト モデルのメソッドとプロパティで空白の null 値Excelする方法について説明します。
ms.date: 09/03/2020
localization_priority: Normal
ms.openlocfilehash: 3f38569f7342bb88c52ce424db426bfa7939be5e
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937518"
---
# <a name="blank-and-null-values-in-excel-add-ins"></a>Excel アドインの空白値と null 値

`null` と空の文字列は、Excel JavaScript API では特別な意味を持ちます。 これらは、空のセル、書式設定なし、既定値を表すために使用されます。 このセクションでは、プロパティの取得や設定を行うときに `null` や空の文字列を使用する方法について詳しく説明します。

## <a name="null-input-in-2-d-array"></a>2 次元配列での null の入力

Excel では、範囲は 2 次元配列で表され、最初のディメンションは行、2 番目のディメンションは列を示します。 範囲内の特定のセルだけに値、数値書式、または数式を設定するには、2 次元配列内のそのセルに値、数値書式、または数式を指定し、2 次元配列内のその他のすべてのセルに `null` を指定します。

たとえば、範囲内の 1 つのセルの数値書式を更新し、範囲内の他のセルすべての既存の数値書式を保持する場合、更新するセルに新しい数値書式を指定し、他のセルすべてに `null` を指定します。 次のコード スニペットでは、範囲内の 4 番目のセルに新しい数値書式を設定し、その前の 3 つのセルについては数値書式を変更せずに保持します。

```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```

## <a name="null-input-for-a-property"></a>プロパティに対する null の入力

`null` は単一プロパティに有効な入力ではありません。たとえば、次のコード スニペットは、範囲の `values` プロパティを `null` に設定できないため無効です。

```js
range.values = null; // This is not a valid snippet. 
```

同様に、次のコード スニペットは、`null` が `color` プロパティで有効な値ではないため無効です。

```js
range.format.fill.color =  null;  // This is not a valid snippet. 
```

## <a name="null-property-values-in-the-response"></a>応答内の null プロパティ値

指定の範囲に複数の値がある場合、`size` および `color` などの書式設定プロパティでは、応答に `null` 値が含まれます。 たとえば、範囲を取得してその `format.font.color` プロパティを読み込む場合:

* 範囲内のすべてのセルのフォントの色が同じ場合、`range.format.font.color` がその色を指定します。
* 範囲内に複数のフォントの色がある場合、`range.format.font.color` は `null` です。

## <a name="blank-input-for-a-property"></a>プロパティに対する空白の入力

プロパティに空白の値 (`''` の間にスペースのない 2 つの引用符) を指定すると、プロパティをクリアまたはリセットする指示として解釈されます。例:

* 範囲の `values` プロパティに空白の値を指定すると、範囲のコンテンツはクリアされます。
* `numberFormat` プロパティに空白の値を指定すると、数値書式は `General` にリセットされます。
* `formula` プロパティと `formulaLocale` プロパティに空白の値を指定すると、数式の値はクリアされます。

## <a name="blank-property-values-in-the-response"></a>応答内の空白のプロパティ値

読み取り操作では、応答内の空白のプロパティ値 (`''` の間にスペースのない、2 つの引用符) は、セルにデータまたは値がないことを示します。 次の 1 番目の例では、範囲内の最初と最後のセルにデータがありません。 2 番目の例では、範囲内の最初の 2 つのセルに数式がありません。

```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```

```js
range.formula = [['', '', '=Rand()']];
```
