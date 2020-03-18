---
ms.date: 12/18/2019
description: Office Excel アドインで、カスタム関数から複数の結果を返します。
title: カスタム関数から複数の結果を返す
localization_priority: Normal
ms.openlocfilehash: a2632c621071f0cbc55f545847d9e9392d884b90
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719295"
---
# <a name="return-multiple-results-from-your-custom-function"></a>カスタム関数から複数の結果を返す

隣接するセルに返される、カスタム関数から複数の結果を返すことができます。 この動作は spilling と呼ばれます。 カスタム関数が結果の配列を返す場合は、動的配列数式と呼ばれます。 Excel の動的配列数式の詳細については、「動的配列」[および「こぼれた配列の動作](https://support.office.com/article/dynamic-arrays-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531)」を参照してください。

次の図は、関数`SORT`が隣接するセルにどのように分解されるかを示しています。 カスタム関数は、次のような複数の結果を返すこともできます。

![複数のセルに複数の結果を表示する ' SORT ' 関数のスクリーンショット。](../images/dynamic-array-spill.png)

動的配列数式であるカスタム関数を作成するには、値の2次元配列を返す必要があります。 結果が、既に値を持つ隣接するセルにスピルされる場合、 `#SPILL!`数式はエラーを表示します。

次の例は、分解した動的配列を返す方法を示しています。

```javascript
/**
 * Get text values that spill down.
 * @customfunction
 * @returns {string[][]} A dynamic array with multiple results.
 */
function spillDown() {
  return [['first'], ['second'], ['third']];
}
```

次の例は、右に液体をこぼれた動的配列を返す方法を示しています。 

```javascript
/**
 * Get text values that spill to the right.
 * @customfunction
 * @returns {string[][]} A dynamic array with multiple results.
 */
function spillRight() {
  return [['first', 'second', 'third']];
}
```

次の例は、右下の配列を返す方法を示しています。

```javascript
/**
 * Get text values that spill both right and down.
 * @customfunction
 * @returns {string[][]} A dynamic array with multiple results.
 */
function spillRectangle() {
  return [
    ['apples', 1, 'pounds'],
    ['oranges', 3, 'pounds'],
    ['pears', 5, 'crates']
  ];
}
```

## <a name="see-also"></a>関連項目

- [動的配列とこぼれた配列の動作](https://support.office.com/article/dynamic-arrays-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531)
- [Excel カスタム関数のオプション](custom-functions-parameter-options.md)