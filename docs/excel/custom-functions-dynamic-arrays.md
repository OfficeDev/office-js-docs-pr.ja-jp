---
ms.date: 05/11/2020
description: アドイン内のカスタム関数から複数の結果Office Excel返します。
title: カスタム関数から複数の結果を返す
localization_priority: Normal
ms.openlocfilehash: b7df6b2c5ca3dca24615a61e11277ac36b42c0df
ms.sourcegitcommit: 69f6492de8a4c91e734250c76681c44b3f349440
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/03/2021
ms.locfileid: "58868450"
---
# <a name="return-multiple-results-from-your-custom-function"></a>カスタム関数から複数の結果を返す

カスタム関数から複数の結果を返し、隣接するセルに返されます。 この動作は、スピルと呼ばれる。 カスタム関数が結果の配列を返す場合は、動的配列式と呼ばれる。 動的配列の数式の詳細については、「Excel動的配列とスピル配列の[動作」を参照してください](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531)。

次の図は、関数が `SORT` 隣接するセルにこぼれる方法を示しています。 カスタム関数は、このような複数の結果を返す場合があります。

![複数の結果を複数のセルに表示する 'SORT' 関数のスクリーン ショット。](../images/dynamic-array-spill.png)

動的配列数式であるカスタム関数を作成するには、値の 2 次元配列を返す必要があります。 結果が既に値を持つ隣接セルにこぼれる場合、数式にエラーが表示 `#SPILL!` されます。

次の例は、流出する動的配列を返す方法を示しています。

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

次の例は、右にこぼれる動的配列を返す方法を示しています。 

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

次の例は、ダウンと右の両方をこぼす動的配列を返す方法を示しています。

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

- [動的配列とスピル配列の動作](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531)
- [カスタム関数Excelオプション](custom-functions-parameter-options.md)