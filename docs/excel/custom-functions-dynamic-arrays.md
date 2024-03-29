---
ms.date: 05/11/2020
description: アドイン内のカスタム関数から複数の結果Office Excel返します。
title: カスタム関数から複数の結果を返す
ms.localizationpriority: medium
ms.openlocfilehash: 9c619b379bc39598bb325180d32ddcbced0ff664
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744350"
---
# <a name="return-multiple-results-from-your-custom-function"></a>カスタム関数から複数の結果を返す

カスタム関数から複数の結果を返し、隣接するセルに返されます。 この動作は、スピルと呼ばれる。 カスタム関数が結果の配列を返す場合は、動的配列式と呼ばれる。 動的配列式の詳細については、「Excel動的配列とスピル配列[の動作」を参照してください](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531)。

次の図は、関数が隣接 `SORT` するセルにこぼれる方法を示しています。 カスタム関数は、このような複数の結果を返す場合があります。

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
