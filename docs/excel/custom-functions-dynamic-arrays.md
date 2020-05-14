---
ms.date: 12/18/2019
description: Office Excel アドインで、カスタム関数から複数の結果を返します。
title: カスタム関数から複数の結果を返す
localization_priority: Normal
ms.openlocfilehash: 753755b481ab3db0de711c80ef082aedc82177ae
ms.sourcegitcommit: 682d18c9149b1153f9c38d28e2a90384e6a261dc
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/13/2020
ms.locfileid: "44217838"
---
# <a name="return-multiple-results-from-your-custom-function"></a>カスタム関数から複数の結果を返す

隣接するセルに返される、カスタム関数から複数の結果を返すことができます。 この動作は spilling と呼ばれます。 カスタム関数が結果の配列を返す場合は、動的配列数式と呼ばれます。 Excel の動的配列数式の詳細については、「動的配列」[および「こぼれた配列の動作](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531)」を参照してください。

次の図は、関数が隣接するセルにどのように分解されるかを示して `SORT` います。 カスタム関数は、次のような複数の結果を返すこともできます。

![複数のセルに複数の結果を表示する ' SORT ' 関数のスクリーンショット。](../images/dynamic-array-spill.png)

動的配列数式であるカスタム関数を作成するには、値の2次元配列を返す必要があります。 結果が、既に値を持つ隣接するセルにスピルされる場合、数式はエラーを表示し `#SPILL!` ます。

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

- [動的配列とこぼれた配列の動作](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531)
- [Excel カスタム関数のオプション](custom-functions-parameter-options.md)