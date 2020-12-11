---
ms.date: 12/09/2020
description: Excel の範囲、オプションのパラメーター、呼び出しコンテキストなど、カスタム関数内で異なるパラメーターを使用する方法について説明します。
title: Excel カスタム関数のオプション
localization_priority: Normal
ms.openlocfilehash: 9f43955324c148a0af030fb796b82f6d72f429c5
ms.sourcegitcommit: b300e63a96019bdcf5d9f856497694dbd24bfb11
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/11/2020
ms.locfileid: "49624667"
---
# <a name="custom-functions-parameter-options"></a>カスタム関数のパラメーター オプション

カスタム関数は、さまざまなパラメーター オプションを使用して構成できます。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="optional-parameters"></a>オプションのパラメーター

ユーザーが Excel で関数を呼び出すと、角かっこで囲まれた省略可能なパラメーターが表示されます。 次のサンプルでは、add 関数は必要に応じて 3 番目の数値を追加できます。 この関数は Excel と `=CONTOSO.ADD(first, second, [third])` 同様に表示されます。

#### <a name="javascript"></a>[JavaScript](#tab/javascript)

```js
/**
 * Calculates the sum of the specified numbers
 * @customfunction
 * @param {number} first First number.
 * @param {number} second Second number.
 * @param {number} [third] Third number to add. If omitted, third = 0.
 * @returns {number} The sum of the numbers.
 */
function add(first, second, third) {
  if (third === null) {
    third = 0;
  }
  return first + second + third;
}
```

#### <a name="typescript"></a>[TypeScript](#tab/typescript)

```typescript
/**
 * Calculates the sum of the specified numbers
 * @customfunction
 * @param first First number.
 * @param second Second number.
 * @param [third] Third number to add. If omitted, third = 0.
 * @returns The sum of the numbers.
 */
function add(first: number, second: number, third?: number): number {
  if (third === null) {
    third = 0;
  }
  return first + second + third;
}
```

---

> [!NOTE]
> 省略可能なパラメーターに値を指定しない場合、Excel によって値が割り当てらされます `null` 。 つまり、TypeScript の既定で初期化されたパラメーターは期待通り動作しません。 この構文は 0 に初期化 `function add(first:number, second:number, third=0):number` されないので使用 `third` してください。 代わりに、前の例で示した TypeScript 構文を使用します。

1 つ以上のオプション パラメーターを含む関数を定義する場合は、オプションパラメーターが null の場合の処理を指定します。 次の例の `zipCode` と `dayOfWeek` は、どちらも `getWeatherReport` 関数の省略可能なパラメーターです。 パラメーターが `zipCode` null の場合、既定値はに設定されます `98052` 。 パラメーターが `dayOfWeek` null の場合は、水曜日に設定されます。

#### <a name="javascript"></a>[JavaScript](#tab/javascript)

```js
/**
 * Gets a weather report for a specified zipCode and dayOfWeek
 * @customfunction
 * @param {number} [zipCode] Zip code. If omitted, zipCode = 98052.
 * @param {string} [dayOfWeek] Day of the week. If omitted, dayOfWeek = Wednesday.
 * @returns {string} Weather report for the day of the week in that zip code.
 */
function getWeatherReport(zipCode, dayOfWeek) {
  if (zipCode === null) {
    zipCode = 98052;
  }

  if (dayOfWeek === null) {
    dayOfWeek = "Wednesday";
  }

  // Get weather report for specified zipCode and dayOfWeek.
  // ...
}
```

#### <a name="typescript"></a>[TypeScript](#tab/typescript)

```typescript
/**
 * Gets a weather report for a specified zipCode and dayOfWeek
 * @customfunction
 * @param zipCode Zip code. If omitted, zipCode = 98052.
 * @param [dayOfWeek] Day of the week. If omitted, dayOfWeek = Wednesday.
 * @returns Weather report for the day of the week in that zip code.
 */
function getWeatherReport(zipCode?: number, dayOfWeek?: string): string {
  if (zipCode === null) {
    zipCode = 98052;
  }

  if (dayOfWeek === null) {
    dayOfWeek = "Wednesday";
  }

  // Get weather report for specified zipCode and dayOfWeek.
  // ...
}
```

---

## <a name="range-parameters"></a>範囲パラメーター

カスタム関数は、入力パラメーターとしてセル データの範囲を受け入れる場合があります。 関数は、データの範囲を返す場合があります。 Excel はセル データの範囲を 2 次元配列として渡します。

例えば、関数が Excel に保存されている数値の範囲から 2 番目に大きい値を返すとします。 次の関数はパラメーターを受け入れ、JSDOC 構文はパラメーターのプロパティをこの関数の JSON メタデータ `values` `number[][]` `dimensionality` `matrix` に設定します。 

```js
/**
 * Returns the second highest value in a matrixed range of values.
 * @customfunction
 * @param {number[][]} values Multiple ranges of values.
 */
function secondHighest(values) {
  let highest = values[0][0],
    secondHighest = values[0][0];
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      if (values[i][j] >= highest) {
        secondHighest = highest;
        highest = values[i][j];
      } else if (values[i][j] >= secondHighest) {
        secondHighest = values[i][j];
      }
    }
  }
  return secondHighest;
}
```

## <a name="repeating-parameters"></a>繰り返しパラメーター

繰り返しパラメーターを使用すると、ユーザーは関数に一連のオプションの引数を入力できます。 関数が呼び出される場合、値はパラメーターの配列で提供されます。 パラメーター名の最後が数値の場合、各引数の数値は徐々に増加します。次に例を示します `ADD(number1, [number2], [number3],…)` 。 これは、組み込みの Excel 関数で使用される規則に一致します。

次の関数は、入力されている場合、数値、セル アドレス、および範囲の合計を合計します。

```TS
/**
* The sum of all of the numbers.
* @customfunction
* @param operands A number (such as 1 or 3.1415), a cell address (such as A1 or $E$11), or a range of cell addresses (such as B3:F12)
*/

function ADD(operands: number[][][]): number {
  let total: number = 0;

  operands.forEach(range => {
    range.forEach(row => {
      row.forEach(num => {
        total += num;
      });
    });
  });

  return total;
}
```

この関数は Excel `=CONTOSO.ADD([operands], [operands]...)` ブックに表示されます。

<img alt="The ADD custom function being entered into cell of an Excel worksheet" src="../images/operands.png" />

### <a name="repeating-single-value-parameter"></a>繰り返し単一値パラメーター

繰り返し単一値パラメーターを使用すると、複数の単一の値を渡す可能性があります。 たとえば、ユーザーは ADD(1,B2,3) と入力できます。 次のサンプルは、1 つの値パラメーターを宣言する方法を示しています。

```JS
/**
 * @customfunction
 * @param {number[]} singleValue An array of numbers that are repeating parameters.
 */
function addSingleValue(singleValue) {
  let total = 0;
  singleValue.forEach(value => {
    total += value;
  })

  return total;
}
```

### <a name="single-range-parameter"></a>単一の範囲パラメーター

1 つの範囲パラメーターは技術的には繰り返しパラメーターではなく、宣言が繰り返しパラメーターと非常に似ているため、ここに含まれています。 ユーザーには ADD(A2:B3) と表示され、Excel から 1 つの範囲が渡されます。 次のサンプルは、1 つの範囲パラメーターを宣言する方法を示しています。

```JS
/**
 * @customfunction
 * @param {number[][]} singleRange
 */
function addSingleRange(singleRange) {
  let total = 0;
  singleRange.forEach(setOfSingleValues => {
    setOfSingleValues.forEach(value => {
      total += value;
    })
  })
  return total;
}
```

### <a name="repeating-range-parameter"></a>繰り返し範囲パラメーター

繰り返し範囲パラメーターを使用すると、複数の範囲または数値を渡できます。 たとえば、ユーザーは ADD(5,B2,C3,8,E5:E8) と入力できます。 繰り返し範囲は、通常、3 次元マトリックス `number[][][]` である型で指定されます。 サンプルについては、繰り返しパラメーター (#repeating-parameters) の一覧にあるメイン サンプルを参照してください。


### <a name="declaring-repeating-parameters"></a>繰り返しパラメーターの宣言
Typescript で、パラメーターが多次元パラメーターかどうかを示します。 たとえば  `ADD(values: number[])` 、1 次元配列を示し `ADD(values:number[][])` 、2 次元配列を示す場合などです。

JavaScript では、1 次元配列、2 次元配列、およびより多くの次元 `@param values {number[]}` `@param <name> {number[][]}` に使用します。

手書き JSON の場合は、JSON ファイルでパラメーターが指定されているのを確認し、パラメーターにマーク `"repeating": true` が付けられているか確認します `"dimensionality": matrix` 。

## <a name="invocation-parameter"></a>呼び出しパラメーター

すべてのカスタム関数には、最後の引数 `invocation` として引数が自動的に渡されます。 この引数は、呼び出し元セルのアドレスなど、追加のコンテキストを取得するために使用できます。 または、関数をキャンセルする関数ハンドラーなどの情報を Excel [に送信するために使用できます](custom-functions-web-reqs.md#make-a-streaming-function)。 パラメーターを宣言していなくても、カスタム関数にはこのパラメーターがあります。 この引数は、Excel のユーザーには表示されません。 カスタム関数で使用 `invocation` する場合は、最後のパラメーターとして宣言します。

次のコード サンプルでは、コンテキスト `invocation` が参照用に明示的に示されています。

```js
/**
 * Add two numbers.
 * @customfunction
 * @param {number} first First number.
 * @param {number} second Second number.
 * @returns {number} The sum of the two (or optionally three) numbers.
 */
function add(first, second, invocation) {
  return first + second;
}
```

## <a name="next-steps"></a>次の手順

カスタム関数で揮発性値 [を使用する方法について説明します](custom-functions-volatile.md)。

## <a name="see-also"></a>関連項目

* [カスタム関数でデータを受信して​​処理する](custom-functions-web-reqs.md)
* [カスタム関数用の JSON メタデータの自動生成](custom-functions-json-autogeneration.md)
* [カスタム関数の JSON メタデータを手動で作成する](custom-functions-json.md)
* [Excel でカスタム関数を作成する](custom-functions-overview.md)
* [Excel カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)
