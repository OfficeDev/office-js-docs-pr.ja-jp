---
ms.date: 03/08/2021
description: Excel の範囲、オプションのパラメーター、呼び出しコンテキストなど、カスタム関数内でさまざまなパラメーターを使用する方法について説明します。
title: Excel カスタム関数のオプション
localization_priority: Normal
ms.openlocfilehash: a168853eeb6a81cf3d0054cb3628b609ec283af7
ms.sourcegitcommit: d153f6d4c3e01d63ed24aa1349be16fa8ad51218
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/10/2021
ms.locfileid: "50613928"
---
# <a name="custom-functions-parameter-options"></a>カスタム関数パラメーター のオプション

カスタム関数は、さまざまなパラメーター オプションで構成できます。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="optional-parameters"></a>オプションのパラメーター

ユーザーが Excel で関数を呼び出すと、角かっこで囲まれた省略可能なパラメーターが表示されます。 次のサンプルでは、add 関数は必要に応じて 3 番目の数値を追加できます。 この関数は Excel のように `=CONTOSO.ADD(first, second, [third])` 表示されます。

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
> 省略可能なパラメーターに値を指定しない場合、Excel は値を割り当てします `null` 。 つまり、TypeScript の既定で初期化されたパラメーターは期待通り動作しません。 構文は 0 に初期化 `function add(first:number, second:number, third=0):number` されないので使用 `third` しない。 代わりに、前の例に示すように TypeScript 構文を使用します。

1 つ以上のオプション パラメーターを含む関数を定義する場合は、省略可能なパラメーターが null の場合の処理を指定します。 次の例の `zipCode` と `dayOfWeek` は、どちらも `getWeatherReport` 関数の省略可能なパラメーターです。 パラメーターが `zipCode` null の場合、既定値は に設定されます `98052` 。 パラメーターが `dayOfWeek` null の場合は、水曜日に設定されます。

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

## <a name="range-parameters"></a>Range パラメーター

カスタム関数は、セル データの範囲を入力パラメーターとして受け入れる場合があります。 関数は、データの範囲を返す場合があります。 Excel は、セル データの範囲を 2 次元配列として渡します。

例えば、関数が Excel に保存されている数値の範囲から 2 番目に大きい値を返すとします。 次の関数はパラメーターを受け入れ、JSDOC 構文はパラメーターのプロパティを `values` `number[][]` `dimensionality` `matrix` この関数の JSON メタデータに設定します。 

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

繰り返しパラメーターを使用すると、ユーザーは関数に対して一連のオプションの引数を入力できます。 関数が呼び出される場合、値はパラメーターの配列に指定されます。 パラメーター名が数値で終わると、各引数の数は増分的に増加します `ADD(number1, [number2], [number3],…)` 。. これは、組み込みの Excel 関数で使用される規則と一致します。

次の関数は、数値、セル アドレス、および範囲の合計を入力した場合に集計します。

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

この関数は `=CONTOSO.ADD([operands], [operands]...)` 、Excel ブックに表示されます。

<img alt="The ADD custom function being entered into cell of an Excel worksheet" src="../images/operands.png" />

### <a name="repeating-single-value-parameter"></a>単一値パラメーターの繰り返し

繰り返しの単一値パラメーターを使用すると、複数の単一の値を渡します。 たとえば、ユーザーは ADD(1,B2,3) と入力できます。 次のサンプルは、1 つの値パラメーターを宣言する方法を示しています。

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

1 つの範囲パラメーターは、技術的には繰り返しパラメーターではなく、宣言が繰り返しパラメーターと非常に似ているため、ここに含まれています。 これは、Excel から 1 つの範囲が渡される ADD(A2:B3) としてユーザーに表示されます。 次のサンプルは、1 つの範囲パラメーターを宣言する方法を示しています。

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

### <a name="repeating-range-parameter"></a>範囲パラメーターの繰り返し

繰り返し範囲パラメーターを使用すると、複数の範囲または数値を渡します。 たとえば、ユーザーは ADD(5,B2,C3,8,E5:E8) と入力できます。 繰り返し範囲は、通常、3 次元マトリックスとして型 `number[][][]` で指定されます。 サンプルについては、パラメーターの繰り返しに関する一覧のメイン [サンプルを参照してください](#repeating-parameters)。


### <a name="declaring-repeating-parameters"></a>繰り返しパラメーターの宣言
Typescript で、パラメーターが多次元かどうかを示します。 たとえば  `ADD(values: number[])` 、1 次元配列を示し、2 次元配列を示す場合など `ADD(values:number[][])` です。

JavaScript では、1 次元配列、2 次元配列など、より多くの次元 `@param values {number[]}` `@param <name> {number[][]}` に使用します。

手書き JSON の場合は、JSON ファイルのようにパラメーターを指定し、パラメーターにマークが付いている `"repeating": true` か確認してください `"dimensionality": matrix` 。

## <a name="invocation-parameter"></a>呼び出しパラメーター

すべてのカスタム関数は、明示的に宣言されていない場合でも、引数を最後の入力パラメーターとして自動的 `invocation` に渡されます。 この `invocation` パラメーターは、呼び出しオブジェクト [に対応](/javascript/api/custom-functions-runtime/customfunctions.invocation) します。 オブジェクトを使用して、カスタム関数を呼び出したセルのアドレスなど、追加の `Invocation` コンテキストを取得できます。 オブジェクトにアクセス `Invocation` するには、カスタム関数 `invocation` の最後のパラメーターとして宣言する必要があります。 

> [!NOTE]
> この `invocation` パラメーターは、Excel のユーザーのカスタム関数引数として表示されません。

次のサンプルは、パラメーターを使用して、カスタム関数を呼び出したセルの `invocation` アドレスを返す方法を示しています。 このサンプルでは、オブジェクト [の address](/javascript/api/custom-functions-runtime/customfunctions.invocation#address) プロパティを使用 `Invocation` します。 オブジェクトにアクセス `Invocation` するには、まず `CustomFunctions.Invocation` JSDoc でパラメーターとして宣言します。 次に `@requiresAddress` 、JSDoc でオブジェクトのプロパティにアクセス `address` する宣言を `Invocation` 行います。 最後に、関数内でプロパティを取得して返 `address` します。 

```js
/**
 * Return the address of the cell that invoked the custom function. 
 * @customfunction
 * @param {number} first First parameter.
 * @param {number} second Second parameter.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @requiresAddress 
 */
function getAddress(first, second, invocation) {
  var address = invocation.address;
  return address;
}
```

Excel では、オブジェクトのプロパティを呼び出すカスタム関数は、関数を呼び出したセルの形式に続く絶対アドレス `address` `Invocation` `SheetName!RelativeCellAddress` を返します。 たとえば、入力パラメーターがセル F6 の **[価格** ] というシートにある場合、返されるパラメーターのアドレス値は 、 になります `Prices!F6` 。 

この `invocation` パラメーターは、Excel に情報を送信するためにも使用できます。 詳細 [については、「ストリーミング機能を作成](custom-functions-web-reqs.md#make-a-streaming-function) する」を参照してください。

## <a name="detect-the-address-of-a-parameter"></a>パラメーターのアドレスを検出する

呼び出しパラメーター [と組み](#invocation-parameter)合わせて [、Invocation](/javascript/api/custom-functions-runtime/customfunctions.invocation) オブジェクトを使用して、カスタム関数入力パラメーターのアドレスを取得できます。 呼び出されると、 [オブジェクトの parameterAddresses](/javascript/api/custom-functions-runtime/customfunctions.invocation#parameterAddresses) プロパティを使用すると、関数は、すべての入力パラメーター `Invocation` のアドレスを返すことができます。 

これは、入力データ型が異なる場合に役立ちます。 入力パラメーターのアドレスを使用して、入力値の数値形式を確認できます。 必要に応じて、数値の形式を入力前に調整できます。 入力パラメーターのアドレスを使用して、入力値に後続の計算に関連する可能性のある関連プロパティが含されているかどうかを検出することもできます。 

>[!NOTE]
> 手動で作成した [JSON](custom-functions-json.md) メタデータを操作して、Yo Office ジェネレーターの代わりにパラメーター アドレスを返す場合、オブジェクトにはプロパティが設定されている必要があります。オブジェクトにはプロパティがに設定されている必要があります `options` `requiresParameterAddresses` `true` `result` `dimensionality` `matrix` 。

次のカスタム関数は、3 つの入力パラメーターを取り込み、各パラメーターのオブジェクトのプロパティを取得し、アドレス `parameterAddresses` `Invocation` を返します。 

```js
/**
 * Return the addresses of three parameters. 
 * @customfunction
 * @param {string} firstParameter First parameter.
 * @param {string} secondParameter Second parameter.
 * @param {string} thirdParameter Third parameter.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @returns {string[][]} The addresses of the parameters, as a 2-dimensional array. 
 * @requiresParameterAddresses
 */
function getParameterAddresses(firstParameter, secondParameter, thirdParameter, invocation) {
  var addresses = [
    [invocation.parameterAddresses[0]],
    [invocation.parameterAddresses[1]],
    [invocation.parameterAddresses[2]]
  ];
  return addresses;
}
```

プロパティを呼び出すカスタム関数が実行されると、関数を呼び出したセルの形式に従ってパラメーター アドレス `parameterAddresses` `SheetName!RelativeCellAddress` が返されます。 たとえば、入力パラメーターがセル D8 の **Costs** というシートにある場合、返されるパラメーターのアドレス値は `Costs!D8` . カスタム関数に複数のパラメーターが含まれていますが、複数のパラメーター アドレスが返される場合、返されるアドレスは、関数を呼び出したセルから垂直方向に降順に、複数のセルにわたってこぼれ落ちします。 

## <a name="next-steps"></a>次の手順

カスタム関数で揮発性値 [を使用する方法について説明します](custom-functions-volatile.md)。

## <a name="see-also"></a>関連項目

* [カスタム関数でデータを受信して​​処理する](custom-functions-web-reqs.md)
* [カスタム関数用の JSON メタデータの自動生成](custom-functions-json-autogeneration.md)
* [カスタム関数の JSON メタデータを手動で作成する](custom-functions-json.md)
* [Excel でカスタム関数を作成する](custom-functions-overview.md)
* [Excel カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)
