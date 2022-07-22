---
title: Excel カスタム関数のオプション
description: Excel 範囲、省略可能なパラメーター、呼び出しコンテキストなど、カスタム関数内でさまざまなパラメーターを使用する方法について説明します。
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: de86afc60d7d0b81820bd742e989e0ee7dd6970c
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958574"
---
# <a name="custom-functions-parameter-options"></a>カスタム関数パラメーター オプション

カスタム関数は、さまざまなパラメーター オプションを使用して構成できます。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="optional-parameters"></a>オプションのパラメーター

ユーザーが Excel で関数を呼び出すと、角かっこで囲まれた省略可能なパラメーターが表示されます。 次の例では、add 関数は必要に応じて 3 番目の数値を追加できます。 この関数は Excel のように `=CONTOSO.ADD(first, second, [third])` 表示されます。

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
> 省略可能なパラメーターに値が指定されていない場合は、Excel によって値 `null`が割り当てられます。 これは、TypeScript で既定で初期化されたパラメーターが期待どおりに機能しないことを意味します。 構文 `function add(first:number, second:number, third=0):number` は 0 に初期化 `third` されないため、使用しないでください。 代わりに、前の例に示すように TypeScript 構文を使用します。

1 つ以上の省略可能なパラメーターを含む関数を定義する場合は、省略可能なパラメーターが null の場合の動作を指定します。 次の例の `zipCode` と `dayOfWeek` は、どちらも `getWeatherReport` 関数の省略可能なパラメーターです。 パラメーターが `zipCode` null の場合、既定値は `98052`. パラメーターが `dayOfWeek` null の場合は、水曜日に設定されます。

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

カスタム関数は、入力パラメーターとしてセル データの範囲を受け入れる場合があります。 関数は、データ範囲を返すこともできます。 Excel はセル データの範囲を 2 次元配列として渡します。

例えば、関数が Excel に保存されている数値の範囲から 2 番目に大きい値を返すとします。 次の関数はパラメーター`values`を受け入れ、JSDOC 構文`number[][]`はパラメーターのプロパティを`matrix`この関数の `dimensionality` JSON メタデータに設定します。

```js
/**
 * Returns the second highest value in a matrixed range of values.
 * @customfunction
 * @param {number[][]} values Multiple ranges of values.
 */
function secondHighest(values) {
  let highest = values[0][0],
    secondHighest = values[0][0];
  for (let i = 0; i < values.length; i++) {
    for (let j = 0; j < values[i].length; j++) {
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

## <a name="repeating-parameters"></a>パラメーターの繰り返し

繰り返しパラメーターを使用すると、ユーザーは関数に対して一連の省略可能な引数を入力できます。 関数が呼び出されると、パラメーターの配列に値が指定されます。 パラメーター名が数値で終わると、各引数の数は増分的に増加します (例: `ADD(number1, [number2], [number3],…)`. これは、組み込みの Excel 関数で使用される規則と一致します。

次の関数は、入力した場合、数値、セル アドレス、範囲の合計を合計します。

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

この関数は Excel ブックに表示 `=CONTOSO.ADD([operands], [operands]...)` されます。

![Excel ワークシートのセルに入力される ADD カスタム関数](../images/operands.png)

### <a name="repeating-single-value-parameter"></a>単一値パラメーターの繰り返し

繰り返し 1 つの値パラメーターを使用すると、複数の単一の値を渡すことができます。 たとえば、ユーザーは ADD(1,B2,3) と入力できます。 次の例では、1 つの値パラメーターを宣言する方法を示します。

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

### <a name="single-range-parameter"></a>単一範囲パラメーター

単一の範囲パラメーターは、技術的には繰り返しパラメーターではありませんが、宣言は繰り返しパラメーターとよく似ているため、ここに含まれています。 ユーザーは ADD(A2:B3) として表示され、Excel から 1 つの範囲が渡されます。 次の例では、1 つの範囲パラメーターを宣言する方法を示します。

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

繰り返し範囲パラメーターを使用すると、複数の範囲または数値を渡すことができます。 たとえば、ユーザーは ADD(5,B2,C3,8,E5:E8) と入力できます。 繰り返し範囲は通常、3 次元マトリックスであるため、型 `number[][][]` で指定されます。 サンプルについては、 [パラメーターの繰り返し](#repeating-parameters)に関する一覧のメイン サンプルを参照してください。

### <a name="declaring-repeating-parameters"></a>繰り返しパラメーターを宣言する

Typescript で、パラメーターが多次元であることを示します。 たとえば、  `ADD(values: number[])` 1 次元配列を示し、 `ADD(values:number[][])` 2 次元配列などを示します。

JavaScript では、1 次元配列、 `@param <name> {number[][]}` 2 次元配列などに使用`@param values {number[]}`し、より多くのディメンションに使用します。

手書きの JSON の場合は、JSON ファイルのように `"repeating": true` パラメーターが指定されていることを確認し、パラメーターが `"dimensionality": matrix`.

## <a name="invocation-parameter"></a>呼び出しパラメーター

すべてのカスタム関数は、明示的に `invocation` 宣言されていない場合でも、引数が最後の入力パラメーターとして自動的に渡されます。 このパラメーターは`invocation`[、呼び出し](/javascript/api/custom-functions-runtime/customfunctions.invocation)オブジェクトに対応します。 オブジェクトは `Invocation` 、カスタム関数を呼び出したセルのアドレスなど、追加のコンテキストを取得するために使用できます。 オブジェクトに `Invocation` アクセスするには、カスタム関数の最後のパラメーターとして宣言 `invocation` する必要があります。

> [!NOTE]
> このパラメーターは `invocation` 、Excel のユーザーのカスタム関数引数として表示されません。

次の例では、パラメーターを使用 `invocation` して、カスタム関数を呼び出したセルのアドレスを返す方法を示します。 このサンプルでは、オブジェクトの [address](/javascript/api/custom-functions-runtime/customfunctions.invocation#custom-functions-runtime-customfunctions-invocation-address-member) プロパティを `Invocation` 使用します。 オブジェクトに `Invocation` アクセスするには、まず JSDoc でパラメーターとして宣言 `CustomFunctions.Invocation` します。 次に、JSDoc で宣言`@requiresAddress`してオブジェクトのプロパティに`Invocation`アクセス`address`します。 最後に、関数内でプロパティを取得して返 `address` します。

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
  const address = invocation.address;
  return address;
}
```

Excel では、オブジェクトの`Invocation`プロパティを`address`呼び出すカスタム関数は、関数を呼び出したセルの形式`SheetName!RelativeCellAddress`に従って絶対アドレスを返します。 たとえば、入力パラメーターがセル F6 の **Price** というシートに配置されている場合、返されるパラメーター アドレスの値は次のようになります `Prices!F6`。

このパラメーターを `invocation` 使用して、Excel に情報を送信することもできます。 詳細については、「 [ストリーミング関数を作成](custom-functions-web-reqs.md#make-a-streaming-function) する」を参照してください。

## <a name="detect-the-address-of-a-parameter"></a>パラメーターのアドレスを検出する

[呼び出しパラメーター](#invocation-parameter)と組み合わせて、[呼び出し](/javascript/api/custom-functions-runtime/customfunctions.invocation)オブジェクトを使用して、カスタム関数入力パラメーターのアドレスを取得できます。 呼び出されると、オブジェクトの `Invocation` [parameterAddresses](/javascript/api/custom-functions-runtime/customfunctions.invocation#custom-functions-runtime-customfunctions-invocation-parameteraddresses-member) プロパティを使用すると、関数はすべての入力パラメーターのアドレスを返すことができます。

これは、入力データ型が異なる可能性があるシナリオで役立ちます。 入力パラメーターのアドレスを使用して、入力値の数値形式を確認できます。 必要に応じて、数値の形式を入力前に調整できます。 入力パラメーターのアドレスは、入力値に後続の計算に関連する可能性のある関連プロパティがあるかどうかを検出するためにも使用できます。

>[!NOTE]
> [Office アドインの Yeoman ジェネレーター](../develop/yeoman-generator-overview.md)の代わりにパラメーター アドレスを返すために[手動で作成された JSON メタデータ](custom-functions-json.md)を使用している場合、オブジェクトにはプロパティが設定`true`され、`options`オブジェクトに`result`プロパティが設定`matrix`されている必要があります`requiresParameterAddresses``dimensionality`。

次のカスタム関数は、3 つの入力パラメーターを `parameterAddresses` 受け取り、各パラメーターのオブジェクトのプロパティを `Invocation` 取得し、アドレスを返します。

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
  const addresses = [
    [invocation.parameterAddresses[0]],
    [invocation.parameterAddresses[1]],
    [invocation.parameterAddresses[2]]
  ];
  return addresses;
}
```

プロパティを呼び出すカスタム関数が実行されると、関数を呼び `parameterAddresses` 出したセルの形式 `SheetName!RelativeCellAddress` に従ってパラメーター アドレスが返されます。 たとえば、入力パラメーターがセル D8 の **Costs** というシート上にある場合、返されるパラメーターのアドレス値は次のようになります `Costs!D8`。 カスタム関数に複数のパラメーターがあり、複数のパラメーター アドレスが返された場合、返されたアドレスは複数のセルにまたがり、関数を呼び出したセルから垂直方向に降下します。

## <a name="next-steps"></a>次の手順

[カスタム関数で揮発性値](custom-functions-volatile.md)を使用する方法について説明します。

## <a name="see-also"></a>関連項目

- [カスタム関数でデータを受信して​​処理する](custom-functions-web-reqs.md)
- [カスタム関数用の JSON メタデータの自動生成](custom-functions-json-autogeneration.md)
- [カスタム関数の JSON メタデータを手動で作成する](custom-functions-json.md)
- [Excel でカスタム関数を作成する](custom-functions-overview.md)
- [Excel カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)
