---
ms.date: 06/18/2019
description: Excel 範囲、省略可能なパラメーター、呼び出しコンテキストなど、カスタム関数内でさまざまなパラメーターを使用する方法について説明します。
title: Excel カスタム関数のオプション
localization_priority: Normal
ms.openlocfilehash: dca85df87f0153c03b2ddd027748e16d3ec79924
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128340"
---
# <a name="custom-functions-parameter-options"></a>カスタム関数のパラメータオプション

カスタム関数は、パラメーターにさまざまなオプションを使用して構成できます。
- [オプションのパラメーター](#custom-functions-optional-parameters)
- [範囲パラメーター](#range-parameters)
- [呼び出しコンテキストパラメーター](#invocation-parameter)

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="custom-functions-optional-parameters"></a>カスタム関数の省略可能なパラメーター

通常のパラメーターは必須ですが、省略可能なパラメーターは必須ではありません。 ユーザーが Excel で関数を呼び出すと、角かっこで囲まれた省略可能なパラメーターが表示されます。 次の例では、add 関数で3番目の番号を追加することもできます。 この関数は Excel `=CONTOSO.ADD(first, second, [third])`のように表示されます。

```js
/**
 * Add two numbers
 * @customfunction 
 * @param {number} first First number.
 * @param {number} second Second number.
 * @param {number} [third] Third number to add. If omitted, third = 0.
 * @returns {number} The sum of the numbers.
 */
function add(first, second, third) {
  if (third !== undefined) {
    return first + second + third;
  }
  return first + second;
}
CustomFunctions.associate("ADD", add);
```

関数の定義時に 1 つ以上の省略可能なパラメーターを含める場合は、省略可能なパラメーターが未定義のときの処理を指定しておく必要があります。 次の例の `zipCode` と `dayOfWeek` は、どちらも `getWeatherReport` 関数の省略可能なパラメーターです。 `zipCode`パラメーターが定義されていない場合、既定値`98052`はに設定されます。 `dayOfWeek` パラメーターが未定義の場合は、Wednesday が設定されます。

```js
/**
 * Gets a weather report for a specified zipCode and dayOfWeek
 * @customfunction
 * @param {number} zipCode Zip code. If omitted, zipCode = 98052.
 * @param {string} dayOfWeek Day of the week. If omitted, dayOfWeek = Wednesday.
 * @returns {string} Weather report for the day of the week in that zip code.
 */
function getWeatherReport(zipCode, dayOfWeek)
{
  if (zipCode === undefined) {
      zipCode = "98052";
  }

  if (dayOfWeek === undefined) {
    dayOfWeek = "Wednesday";
  }

  // Get weather report for specified zipCode and dayOfWeek.
  // ...
}
```

## <a name="range-parameters"></a>範囲パラメーター

カスタム関数は、入力パラメーターとして範囲のセルデータを受け入れることができます。 関数は、データの範囲を返すこともできます。 Excel は、セルデータの範囲を2次元配列として渡します。

例えば、関数が Excel に保存されている数値の範囲から 2 番目に大きい値を返すとします。 次の関数は、`Excel.CustomFunctionDimensionality.matrix` 型の `values` パラメーターを受け入れます。 この関数の JSON メタデータでは、パラメーターの`type`プロパティがに`matrix`設定されていることに注意してください。

```js
/**
 * Returns the second highest value in a matrixed range of values.
 * @customfunction
 * @param {number[][]} values Multiple ranges of values.  
 */
function secondHighest(values){
  let highest = values[0][0], secondHighest = values[0][0];
  for(var i = 0; i < values.length; i++){
    for(var j = 0; j < values[i].length; j++){
      if(values[i][j] >= highest){
        secondHighest = highest;
        highest = values[i][j];
      }
      else if(values[i][j] >= secondHighest){
        secondHighest = values[i][j];
      }
    }
  }
  return secondHighest;
}
CustomFunctions.associate("SECONDHIGHEST", secondHighest);
```

## <a name="invocation-parameter"></a>呼び出しパラメーター

すべてのカスタム関数には、 `invocation`最後の引数として引数が自動的に渡されます。 この引数は、呼び出し元のセルのアドレスなど、追加のコンテキストを取得するために使用できます。 また、関数[をキャンセル](custom-functions-web-reqs.md#make-a-streaming-function)する関数ハンドラーなど、Excel に情報を送信するために使用することもできます。 パラメーターを宣言しない場合でも、カスタム関数にはこのパラメーターがあります。 この引数は、Excel のユーザーには表示されません。 カスタム関数でを使用`invocation`する場合は、最後のパラメーターとして宣言します。

次のコードサンプルでは、 `invocation`コンテキストが参照に対して明示的に指定されています。

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
CustomFunctions.associate("ADD", add);
```

このパラメーターを使用すると、呼び出し元のセルのコンテキストを取得できます。これは、[カスタム関数を呼び出すセルのアドレスを検索](#addressing-cells-context-parameter)するなどの一部のシナリオで役立ちます。

### <a name="addressing-cells-context-parameter"></a>アドレス指定セルのコンテキストパラメーター

場合によっては、カスタム関数を呼び出したセルのアドレスを取得する必要があります。 これは、次のシナリオで役立ちます。

- 範囲の書式設定: セルのアドレスをキーとして使用し、データを保存します[。](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data) Excel で [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) を使用して`OfficeRuntime.storage` からキーを読み込みます。
- キャッシュされた値を表示させる: 関数がオフラインで使用される場合、`onCalculated` を使用して `OfficeRuntime.storage` に格納されているキャッシュされた値を表示します。
- 調整: セル アドレスを使用して元のセルを検出し、処理が発生している場所での調整を行えます。

関数内のアドレス指定セルのコンテキストを要求するには、次の例のように、関数を使用してセルのアドレスを検索する必要があります。 セルのアドレスに関する情報は、関数のコメント`@requiresAddress`にタグ付けされている場合にのみ公開されます。

```js
/**
 * Function that gets the address of a cell.
 * @customfunction
 * @param {CustomFunctions.Invocation} invocation Uses the invocation parameter present in each cell.
 * @requiresAddress
 * @returns {string} Returns address of cell.
 */

function getAddress(invocation) {
  return invocation.address;
}
CustomFunctions.associate("GETADDRESS", getAddress);
```

既定では、`getAddress` 関数が返す値は次の形式に従います: `SheetName!CellNumber`。 たとえば、ある関数が Expenses という名前のシートのセル B2 から呼び出される場合の戻り値は `Expenses!B2` になります。

## <a name="next-steps"></a>次のステップ
カスタム関数の[状態を保存](custom-functions-save-state.md)する方法、または[カスタム関数で揮発性の値](custom-functions-volatile.md)を使用する方法について説明します。

## <a name="see-also"></a>関連項目

* [カスタム関数でデータを受信して​​処理する](custom-functions-web-reqs.md)
* [カスタム関数のベスト プラクティス](custom-functions-best-practices.md)
* [カスタム関数のメタデータ](custom-functions-json.md)
* [カスタム関数用の JSON メタデータの自動生成](custom-functions-json-autogeneration.md)
* [Excel でカスタム関数を作成する](custom-functions-overview.md)
* [Excel カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)
