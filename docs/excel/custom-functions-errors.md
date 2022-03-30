---
title: カスタム関数のエラーを処理して返す
description: '#NULL! のようなエラーを処理して返す カスタム関数から。'
ms.date: 08/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: c8f2667f47c1c983b135f38ce2c67ad1f31502c9
ms.sourcegitcommit: 287a58de82a09deeef794c2aa4f32280efbbe54a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/28/2022
ms.locfileid: "64496293"
---
# <a name="handle-and-return-errors-from-your-custom-function"></a>カスタム関数のエラーを処理して返す

カスタム関数の実行中に問題が発生した場合は、エラーを返してユーザーに通知します。 正の数値のみなど、特定のパラメーター要件がある場合は、パラメーターをテストし、正しい値でない場合はエラーをスローします。 ブロックを使用して、 [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) カスタム関数の実行中に発生するエラーをキャッチできます。

## <a name="detect-and-throw-an-error"></a>エラーを検出してスローする

カスタム関数が正しい形式で動作していることを確認する必要がある場合について説明します。 次のカスタム関数は、正規表現を使用して郵便番号を確認します。 郵便番号の形式が正しい場合は、別の関数を使用して都市を参照し、値を返します。 書式が無効な場合、関数はセルにエラー `#VALUE!` を返します。

```typescript
/**
* Gets a city name for the given U.S. zip code.
* @customfunction
* @param {string} zipCode
* @returns The city of the zip code.
*/
function getCity(zipCode: string): string {
  let isValidZip = /(^\d{5}$)|(^\d{5}-\d{4}$)/.test(zipCode);
  if (isValidZip) return cityLookup(zipCode);
  let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "Please provide a valid U.S. zip code.");
  throw error;
}
```

## <a name="the-customfunctionserror-object"></a>The CustomFunctions.Error object

[CustomFunctions.Error](/javascript/api/custom-functions-runtime/customfunctions.error) オブジェクトを使用して、セルにエラーを返します。 オブジェクトを作成する場合は、次のいずれかの列挙値を選択して、使用するエラーを `ErrorCode` 指定します。

|ErrorCode enum value  |Excel のセル値  |説明  |
|---------------|---------|---------|
|`divisionByZero` | `#DIV/0`  | 関数は 0 で除算を試行しています。 |
|`invalidName`    | `#NAME?`  | 関数名に入力ミスがあります。 このエラーはカスタム関数入力エラーとしてサポートされますが、カスタム関数出力エラーとしてサポートされません。 |
|`invalidNumber`  | `#NUM!`   | 数式の数値に問題があります。 |
|`invalidReference` | `#REF!` | この関数は、無効なセルを参照します。 このエラーはカスタム関数入力エラーとしてサポートされますが、カスタム関数出力エラーとしてサポートされません。|
|`invalidValue`   | `#VALUE!` | 数式の値が正しい型です。 |
|`notAvailable`   | `#N/A`    | 関数またはサービスは使用できません。 |
|`nullReference`  | `#NULL!`  | 数式内の範囲は交差しません。 |

次のコードサンプルは、無効な番号 (`#NUM!`) に対してエラーを作成して返す方法を示しています。

```typescript
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidNumber);
throw error;
```

また、 `#VALUE!` エラー `#N/A` はカスタム エラー メッセージもサポートします。 カスタム エラー メッセージはエラー インジケーター メニューに表示され、エラーのある各セルのエラー フラグにカーソルを合わせるとアクセスされます。 次の例は、エラーを含むカスタム エラー メッセージを返す方法を示 `#VALUE!` しています。

```typescript
// You can only return a custom error message with the #VALUE! and #N/A errors.
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "The parameter can only contain lowercase characters.");
throw error;
```

### <a name="handle-errors-when-working-with-dynamic-arrays"></a>動的配列を操作するときにエラーを処理する

1 つのエラーを返すだけでなく、カスタム関数は、エラーを含む動的配列を出力できます。 たとえば、カスタム関数は配列を出力できます `[1],[#NUM!],[3]`。 次のコード `#NUM!` サンプルは、3 つのパラメーターをカスタム関数に入力し、入力パラメーターの 1 つをエラーに置き換え、2 次元配列を各入力パラメーターの処理結果で返す方法を示しています。

```js
/**
* Returns the #NUM! error as part of a 2-dimensional array.
* @customfunction
* @param {number} first First parameter.
* @param {number} second Second parameter.
* @param {number} third Third parameter.
* @returns {number[][]} Three results, as a 2-dimensional array.
*/
function returnInvalidNumberError(first, second, third) {
  // Use the `CustomFunctions.Error` object to retrieve an invalid number error.
  var error = new CustomFunctions.Error(
    CustomFunctions.ErrorCode.invalidNumber, // Corresponds to the #NUM! error in the Excel UI.
  );

  // Enter logic that processes the first, second, and third input parameters.
  // Imagine that the second calculation results in an invalid number error. 
  var firstResult = first;
  var secondResult =  error;
  var thirdResult = third;

  // Return the results of the first and third parameter calculations and a #NUM! error in place of the second result. 
  return [[firstResult], [secondResult], [thirdResult]];
}
```

### <a name="errors-as-custom-function-inputs"></a>カスタム関数入力としてのエラー

カスタム関数は、入力範囲にエラーが含まれている場合でも評価できます。 たとえば、**A6:A7** にエラーが含まれている場合でも、カスタム関数は範囲 **A2:A7** を入力として受け取る場合があります。

エラーを含む入力を処理するには、カスタム関数に JSON メタデータ プロパティが設定されている `allowErrorForDataTypeAny` 必要があります `true`。 詳細については [、「カスタム関数の JSON メタデータを手動で作成する](custom-functions-json.md#metadata-reference) 」を参照してください。

> [!IMPORTANT]
> この `allowErrorForDataTypeAny` プロパティは、手動で作成された [JSON メタデータでのみ使用できます](custom-functions-json.md)。 このプロパティは、自動生成された JSON メタデータ プロセスでは機能しません。

## <a name="use-trycatch-blocks"></a>ブロックを使用 `try...catch` する

一般に、カスタム関数 [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) でブロックを使用して、発生する可能性のあるエラーをキャッチします。 コード内の例外を処理しない場合、例外はコードに返Excel。 既定では、Excelエラー`#VALUE!`または例外が返されます。

次のコードサンプルでは、カスタム関数を使用して REST サービスの呼び出しを行ないます。 たとえば REST サービスがエラーを返したり、ネットワークがダウンした場合には、呼び出しが失敗することもあります。 この場合、カスタム関数は Web 呼び出 `#N/A` しが失敗したと示すために返されます。

```typescript
/**
 * Gets a comment from the hypothetical contoso.com/comments API.
 * @customfunction
 * @param {number} commentID ID of a comment.
 */
function getComment(commentID) {
  let url = "https://www.contoso.com/comments/" + commentID;
  return fetch(url)
    .then(function (data) {
      return data.json();
    })
    .then(function (json) {
      return json.body;
    })
    .catch(function (error) {
      throw new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable);
    })
}
```

## <a name="next-steps"></a>次の手順

[自分のカスタム関数で問題をトラブルシューティングを行う](custom-functions-troubleshooting.md)方法についての詳細を確認する。

## <a name="see-also"></a>関連項目

* [カスタム関数のデバッグ](custom-functions-debugging.md)
* [カスタム関数の要件セット](/javascript/api/requirement-sets/excel/custom-functions-requirement-sets)
* [Excel でカスタム関数を作成する](custom-functions-overview.md)
