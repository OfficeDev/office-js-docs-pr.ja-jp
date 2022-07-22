---
title: カスタム関数のエラーを処理して返す
description: '#NULL! のようなエラーを処理して返す カスタム関数から取得します。'
ms.date: 08/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: c93c13aac1457e776ba8441565c11a23074a8d97
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958567"
---
# <a name="handle-and-return-errors-from-your-custom-function"></a>カスタム関数のエラーを処理して返す

カスタム関数の実行中に問題が発生した場合は、エラーを返してユーザーに通知します。 正の数値のみなど、特定のパラメーター要件がある場合は、パラメーターをテストし、正しくない場合はエラーをスローします。 また、ブロックを使用して、 [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) カスタム関数の実行中に発生するエラーをキャッチすることもできます。

## <a name="detect-and-throw-an-error"></a>エラーを検出してスローする

カスタム関数が機能するように郵便番号パラメーターが正しい形式であることを確認する必要があるケースを見てみましょう。 次のカスタム関数は、正規表現を使用して郵便番号を確認します。 郵便番号の形式が正しい場合は、別の関数を使用して市区町村を検索し、値を返します。 形式が無効な場合、関数はセルにエラーを返 `#VALUE!` します。

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

[CustomFunctions.Error](/javascript/api/custom-functions-runtime/customfunctions.error) オブジェクトは、セルにエラーを返すために使用されます。 オブジェクトを作成するときは、次 `ErrorCode` のいずれかの列挙値を選択して、使用するエラーを指定します。

|ErrorCode enum value  |Excel のセル値  |説明  |
|---------------|---------|---------|
|`divisionByZero` | `#DIV/0`  | 関数は 0 で除算しようとしています。 |
|`invalidName`    | `#NAME?`  | 関数名に入力ミスがあります。 このエラーはカスタム関数入力エラーとしてサポートされていますが、カスタム関数の出力エラーとしてサポートされないことに注意してください。 |
|`invalidNumber`  | `#NUM!`   | 数式の数値に問題があります。 |
|`invalidReference` | `#REF!` | 関数は無効なセルを参照します。 このエラーはカスタム関数入力エラーとしてサポートされていますが、カスタム関数の出力エラーとしてサポートされないことに注意してください。|
|`invalidValue`   | `#VALUE!` | 数式の値が間違った型です。 |
|`notAvailable`   | `#N/A`    | 関数またはサービスは使用できません。 |
|`nullReference`  | `#NULL!`  | 数式の範囲が交差しません。 |

次のコードサンプルは、無効な番号 (`#NUM!`) に対してエラーを作成して返す方法を示しています。

```typescript
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidNumber);
throw error;
```

また、エラーと`#N/A`エラーは`#VALUE!`カスタム エラー メッセージもサポートします。 カスタム エラー メッセージは、エラー インジケーター メニューに表示されます。これは、エラーを含む各セルのエラー フラグの上にマウス ポインターを置いてアクセスします。 次の例は、エラーを含むカスタム エラー メッセージを返す方法を `#VALUE!` 示しています。

```typescript
// You can only return a custom error message with the #VALUE! and #N/A errors.
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "The parameter can only contain lowercase characters.");
throw error;
```

### <a name="handle-errors-when-working-with-dynamic-arrays"></a>動的配列を操作するときにエラーを処理する

カスタム関数は、単一のエラーを返すだけでなく、エラーを含む動的配列を出力することもできます。 たとえば、カスタム関数は配列 `[1],[#NUM!],[3]`を出力できます。 次のコード サンプルは、3 つのパラメーターをカスタム関数に入力し、入力パラメーターの 1 つをエラーに `#NUM!` 置き換え、各入力パラメーターを処理した結果で 2 次元配列を返す方法を示しています。

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
  const error = new CustomFunctions.Error(
    CustomFunctions.ErrorCode.invalidNumber, // Corresponds to the #NUM! error in the Excel UI.
  );

  // Enter logic that processes the first, second, and third input parameters.
  // Imagine that the second calculation results in an invalid number error. 
  const firstResult = first;
  const secondResult =  error;
  const thirdResult = third;

  // Return the results of the first and third parameter calculations and a #NUM! error in place of the second result. 
  return [[firstResult], [secondResult], [thirdResult]];
}
```

### <a name="errors-as-custom-function-inputs"></a>カスタム関数入力としてのエラー

カスタム関数は、入力範囲にエラーが含まれている場合でも評価できます。 たとえば、カスタム関数は **、A6:A7** にエラーが含まれている場合でも、 **A2:A7** の範囲を入力として受け取ることができます。

エラーを含む入力を処理するには、カスタム関数に JSON メタデータ プロパティ `allowErrorForDataTypeAny` を設定する `true`必要があります。 詳細については、「 [カスタム関数の JSON メタデータを手動で作成](custom-functions-json.md#metadata-reference) する」を参照してください。

> [!IMPORTANT]
> このプロパティは `allowErrorForDataTypeAny` 、 [手動で作成された JSON メタデータ](custom-functions-json.md)でのみ使用できます。 このプロパティは、自動生成された JSON メタデータ プロセスでは機能しません。

## <a name="use-trycatch-blocks"></a>ブロックを使用する`try...catch`

一般に、カスタム関数でブロックを使用 [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) して、発生する可能性のあるエラーをキャッチします。 コードで例外を処理しない場合は、Excel に返されます。 既定では、未処理の `#VALUE!` エラーまたは例外が Excel から返されます。

次のコードサンプルでは、カスタム関数を使用して REST サービスの呼び出しを行ないます。 たとえば REST サービスがエラーを返したり、ネットワークがダウンした場合には、呼び出しが失敗することもあります。 この場合、カスタム関数は Web 呼び出しが失敗したことを示すために戻ります `#N/A` 。

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
