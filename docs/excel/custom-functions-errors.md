---
ms.date: 09/23/2020
description: '#NULL! のようなエラーを処理して返す カスタム関数から。'
title: カスタム関数のエラーを処理して返す
localization_priority: Normal
ms.openlocfilehash: 2822b3e93f7e5f16410e49d4414110e37172f3569b8f3c5d7d4dd98d5c5ecf6a
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57079675"
---
# <a name="handle-and-return-errors-from-your-custom-function"></a>カスタム関数のエラーを処理して返す

カスタム関数の実行中に問題が発生した場合は、エラーを返してユーザーに通知します。 正の数値のみなど、特定のパラメーター要件がある場合は、パラメーターをテストし、正しい値でない場合はエラーをスローします。 `try` - `catch` ブロックを使用して、カスタム関数の実行中に発生したエラーを検出することもできます。

## <a name="detect-and-throw-an-error"></a>エラーを検出してスローする

カスタム関数が正しい形式で動作していることを確認する必要がある場合について説明します。 次のカスタム関数は、正規表現を使用して郵便番号を確認します。 郵便番号の形式が正しい場合は、別の関数を使用して都市を参照し、値を返します。 書式が無効な場合、関数はセルに `#VALUE!` エラーを返します。

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

[CustomFunctions.Error](/javascript/api/custom-functions-runtime/customfunctions.error)オブジェクトを使用して、セルにエラーを返します。 オブジェクトを作成する場合は、次のいずれかの列挙値を選択して、使用するエラーを `ErrorCode` 指定します。


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

また `#VALUE!` 、エラー `#N/A` はカスタム エラー メッセージもサポートします。 カスタム エラー メッセージはエラー インジケーター メニューに表示され、エラーのある各セルのエラー フラグにカーソルを合わせるとアクセスされます。 次の例は、エラーを含むカスタム エラー メッセージを返す方法を示 `#VALUE!` しています。

```typescript
// You can only return a custom error message with the #VALUE! and #N/A errors.
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "The parameter can only contain lowercase characters.");
throw error;
```

## <a name="use-try-catch-blocks"></a>Use try-catch blocks

一般に、カスタム `try` - `catch` 関数でブロックを使用して、発生する可能性のあるエラーをキャッチします。 コードで例外を処理しない場合は、Excel に返されます。 既定では、Excelエラーまたは例外 `#VALUE!` が返されます。

次のコードサンプルでは、カスタム関数を使用して REST サービスの呼び出しを行ないます。 たとえば REST サービスがエラーを返したり、ネットワークがダウンした場合には、呼び出しが失敗することもあります。 この場合、カスタム関数は Web 呼び出しが失敗 `#N/A` したと示すために返されます。


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
* [カスタム関数の要件](custom-functions-requirement-sets.md)
* [Excel でカスタム関数を作成する](custom-functions-overview.md)
