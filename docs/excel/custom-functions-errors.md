---
ms.date: 09/21/2020
description: '#NULL! のようなエラーを処理して返す カスタム関数から。'
title: カスタム関数を処理し、エラーを返します。
localization_priority: Normal
ms.openlocfilehash: 58c2ab432a4525f660e2d89735fd3add6e76fa7f
ms.sourcegitcommit: 4a03d8b3f676ee2d91114813cb81bce5da3c8d6b
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/22/2020
ms.locfileid: "48175529"
---
# <a name="handle-and-return-errors-from-your-custom-function"></a>カスタム関数を処理し、エラーを返します。

カスタム関数の実行中に何らかの問題が発生した場合は、ユーザーに通知するエラーを返します。 正の数だけなど、特定のパラメーター要件がある場合は、パラメーターをテストし、正しくない場合はエラーをスローします。 `try` - `catch` ブロックを使用して、カスタム関数の実行中に発生したエラーを検出することもできます。

## <a name="detect-and-throw-an-error"></a>エラーを検出してスローする

カスタム関数が動作するために zip コードパラメーターが正しい形式であることを確認する必要があるケースを見てみましょう。 次のカスタム関数は、正規表現を使用して郵便番号を確認します。 郵便番号の形式が正しい場合は、別の関数を使用して都市を検索し、その値を返します。 形式が有効でない場合、この関数は `#VALUE!` セルにエラーを返します。

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

Error オブジェクトは、セルにエラーを返すために使用されます[。](/javascript/api/custom-functions-runtime/customfunctions.error) オブジェクトを作成するときに、次の列挙値のいずれかを選択して、使用するエラーを指定し `ErrorCode` ます。


|ErrorCode enum value  |Excel のセル値  |意味  |
|---------------|---------|---------|
|`divisionByZero` | `#DIV/0`  | JavaScript ではゼロ除算が許可されるため、この状態を検出するには、慎重にエラーハンドラをに記述する必要があります。 |
|`invalidName`    | `#NAME?`  | 関数名に入力ミスがあります。 このエラーは、カスタム関数の入力エラーとしてサポートされますが、カスタム関数の出力エラーとしてはサポートされていないことに注意してください。 | 
|`invalidNumber`  | `#NUM!`   | 数式の数値に問題があります。 |
|`invalidReference` | `#REF!` | 関数が無効なセルを参照しています。 このエラーは、カスタム関数の入力エラーとしてサポートされますが、カスタム関数の出力エラーとしてはサポートされていないことに注意してください。|
|`invalidValue`   | `#VALUE!` | 数式の値の種類が正しくありません。 |
|`notAvailable`   | `#N/A`    | 関数またはサービスを使用できません。 |
|`nullReference`  | `#NULL!`  | 数式の範囲は交差しません。 |

次のコードサンプルは、無効な番号 (`#NUM!`) に対してエラーを作成して返す方法を示しています。

```typescript
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidNumber);
throw error;
```

およびエラーでは、 `#VALUE!` `#N/A` カスタムエラーメッセージもサポートされます。 エラーインジケーターメニューにカスタムエラーメッセージが表示されます。このメニューでは、エラーが発生した各セルのエラーフラグの上にカーソルがアクセスします。 次の例は、エラーが発生したカスタムエラーメッセージを返す方法を示して `#VALUE!` います。

```typescript
// You can only return a custom error message with the #VALUE! and #N/A errors.
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "The parameter can only contain lowercase characters.");
throw error;
```

## <a name="use-try-catch-blocks"></a>Use try-catch blocks

一般的には、 `try` - `catch` カスタム関数のブロックを使用して発生する可能性のあるエラーを検出します。 コードで例外を処理しない場合は、Excel に返されます。 既定で `#VALUE!` は、処理されないエラーまたは例外に対して Excel が返します。

次のコードサンプルでは、カスタム関数を使用して REST サービスの呼び出しを行ないます。 たとえば REST サービスがエラーを返したり、ネットワークがダウンした場合には、呼び出しが失敗することもあります。 このような場合、カスタム関数は、 `#N/A` web 呼び出しが失敗したことを示すためにを返します。


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
