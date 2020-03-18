---
ms.date: 03/11/2020
description: '#NULL! のようなエラーを処理して返す カスタム関数で'
title: カスタム関数でエラーを処理して返す (プレビュー)
localization_priority: Normal
ms.openlocfilehash: 10bb7ca6ff612ef38b26b88fed5ce9ce81ed7edb
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717048"
---
# <a name="handle-and-return-errors-from-your-custom-function-preview"></a>カスタム関数でエラーを処理して返す (プレビュー)

> [!NOTE]
> この記事で説明する機能は現在プレビュー中であり、変更される可能性があります。 これらを運用環境で使用することは現在サポートされていません。 プレビュー機能を試すには、 [Office Insider](https://insider.office.com/join)プログラムに参加する必要があります。  プレビュー機能を試す良い方法は、Office 365 サブスクリプションを使用することです。 Office 365 サブスクリプションをまだお持ちでない場合は、[Office 365 開発者プログラム](https://developer.microsoft.com/office/dev-program)に参加することで 90 日間の更新可能な無料の Office 365 サブスクリプションを入手できます。

カスタム関数の実行中に問題が発生した場合、エラーを返してユーザーに通知する必要があります。 正数のみなど、特定のパラメーター要件がある場合は、パラメーターをテストし、正しくない場合はエラーをスローする必要があります。 `try` - `catch` ブロックを使用して、カスタム関数の実行中に発生したエラーを検出することもできます。

## <a name="detect-and-throw-an-error"></a>エラーを検出してスローする

カスタム関数が動作するために zip コードパラメーターが正しい形式であることを確認する必要があるケースを見てみましょう。 次のカスタム関数は、正規表現を使用して郵便番号を確認します。 正しい場合は、(別の関数で) 都市を検索し、その値を返します。 正しくない場合は、セルに `#VALUE!` エラーを返します。

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

`CustomFunctions.Error` オブジェクトは、セルにエラーを返すために使用されます。 オブジェクトを作成するときに、次の `ErrorCode` 列挙値のいずれかを使用して、使用するエラーを指定します。


|ErrorCode enum value  |Excel のセル値  |意味  |
|---------------|---------|---------|
|`invalidValue`   | `#VALUE!` | 数式で使用されている値の型が間違っている。 |
|`notAvailable`   | `#N/A`    | 機能またはサービスが利用できない。 |
|`divisionByZero` | `#DIV/0`  | JavaScript ではゼロ除算が許可されるため、この状態を検出するには、慎重にエラーハンドラをに記述する必要があります。 |
|`invalidNumber`  | `#NUM!`   | 数式で使用されている番号に問題がある。 |
|`nullReference`  | `#NULL!`  | 数式の範囲が交わることはありません。 |

次のコードサンプルは、無効な番号 (`#NUM!`) に対してエラーを作成して返す方法を示しています。

```typescript
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidNumber);
throw error;
```

`#VALUE!` エラーを返す場合、ユーザーがセルにカーソルを合わせたときにポップアップに表示されるカスタムメッセージを含めることもできます。 次の例は、カスタムのエラーメッセージを返す方法を示しています。

```typescript
// You can only return a custom error message with the #VALUE! error
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "The parameter can only contain lowercase characters.");
throw error;
```

## <a name="use-try-catch-blocks"></a>Use try-catch blocks

通常、発生する可能性があるエラーをキャッチするには、カスタム関数で `try` - `catch` ブロックを使用する必要があります。 コードで例外を処理しない場合は、Excel に返されます。 既定では、Excel は未処理の例外に対して `#VALUE!` を返します。

次のコードサンプルでは、カスタム関数を使用して REST サービスの呼び出しを行ないます。 たとえば REST サービスがエラーを返したり、ネットワークがダウンした場合には、呼び出しが失敗することもあります。 この場合、カスタム関数は Web 呼び出しが失敗したことを示す `#N/A` を返します。


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
