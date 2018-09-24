---
ms.date: 09/20/2018
description: Excel のカスタム関数のベスト プラクティスと推奨パターンについて説明します。
title: カスタム関数のベスト プラクティス
ms.openlocfilehash: 1f2c0a80e62b65523fcc1673ba2ca4be444e6ce0
ms.sourcegitcommit: 470d8212b256275587e651abaa6f28beafebcab4
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/21/2018
ms.locfileid: "24068823"
---
# <a name="custom-functions-best-practices"></a>カスタム関数のベスト プラクティス

この記事は、Excel でカスタム関数を開発するためのベスト プラクティスについて説明します。

## <a name="error-handling"></a>エラー処理

カスタム関数を定義するアドインを作成する場合は、実行時エラーを含むエラー処理ロジックを含めてください。 カスタム関数のエラー処理は、「[一般的な Excel JavaScript API のエラー処理](excel-add-ins-error-handling.md) 」と同じです。 次のコード サンプルでは、`.catch` がコード内で発生するエラーを処理します。

```js
function getComment(x) {
    let url = "https://jsonplaceholder.typicode.com/comments/" + x; 
    return fetch(url)
        .then(function (data) {
            return data.json();
        })
        .then((json) => {
            return json.body;
        })
        .catch(function (error) {
            throw error;
        })
}
```

## <a name="error-logging"></a>エラー ログ

カスタム関数のエラーログは、次のような複数の方法で有効にすることができます。 

- アドインの XML マニフェスト ファイルをデバッグするために、[ 実行時ログを使用する](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in-manifest)。 

- カスタム関数内の `console.log` 文を使用し、コンソールにリアルタイムに出力を送信する。

> [!NOTE]
> 現時点では、実行時ログ機能は Office 2016 デスクトップでのみ利用可能です。

## <a name="debugging"></a>デバッグ

現時点で、Excel のカスタム関数をデバッグする最良の方法は、 [Excel Online](https://www.office.com/launch/excel) を使用し、使用するブラウザに対応する F12 デバッグ ツールを使用する方法です。 将来的には、カスタム関数用の他のデバッグ ツールも利用できる可能性があります。

## <a name="mapping-names"></a>名前のマッピング

デフォルトでは、JavaScript ファイル内のカスタム関数の名前は通常すべて大文字を使用して宣言し、エンド ユーザーに Excel で表示される関数の名前と正確に対応します。 ただし、`CustomFunctionsMappings` オブジェクトを使用して、JavaScript ファイルの 1 つ以上の関数名を、Excel でエンド ユーザーに関数名として表示する他の値にマップするように変更できます。 `CustomFunctionsMapping` を使用する必要はありませんが、大文字の関数名では問題が生じる uglifier、webpack、import 構文などを使用している場合に役立ちます。
  
次のコード サンプルは、JavaScript 関数名 `plusFortyTwo` を、Excel UI の `ADD42` 関数名にマップする単一のキーと値のペアを定義しています。 エンド ユーザーが Excel で `ADD42` 関数を選択すると、`plusFortyTwo` JavaScript 関数が実行されます。

```js
function plusFortyTwo(num) {
    return num + 42;  
}  
  
CustomFunctionsMappings = {
    "plusFortyTwo" : ADD42
}
```

次のコード サンプルは、2 つのキーと値のペアを定義しています。 最初のペアは、JavaScript 関数名 `plusFifty` を Excel UI の `ADD50` 関数名にマップし、2 番目のペアは、JavaScript 関数名 `plusOneHundred` を Excel UI の `ADD100` 関数名にマップします。 エンド ユーザーが Excel で `ADD50` 関数を選択すると、`plusFifty` JavaScript 関数が実行されます。 エンド ユーザーが Excel で `ADD100` 関数を選択すると、`plusOneHundred` JavaScript 関数が実行されます。

```js
function plusFifty(num) {
    return num + 50;  
} 

function plusOneHundred(num) {
    return num + 100;  
}  
  
CustomFunctionsMappings = {
    "plusFifty" : ADD50,  
    "plusOneHundred" : ADD100
}
 ```

 ## <a name="see-also"></a>関連項目

* [Excel でカスタム関数を作成する](custom-functions-overview.md)
* [カスタム関数のメタデータ](custom-functions-json.md)
* [Excel のカスタム関数のランタイム](custom-functions-runtime.md)