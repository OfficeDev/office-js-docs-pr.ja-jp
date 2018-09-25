---
ms.date: 09/20/2018
description: Excel のカスタム関数のベスト プラクティスと推奨パターンについて説明します。
title: カスタム関数のベスト プラクティス
ms.openlocfilehash: 3934910c397aea348c4fe2d7f95f1dc20ebeb4d3
ms.sourcegitcommit: 8ce9a8d7f41d96879c39cc5527a3007dff25bee8
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/24/2018
ms.locfileid: "24985789"
---
# <a name="custom-functions-best-practices"></a>カスタム関数のベスト プラクティス

この記事は、Excel でカスタム関数を開発するためのベスト プラクティスについて説明します。

## <a name="error-handling"></a>エラー処理

カスタム関数を定義するアドインを作成する場合は、実行時エラーに対処するエラー処理ロジックを含めてください。 カスタム関数のエラー処理は、[一般的な Excel JavaScript API のエラー処理](excel-add-ins-error-handling.md) と同じです。 次のコード サンプルでは、`.catch` がコード内で発生するエラーを処理します。

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

現時点で Excel カスタム関数をデバッグするための最良の方法は、Excel Online 内でアドインを最初に[サイドロード](../testing/sideload-office-add-ins-for-testing.md)することです。  [お使いのブラウザーにネイティブの F12 デバッグ ツール](../testing/debug-add-ins-in-office-online.md)を使用して、カスタム関数をデバッグできます。

アドインの登録に失敗した場合は、アドイン アプリケーションをホストしている Web サーバーに、 [SSL 証明書が正しく構成されていることを確認してください](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) 。

## <a name="mapping-names"></a>名前のマッピング

デフォルトでは、JavaScript ファイル内のカスタム関数の名前は通常すべて大文字を使用して宣言し、エンド ユーザーに Excel で表示される関数の名前と正確に対応します。 ただし、`CustomFunctionsMappings` オブジェクトを使用して、JavaScript ファイルの 1 つ以上の関数名を、Excel でエンド ユーザーに関数名として表示する他の値にマップするように変更できます。 Uglifier、webpack、または大文字の関数名が困難なすべてのインポートの構文を使用している場合に便利です。 `CustomFunctionsMappings` プロジェクトが JavaScript を使用するのは恐らくオプションですが、プロジェクトが TypeScript を使用している場合は、使用する必要があります。  
  
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
