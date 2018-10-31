---
ms.date: 10/03/2018
description: Excel のカスタム関数のベスト プラクティスと推奨パターンについて説明します。
title: カスタム関数のベスト プラクティス
ms.openlocfilehash: 218e62cd074ccf3f3708bba90c938f7ddef059cb
ms.sourcegitcommit: f47654582acbe9f618bec49fb97e1d30f8701b62
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/17/2018
ms.locfileid: "25579822"
---
# <a name="custom-functions-best-practices-preview"></a>カスタム関数のベスト プラクティス (プレビュー)

この記事は、Excel でカスタム関数を開発するためのベスト プラクティスについて説明します。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="error-handling"></a>エラー処理

カスタム関数を定義するアドインをビルドする場合は、実行時エラーを考慮するためのロジックを含めるようにしてください。カスタム関数のエラー処理は、 [大規模な Excel の JavaScript API のエラーの処理](excel-add-ins-error-handling.md)と同じです。次のコード サンプルでは、 `.catch` 、コード内で以前に発生したエラーを処理します。

```js
function getComment(x) {
  let url = "https://www.contoso.com/comments/" + x;
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

## <a name="debugging"></a>デバッグ

現時点で Excel カスタム関数をデバッグするための最良の方法は、[ Excel Online ](../testing/sideload-office-add-ins-for-testing.md) 内でアドインを最初に** サイドロード** することです。次の手法と組み合わせて [お使いのブラウザでネイティブのF12 デバッグ ツール](../testing/debug-add-ins-in-office-online.md) を使用して、カスタム関数をデバッグできます。

- カスタム関数のコード内で `console.log` 文を使用して、コンソールにリアルタイムに出力を送信します。

- カスタム関数コード内の `debugger;` 文を使用して、f12  ウィンドウが開いているときに実行が一時停止するブレークポイントを指定します、例えばf12  ウィンドウが開いているとき以下の関数が動作している場合には、`debugger;`文上で実行が停止し、 関数が返される前に、パラメーター値を手動で検査することができます。 `debugger;` 文は、F12 ウィンドウが開いていない場合、Excel Onlineには影響しません。。現在、 `debugger;` 文はwindows 版 Excel には効果がありません。

    ```js
    function add(first, second){
      debugger;
      return first + second;
    }
    ```

アドインが登録に失敗した場合は、アドイン アプリケーションをホストしている Web サーバーに、 [SSL 証明書が正しく構成されていることを確認してください](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) 。

WindowsのデスクトップでOfficeのアドインをテストする場合、いくつかのインストールと実行時の条件と同様に、アドインの XML マニフェスト ファイルの問題をデバッグする [実行時のログ](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in) を有効にできます。

## <a name="mapping-function-names-to-json-metadata"></a>関数名を JSON のメタデータにマップする

[カスタム関数の概要](custom-functions-overview.md) 資料で説明したように、カスタム関数プロジェクトには、カスタム関数を登録し、エンド ユーザーが利用できるように Excel が必要とする情報を提供する JSON のメタデータ ファイルを含める必要があります。さらに、カスタム関数を定義する JavaScript ファイル内で、JSON のメタデータ ファイルにあるどの関数オブジェクトが JavaScript ファイル内の各ユーザー定義関数に対応するかを指定する情報を提供する必要があります。

たとえば、次のコード サンプルは、カスタム関数 `add` を定義し、`id` プロパティの値が **ADD**である JSON のメタデータ ファイル内のオブジェクトに関数 `add` が対応するよう指定します。

```js
function add(first, second){
  return first + second;
}

CustomFunctionMappings.ADD = add;
```

JavaScript ファイルでカスタム関数を作成し、JSON のメタデータ ファイルに対応する情報を指定するときは、次のベスト プラクティスに留意してください。

* JavaScript fileでは関数名をcamelCaseで記述します。たとえば、関数名 `addTenToInput` はcamelCaseで記述されています：名前の最初の単語は小文字で開始し、後続の各単語は大文字で開始します。

* JSON メタデータ ファイル内で、各`name` プロパティの値に大文字を指定します。 `name` プロパティは、Excel でエンド ユーザーに表示される関数の名前を定義します。各カスタム関数の名前の大文字を使用することで、すべての組み込み関数の名前は大文字である Excel で、一貫性のあるエクスペリエンスをエンド ユーザーに提供します。

* JSON メタデータ ファイル内で、各 `id` プロパティの値に大文字を指定します。このようにすると、JavaScript コード内の `CustomFunctionMappings` 文のどの部分が、JSON のメタデータの `id`プロパティに対応するか明らかにしますのステートメントの対応するか明らかになります (推奨したように、関数名はcamelCaseを使用します) 。

* JSON のメタデータ ファイルに確実にそれぞれの値 `id` プロパティには、英数字とピリオドのみが含まれています。 

* JSON のメタデータ ファイルで、各 `id` プロパティの値が、ファイルのスコープ内で一意であることを確認します。すなわち、メタデータ ファイル内の 2 つの関数オブジェクトは同じ `id` 値であってはいけません。さらに、2 つの 大文字と小文字だけが異なるメタデータ ファイル内の`id` 値を指定しないでください。たとえば、 **add**の値`id`  の関数オブジェクトを、**ADD**の値`id`  の別の関数オブジェクトと定義しないでください。

* 対応する JavaScript 関数の名前にマップされた後では、JSON のメタデータ ファイル内の `id` プロパティの値を変更しないでください。JSON のメタデータ ファイル内の`name` プロパティを更新することによって Excel でエンド ユーザーに表示される関数の名前を変更することができます。しかし、確立された後は、  `id` プロパティの値を決して変更しないでください。

* JavaScript ファイルで同じ場所にすべてのカスタム関数のマッピングを指定します。例えば次のコード サンプル は、2 つのカスタム関数を定義し、両方の関数のマッピング情報を指定します。

    ```js
    function add(first, second){
      return first + second;
    }

    function increment(incrementBy, callback) {
      var result = 0;
      var timer = setInterval(function() {
        result += incrementBy;
        callback.setResult(result);
      }, 1000);

      callback.onCanceled = function() {
        clearInterval(timer);
      };
    }

    // map `id` values in the JSON metadata file to JavaScript function names
    CustomFunctionMappings.ADD = add;
    CustomFunctionMappings.INCREMENT = increment;
    ```

    次のサンプルは、JavaScript コード サンプルで定義された関数に対応する JSON メタデータを示します。

    ```json
    {
      "$schema": "https://developer.microsoft.com/en-us/json-schemas/office-js/custom-functions.schema.json",
      "functions": [
        {
          "id": "ADD",
          "name": "ADD",
          ...
        },
        {
          "id": "INCREMENT",
          "name": "INCREMENT",
          ...
        }
      ]
    }
    ```

## <a name="additional-considerations"></a>その他の考慮事項

複数のプラットフォーム（Officeアドインの主要テナントの一つ）で動作するアドインを作成する際は、カスタム関数でドキュメント オブジェクト モデル (DOM) にアクセスしたり、jQueryのようなDOMに依存するライブラリーを使用してはいけません。 カスタム関数が [JavaScript のランタイム](custom-functions-runtime.md)を使用するExcel for Windows のウィンドウでは、ユーザー定義関数はDOM にアクセスできません。

## <a name="see-also"></a>関連項目

* [Excel でカスタム関数を作成する](custom-functions-overview.md)
* [カスタム関数のメタデータ](custom-functions-json.md)
* [Excel カスタム関数のランタイム](custom-functions-runtime.md)
* [Excel カスタム関数のチュートリアル](excel-tutorial-custom-functions.md)
