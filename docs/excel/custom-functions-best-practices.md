---
ms.date: 09/27/2018
description: Excel のカスタム関数のベスト プラクティスと推奨パターンについて説明します。
title: カスタム関数のベスト プラクティス
ms.openlocfilehash: d157464a3a8bf453cd0970281f1a4fdd27df5d25
ms.sourcegitcommit: 1852ae367de53deb91d03ca55d16eb69709340d3
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/29/2018
ms.locfileid: "25348788"
---
# <a name="custom-functions-best-practices-preview"></a>カスタム関数のベスト プラクティス (プレビュー)

この記事は、Excel でカスタム関数を開発するためのベスト プラクティスについて説明します。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="error-handling"></a>エラー処理

カスタム関数を定義するアドインを作成する場合は、実行時エラーに対処するエラー処理ロジックを含めてください。 カスタム関数のエラー処理は、[一般的な Excel JavaScript API のエラー処理](excel-add-ins-error-handling.md) と同じです。 以下のコード サンプルでは、`.catch` がコード内で発生するエラーを処理します。

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

現時点で Excel カスタム関数をデバッグするための最良の方法は、[ Excel Online ](../testing/sideload-office-add-ins-for-testing.md) 内でアドインを最初に** サイドロード** することです。次の手法と組み合わせて [お使いのブラウザーにネイティブのF12 デバッグ ツール](../testing/debug-add-ins-in-office-online.md) を使用して、カスタム関数をデバッグできます。

- カスタム関数内の `console.log` 文を使用し、コンソールにリアルタイムに出力を送信する。

- カスタム関数のコード内の`debugger;` 文を使用して、F12 ウィンドウが開いているときに実行が一時停止するブレークポイントを指定します。 たとえば、F12 ウィンドウが開いているときに次の関数を実行する場合は、実行が `debugger;` 文で一時停止し、関数が返される前にパラメーターの値を手動で検査することを有効にします。 F12 ウィンドウが開いていない場合、`debugger;` 文は Excel Online で影響を及ぼしません。 現在、 `debugger;` 文は Windows 版 Excel で効果がありません。

    ```js
    function add(first, second){
      debugger;
      return first + second;
    }
    ```

アドインの登録に失敗した場合は、アドイン アプリケーションをホストしている Web サーバーに、 [SSL 証明書が正しく構成されていることを確認してください](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) 。

Office 2016 のデスクトップでアドインをテストする場合、いくつかのインストールと実行時の条件と同様に、アドインの XML マニフェスト ファイルの問題をデバッグする [実行時のログ](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in) を有効にできます。

## <a name="mapping-function-names-to-json-metadata"></a>関数名を JSON のメタデータにマップする

[カスタム関数の概要](custom-functions-overview.md) で説明したとおり、カスタム関数プロジェクトは、カスタム関数を登録してエンド ユーザーが利用できるように Excel が必要とする情報を提供する、JSON のメタデータ ファイルを含める必要があります。 さらに、カスタム関数を定義する JavaScript ファイル内で、JSON のメタデータ ファイルにある関数オブジェクトを指定する情報は JavaScript ファイル内の各カスタム関数に対応を提供する必要があります。

たとえば、次のコード サンプルがカスタム関数 `add` を定義し、`id` プロパティの値が **追加**される JSON のメタデータ ファイル内のオブジェクトに関数 `add` が対応するよう指定します。

```js
function add(first, second){
  return first + second;
}

CustomFunctionMappings.ADD = add;
```

JavaScript ファイルでカスタム関数を作成し、JSON のメタデータ ファイルに対応する情報を指定するときは、次のベスト プラクティスに留意してください。

* JavaScript ファイルでは、キャメルケース (camenCase) で関数名を指定します。 たとえば、関数名 `addTenToInput` はキャメル ケースで記述します: 名前の最初の単語は小文字で開始し、後続の各単語は大文字で開始します。

* JSON メタデータ ファイルで、大文字で各 `name` プロパティの値を指定します。 `name` プロパティは、Excel でエンド ユーザーに表示される関数名を定義します。 各カスタム関数の名前の大文字を使用すると、Excel のエンド ユーザーに、すべての組み込み関数の名前が大文字であるという一貫性のあるエクスペリエンスを提供できます。

* JSON メタデータ ファイル内で、大文字で各 `id` プロパティの値を指定します。 これにより、JavaScript コードの`CustomFunctionMappings` 文のどの部分が、JSON メタデータ ファイルの`id` プロパティに対応しているかを明らかにします(上記で推奨したように、関数名にキャメル ケースを使用している場合)。

* JSON のメタデータ ファイルで、各 `id` プロパティの値が、ファイルのスコープ内で一意であることを確認します。 すなわち、メタデータ ファイル内で 2 つの関数オブジェクトが同じ `id` の値を持つことはありません。 さらに、大文字と小文字が異なるだけの 2 つの `id` の値をメタデータ ファイル内で指定しないでください。 たとえば、 **add**の値`id` の関数オブジェクトを、**ADD**の値`id` の別の関数オブジェクトと定義しないでください。

* 対応する JavaScript 関数名にマップした後、JSON メタデータ ファイルの `id` プロパティの値を変更しないでください。 JSON のメタデータ ファイル内の`name` プロパティを更新して、Excel でエンド ユーザーに表示される関数名を変更することができます。ただし、設定された後、 `id` プロパティの値は変更しないでください。

* JavaScript ファイルで、すべてのカスタム関数のマッピングを同じ場所で指定します。 たとえば、次のコード サンプルでは、2 つのカスタム関数を定義し、両方の関数のマッピング情報を指定しています。

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

## <a name="see-also"></a>関連項目

* [Excel でカスタム関数を作成する](custom-functions-overview.md)
* [カスタム関数のメタデータ](custom-functions-json.md)
* [Excel カスタム関数のランタイム](custom-functions-runtime.md)
* [Excel カスタム関数のチュートリアル](excel-tutorial-custom-functions.md)
