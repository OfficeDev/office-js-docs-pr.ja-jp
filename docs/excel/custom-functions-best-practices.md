---
ms.date: 11/29/2018
description: Excel のカスタム関数を開発する際のベスト プラクティスについて説明します。
title: カスタム関数のベスト プラクティス
ms.openlocfilehash: c1be1d01a88d50bb0f3aee8af1aea7c47658bc10
ms.sourcegitcommit: 3007bf57515b0811ff98a7e1518ecc6fc9462276
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/04/2019
ms.locfileid: "27724887"
---
# <a name="custom-functions-best-practices-preview"></a>カスタム関数のベスト プラクティス (プレビュー)

この記事では、Excel でカスタム関数を開発するためのベスト プラクティスについて説明します。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="error-handling"></a>エラー処理

カスタム関数を定義するアドインをビルドする場合は、実行時エラーを考慮して、エラー処理ロジックを含めるようにします。 カスタム関数のエラー処理は、[全体的な Excel の JavaScript API のエラー処理](excel-add-ins-error-handling.md)と同じです。 次のコード サンプルでは、`.catch` がコード内で以前に発生したエラーを処理します。

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

## <a name="troubleshooting"></a>トラブルシューティング

Windows 版 Office でアドインを検証する場合は、アドインの XML マニフェスト ファイルおよびインストールと実行時の条件のトラブルシューティングを行うために**[実行時ログ](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)** を有効にすることをおすすめします。 実行時ログは、ログファイルに `console.log` ステートメントを書き込んで、問題を発見します。

このトラブルシューティングの方法に関するフィードバックを Excel のユーザー設定関数チームに報告するには、チームにフィードバックを送信します。 これを行うには、**[ファイル] > [フィードバック] > [問題点、改善点の報告]** の順に選択します。 問題点や改善点の報告では、発生した問題を理解するために必要なログが提供されます。

## <a name="debugging"></a>デバッグ

現時点で Excel カスタム関数をデバッグするための最良の方法は、**Excel Online** 内で最初にアドインを[サイドロード](../testing/sideload-office-add-ins-for-testing.md)する方法です。 その後に、次の手法と組み合わせて [お使いのブラウザでネイティブのF12 デバッグ ツール](../testing/debug-add-ins-in-office-online.md)を使用して、カスタム関数をデバッグできます。

- カスタム関数のコード内で `console.log` ステートメントを使用して、コンソールにリアルタイムに出力を送信します。

- カスタム関数コード内の `debugger;` ステートメントを使用して、F12 ウィンドウが開いているときに実行が一時停止するブレークポイントを指定します。 例えば F12 ウィンドウが開いているときに以下の関数が動作している場合には、`debugger;` ステートメント上で実行が停止し、 関数が返される前に、パラメーター値を手動で検査することができます。 `debugger;` ステートメントは、F12 ウィンドウが開いていない場合、Excel Online には影響しません。 現在、`debugger;` ステートメントは Windows 版 Excel には効果がありません。

    ```js
    function add(first, second){
      debugger;
      return first + second;
    }
    ```

アドインが登録に失敗した場合は、アドイン アプリケーションをホストしている Web サーバーに、 [SSL 証明書が正しく構成されていることを確認してください](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)。

## <a name="mapping-function-names-to-json-metadata"></a>関数名を JSON のメタデータにマップする

[カスタム関数の概要](custom-functions-overview.md)資料で説明したように、カスタム関数プロジェクトには、カスタム関数を登録し、エンド ユーザーが利用できるように Excel が必要とする情報を提供する JSON のメタデータ ファイルを含める必要があります。 さらに、カスタム関数を定義する JavaScript ファイル内で、JSON のメタデータ ファイルにあるどの関数オブジェクトが JavaScript ファイル内の各カスタム関数に対応するかを指定する情報を提供する必要があります。

たとえば、次のコード サンプルは、カスタム関数 `add` を定義し、`id` プロパティの値が **ADD** である JSON のメタデータ ファイル内のオブジェクトに関数 `add` が対応するよう指定します。

```js
function add(first, second){
  return first + second;
}

CustomFunctionMappings.ADD = add;
```

JavaScript ファイルでカスタム関数を作成し、JSON のメタデータ ファイルに対応する情報を指定するときは、次のベスト プラクティスに留意してください。

* JavaScript ファイルでは関数名を キャメルケースで記述します。 たとえば、関数名 `addTenToInput` はキャメルケースで記述されています: 名前の最初の単語は小文字で開始し、後続の各単語は大文字で開始します。

* JSON メタデータ ファイル内で、各 `name` プロパティの値に大文字を指定します。  `name` プロパティは、Excel でエンド ユーザーに表示される関数の名前を定義します。 各カスタム関数の名前の大文字を使用することで、すべての組み込み関数の名前は大文字である Excel で、一貫性のあるエクスペリエンスをエンド ユーザーに提供します。

* JSON メタデータ ファイル内で、各 `id` プロパティの値に大文字を指定します。  このようにすると、JavaScript コード内の `CustomFunctionMappings` ステートメントのどの部分が、JSON のメタデータの `id` プロパティに対応するかが明らかになります (推奨したように、関数名はキャメルケースを使用している前提で)。

* JSON のメタデータ ファイルにそれぞれの `id` プロパティには、英数字とピリオドのみが含まれています。 

* JSON のメタデータ ファイルで、各 `id` プロパティの値が、ファイルのスコープ内で一意であることを確認します。 すなわち、メタデータ ファイル内の 2 つの関数オブジェクトは同じ `id` 値であってはいけません。 さらに、2 つの大文字と小文字だけが異なるメタデータ ファイル内の `id` 値を指定しないでください。 たとえば、 **add**の値 `id` の関数オブジェクトを、**ADD**の値 `id` の別の関数オブジェクトと定義しないでください。

* 対応する JavaScript 関数の名前にマップされた後では、JSON のメタデータ ファイル内の `id` プロパティの値を変更しないでください。 JSON のメタデータ ファイル内の `name` プロパティを更新することによって Excel でエンド ユーザーに表示される関数の名前を変更することができます。しかし、確立された後は、 `id` プロパティの値を決して変更しないでください。

* JavaScript ファイルで同じ場所にすべてのカスタム関数のマッピングを指定します。 例えば次のコード サンプルは、2 つのカスタム関数を定義し、両方の関数のマッピング情報を指定します。

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

## <a name="declaring-optional-parameters"></a>省略可能なパラメーターの宣言 
Windows 版 Excel (バージョン 1812 以降) では、カスタム関数に省略可能なパラメーターを宣言できます。 ユーザーが Excel で関数を呼び出すと、角かっこで囲まれた省略可能なパラメーターが表示されます。 たとえば、関数 `FOO` に 1 つの必須パラメーター `parameter1` と 1 つの省略可能なパラメーター `parameter2` があるとすると、Excel では `=FOO(parameter1, [parameter2])` のように表示されます。

パラメーターを省略可能にするには、関数を定義している JSON メタデータ ファイルでパラメーターに `"optional": true` を追加します。 次の例では、関数 `=ADD(first, second, [third])` について、これがどのような内容になるかを示しています。 省略可能な `[third]` パラメーターが 2 つの必須パラメーターの後にある点に注目してください。 Excel の数式 UI では、必須パラメーターが最初に表示されます。

```json
{
    "id": "add",
    "name": "ADD",
    "description": "Add two numbers",
    "helpUrl": "http://www.contoso.com",
    "result": {
        "type": "number",
        "dimensionality": "scalar"
        },
    "parameters": [
        {
            "name": "first",
            "description": "first number to add",
            "type": "number",
            "dimensionality": "scalar"
        },
        {
            "name": "second",
            "description": "second number to add",
            "type": "number",
            "dimensionality": "scalar",
        },
        {
            "name": "third",
            "description": "third optional number to add",
            "type": "number",
            "dimensionality": "scalar",
            "optional": true
        }
    ],
    "options": {
        "sync": false
    }
}
```

関数の定義時に 1 つ以上の省略可能なパラメーターを含める場合は、省略可能なパラメーターが未定義のときの処理を指定しておく必要があります。 次の例の `zipCode` と `dayOfWeek` は、どちらも `getWeatherReport` 関数の省略可能なパラメーターです。 `zipCode` パラメーターが未定義の場合は、既定値が 98052 に設定されます。 `dayOfWeek` パラメーターが未定義の場合は、Wednesday が設定されます。

```js
function getWeatherReport(zipCode, dayOfWeek)
{
  if (zipCode === undefined) {
      zipCode = "98052";
  }

  if (dayOfWeek === undefined) {
    dayOfWeek = "Wednesday";
  }

  // Get weather report for specified zipCode and dayOfWeek
  // ...
}
```

## <a name="additional-considerations"></a>その他の考慮事項

複数のプラットフォーム (Office アドインの主要テナントの 1 つ) で動作するアドインを作成する際は、カスタム関数でドキュメント オブジェクト モデル (DOM) にアクセスしたり、jQuery のような DOM に依存するライブラリを使用してはいけません。  カスタム関数が [JavaScript ランタイム](custom-functions-runtime.md)を使用する Excel for Windows では、カスタム関数はDOM にアクセスできません。

## <a name="see-also"></a>関連項目

* [Excel でカスタム関数を作成する](custom-functions-overview.md)
* [カスタム関数のメタデータ](custom-functions-json.md)
* [Excel カスタム関数のランタイム](custom-functions-runtime.md)
* [Excel カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)
