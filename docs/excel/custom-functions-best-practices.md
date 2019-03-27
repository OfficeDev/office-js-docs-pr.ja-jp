---
ms.date: 01/08/2019
description: Excel のカスタム関数を開発する際のベスト プラクティスについて説明します。
title: カスタム関数のベスト プラクティス (プレビュー)
localization_priority: Normal
ms.openlocfilehash: ae04169044336f7e42d341c1e904090e55d568af
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871347"
---
# <a name="custom-functions-best-practices-preview"></a>カスタム関数のベスト プラクティス (プレビュー)

この記事では、Excel でカスタム関数を開発するためのベスト プラクティスについて説明します。

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="troubleshooting"></a>トラブルシューティング

1. Windows 版 Office でアドインを検証する場合は、アドインの XML マニフェスト ファイルおよびインストールと実行時の条件のトラブルシューティングを行うために**[実行時ログ](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)** を有効にすることをおすすめします。 実行時ログは、ログファイルに `console.log` ステートメントを書き込んで、問題を発見します。

2. 1つ以上のカスタム関数が以前に登録されたアドインのカスタム関数と競合する場合、アドインは読み込まれません。 この場合、既存のアドインを削除するか、アドインの開発時にこのエラーが発生した場合は、マニフェストで別の名前空間名を指定することができます。

3. このトラブルシューティングの方法に関するフィードバックを Excel のユーザー設定関数チームに報告するには、チームにフィードバックを送信します。 これを行うには、**[ファイル] > [フィードバック] > [問題点、改善点の報告]** の順に選択します。 問題点や改善点の報告では、発生した問題を理解するために必要なログが提供されます。

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

## <a name="associating-function-names-with-json-metadata"></a>関数名を JSON メタデータに関連付ける

[カスタム関数の概要](custom-functions-overview.md)という記事で取り上げたように、カスタム関数プロジェクトには、カスタム関数を作成するために、JSON メタデータ ファイルとスクリプト (JavaScript または TypeScript) の両方を含める必要があります。 関数が正しく動作するには、スクリプト ファイル内の関数名を、JSON ファイルに記載されている ID にバインドしなければなりません。 このプロセスは関連付けと呼ばれます。 JavaScript コード ファイルの最後に関連付けを含める点に注意してください。そのようにしない限り、関数は動作しません。

次のコード サンプルは、この関連付けを実行する方法を示しています。 このサンプルではカスタム関数 `add` を定義し、それを `id` プロパティ値が **ADD** の、JSON メタデータ ファイル内のオブジェクトに関連付けます。

```js
function add(first, second){
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

JavaScript ファイルでカスタム関数を作成し、JSON のメタデータ ファイルに対応する情報を指定するときは、次のベスト プラクティスに留意してください。

* JSON メタデータ ファイルでは関数の `name` と `id` に大文字のみを使用します。 小文字と大文字を組み合わせたり、小文字のみを使用したりしないでください。 このような文字を使用すると、大文字小文字だけが異なる 2 つの値が存在するようになり、関数で意図しない上書きが生じる原因となる場合があります。 たとえば、`id` 値が **add** の関数オブジェクトが、`id` 値 **ADD** の関数オブジェクトのファイルに含まれる宣言によって後ほど上書きされる場合があります。 また `name` プロパティは、Excel でエンド ユーザーに表示される関数の名前を定義します。 各カスタム関数の名前の大文字を使用することで、すべての組み込み関数の名前は大文字である Excel で、一貫性のあるエクスペリエンスを提供します。

* ただし、関連付けを行う場合、関数の `name` を大文字にしなければならないわけではありません。 たとえば、`CustomFunctions.associate("add", add)` は `CustomFunctions.associate("ADD", add)` に相当します。

* JSON のメタデータ ファイルにそれぞれの `id` プロパティには、英数字とピリオドのみが含まれています。

* JSON のメタデータ ファイルで、各 `id` プロパティの値が、ファイルのスコープ内で一意であることを確認します。 すなわち、メタデータ ファイル内の 2 つの関数オブジェクトは同じ `id` 値であってはいけません。 

* 対応する JavaScript 関数の名前に関連付けられた後では、JSON のメタデータ ファイル内の `id` プロパティの値を変更しないでください。 JSON のメタデータ ファイル内の `name` プロパティを更新することによって Excel でエンド ユーザーに表示される関数の名前を変更することができます。しかし、確立された後は、 `id` プロパティの値を決して変更しないでください。

* JavaScript ファイルで同じ場所にすべてのカスタム関数の関連付けを指定します。 たとえば次のコード サンプルは、2 つのカスタム関数を定義し、両方の関数の関連付け情報を指定します。

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

    // associate `id` values in the JSON metadata file to JavaScript function names
    CustomFunctions.associate("ADD", add);
    CustomFunctions.associate("INCREMENT", increment);
    ```

    次のサンプルは、JavaScript コード サンプルで定義された関数に対応する JSON メタデータを示します。 `id` と `name` のプロパティがこのファイル内で大文字であることに注意してください。 

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
    "id": "ADD",
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
* [カスタム関数の変更ログ](custom-functions-changelog.md)
* [Excel カスタム関数のチュートリアル](../tutorials/excel-tutorial-create-custom-functions.md)
