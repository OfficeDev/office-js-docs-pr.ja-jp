---
ms.date: 10/17/2018
description: Excel のカスタム関数のベスト プラクティスと推奨パターンについて説明します。
title: カスタム関数のベスト プラクティス
ms.openlocfilehash: 10ba29966c1e991ca23674ce3e5da88de2772e00
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/19/2018
ms.locfileid: "25640002"
---
# <a name="custom-functions-best-practices-preview"></a><span data-ttu-id="3d36a-103">カスタム関数のベスト プラクティス (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="3d36a-103">Custom functions best practices</span></span>

<span data-ttu-id="3d36a-104">この記事は、Excel でカスタム関数を開発するためのベスト プラクティスについて説明します。</span><span class="sxs-lookup"><span data-stu-id="3d36a-104">This article describes best practices for developing custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="error-handling"></a><span data-ttu-id="3d36a-105">エラー処理</span><span class="sxs-lookup"><span data-stu-id="3d36a-105">Error handling</span></span>

<span data-ttu-id="3d36a-p101">カスタム関数を定義するアドインをビルドする場合は、実行時エラーを考慮するためのロジックを含めるようにしてください。カスタム関数のエラー処理は、 [大規模な Excel の JavaScript API のエラーの処理](excel-add-ins-error-handling.md)と同じです。次のコード サンプルでは、 `.catch` 、コード内で以前に発生したエラーを処理します。</span><span class="sxs-lookup"><span data-stu-id="3d36a-p101">When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors. Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md). In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

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

## <a name="debugging"></a><span data-ttu-id="3d36a-109">デバッグ</span><span class="sxs-lookup"><span data-stu-id="3d36a-109">Debugging</span></span>

<span data-ttu-id="3d36a-p102">現時点で Excel カスタム関数をデバッグするための最良の方法は、[ Excel Online ](../testing/sideload-office-add-ins-for-testing.md) 内でアドインを最初に\*\* サイドロード\*\* することです。次の手法と組み合わせて [お使いのブラウザでネイティブのF12 デバッグ ツール](../testing/debug-add-ins-in-office-online.md) を使用して、カスタム関数をデバッグできます。</span><span class="sxs-lookup"><span data-stu-id="3d36a-p102">Currently, the best method for debugging Excel custom functions is to first [sideload](../testing/sideload-office-add-ins-for-testing.md) your add-in within **Excel Online**. You can then debug your custom functions by using the [F12 debugging tool native to your browser](../testing/debug-add-ins-in-office-online.md) in combination with the following techniques:</span></span>

- <span data-ttu-id="3d36a-112">カスタム関数のコード内で `console.log` 文を使用して、コンソールにリアルタイムに出力を送信します。</span><span class="sxs-lookup"><span data-stu-id="3d36a-112">Use `console.log` statements within your custom functions code to send output to the console in real time.</span></span>

- <span data-ttu-id="3d36a-p103">カスタム関数コード内の `debugger;` 文を使用して、f12  ウィンドウが開いているときに実行が一時停止するブレークポイントを指定します、例えばf12  ウィンドウが開いているとき以下の関数が動作している場合には、`debugger;`文上で実行が停止し、 関数が返される前に、パラメーター値を手動で検査することができます。 `debugger;` 文は、F12 ウィンドウが開いていない場合、Excel Onlineには影響しません。。現在、 `debugger;` 文はwindows 版 Excel には効果がありません。</span><span class="sxs-lookup"><span data-stu-id="3d36a-p103">Use `debugger;` statements within your custom functions code to specify breakpoints where execution will pause when the F12 window is open. For example, if the following function runs while the F12 window is open, execution will pause on the `debugger;` statement, enabling you to manually inspect parameter values before the function returns. The `debugger;` statement has no effect in Excel Online when the F12 window is not open. Currently, the `debugger;` statement has no effect in Excel for Windows.</span></span>

    ```js
    function add(first, second){
      debugger;
      return first + second;
    }
    ```

<span data-ttu-id="3d36a-117">アドインが登録に失敗した場合は、アドイン アプリケーションをホストしている Web サーバーに、 [SSL 証明書が正しく構成されていることを確認してください](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) 。</span><span class="sxs-lookup"><span data-stu-id="3d36a-117">If your add-in fails to register, [verify that SSL certificates are correctly configured](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for the web server that's hosting your add-in application.</span></span>

<span data-ttu-id="3d36a-118">WindowsのデスクトップでOfficeのアドインをテストする場合、いくつかのインストールと実行時の条件と同様に、アドインの XML マニフェスト ファイルの問題をデバッグする [実行時のログ](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in) を有効にできます。</span><span class="sxs-lookup"><span data-stu-id="3d36a-118">If you are testing your add-in in Office 2016 desktop, you can enable [runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in) to debug issues with your add-in's XML manifest file as well as several installation and runtime conditions.</span></span>

## <a name="mapping-function-names-to-json-metadata"></a><span data-ttu-id="3d36a-119">関数名を JSON のメタデータにマップする</span><span class="sxs-lookup"><span data-stu-id="3d36a-119">Mapping function names to JSON metadata</span></span>

<span data-ttu-id="3d36a-p104">[カスタム関数の概要](custom-functions-overview.md) 資料で説明したように、カスタム関数プロジェクトには、カスタム関数を登録し、エンド ユーザーが利用できるように Excel が必要とする情報を提供する JSON のメタデータ ファイルを含める必要があります。さらに、カスタム関数を定義する JavaScript ファイル内で、JSON のメタデータ ファイルにあるどの関数オブジェクトが JavaScript ファイル内の各ユーザー定義関数に対応するかを指定する情報を提供する必要があります。</span><span class="sxs-lookup"><span data-stu-id="3d36a-p104">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users. Additionally, within the JavaScript file that defines your custom functions, you must provide information to specify which function object in the JSON metadata file corresponds to each custom function in the JavaScript file.</span></span>

<span data-ttu-id="3d36a-122">たとえば、次のコード サンプルは、カスタム関数 `add` を定義し、`id` プロパティの値が **ADD**である JSON のメタデータ ファイル内のオブジェクトに関数 `add` が対応するよう指定します。</span><span class="sxs-lookup"><span data-stu-id="3d36a-122">For example, the following code sample defines the custom function `add` and then specifies that the function `add` corresponds to the object in the JSON metadata file where the value of the `id` property is **ADD**.</span></span>

```js
function add(first, second){
  return first + second;
}

CustomFunctionMappings.ADD = add;
```

<span data-ttu-id="3d36a-123">JavaScript ファイルでカスタム関数を作成し、JSON のメタデータ ファイルに対応する情報を指定するときは、次のベスト プラクティスに留意してください。</span><span class="sxs-lookup"><span data-stu-id="3d36a-123">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

* <span data-ttu-id="3d36a-p105">JavaScript fileでは関数名をcamelCaseで記述します。たとえば、関数名 `addTenToInput` はcamelCaseで記述されています：名前の最初の単語は小文字で開始し、後続の各単語は大文字で開始します。</span><span class="sxs-lookup"><span data-stu-id="3d36a-p105">In the JavaScript file, specify function names in camelCase. For example, the function name `addTenToInput` is written in camelCase: the first word in the name starts with a lowercase letter and each subsequent word in the name starts with an uppercase letter.</span></span>

* <span data-ttu-id="3d36a-p106">JSON メタデータ ファイル内で、各`name` プロパティの値に大文字を指定します。 `name` プロパティは、Excel でエンド ユーザーに表示される関数の名前を定義します。各カスタム関数の名前の大文字を使用することで、すべての組み込み関数の名前は大文字である Excel で、一貫性のあるエクスペリエンスをエンド ユーザーに提供します。</span><span class="sxs-lookup"><span data-stu-id="3d36a-p106">In the JSON metadata file, specify the value of each `name` property in uppercase. The `name` property defines the function name that end users will see in Excel. Using uppercase letters for the name of each custom function provides a consistent experience for end users in Excel, where all built-in function names are uppercase.</span></span>

* <span data-ttu-id="3d36a-p107">JSON メタデータ ファイル内で、各 `id` プロパティの値に大文字を指定します。このようにすると、JavaScript コード内の `CustomFunctionMappings` 文のどの部分が、JSON のメタデータの `id`プロパティに対応するか明らかにしますのステートメントの対応するか明らかになります (推奨したように、関数名はcamelCaseを使用します) 。</span><span class="sxs-lookup"><span data-stu-id="3d36a-p107">In the JSON metadata file, specify the value of each `id` property in uppercase. Doing so makes it obvious which part of the `CustomFunctionMappings` statement in your JavaScript code corresponds to the `id` property in the JSON metadata file (provided that your function name uses camelCase, as recommended earlier).</span></span>

* <span data-ttu-id="3d36a-131">JSON のメタデータ ファイルに確実にそれぞれの値 `id` プロパティには、英数字とピリオドのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="3d36a-131">In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.</span></span> 

* <span data-ttu-id="3d36a-p108">JSON のメタデータ ファイルで、各 `id` プロパティの値が、ファイルのスコープ内で一意であることを確認します。すなわち、メタデータ ファイル内の 2 つの関数オブジェクトは同じ `id` 値であってはいけません。さらに、2 つの 大文字と小文字だけが異なるメタデータ ファイル内の`id` 値を指定しないでください。たとえば、 **add**の値`id`  の関数オブジェクトを、**ADD**の値`id`  の別の関数オブジェクトと定義しないでください。</span><span class="sxs-lookup"><span data-stu-id="3d36a-p108">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file. That is, no two function objects in the metadata file should have the same `id` value. Additionally, do not specify two `id` values in the metadata file that only differ by case. For example, do not define one function object with an `id` value of **add** and another function object with an `id` value of **ADD**.</span></span>

* <span data-ttu-id="3d36a-p109">対応する JavaScript 関数の名前にマップされた後では、JSON のメタデータ ファイル内の `id` プロパティの値を変更しないでください。JSON のメタデータ ファイル内の`name` プロパティを更新することによって Excel でエンド ユーザーに表示される関数の名前を変更することができます。しかし、確立された後は、  `id` プロパティの値を決して変更しないでください。</span><span class="sxs-lookup"><span data-stu-id="3d36a-p109">Do not change the value of an `id` property in the JSON metadata file after it's been mapped to a corresponding JavaScript function name. You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

* <span data-ttu-id="3d36a-p110">JavaScript ファイルで同じ場所にすべてのカスタム関数のマッピングを指定します。例えば次のコード サンプル は、2 つのカスタム関数を定義し、両方の関数のマッピング情報を指定します。</span><span class="sxs-lookup"><span data-stu-id="3d36a-p110">In the JavaScript file, specify all custom function mappings in the same location. For example, the following code sample defines two custom functions and then specifies the mapping information for both functions.</span></span>

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

    <span data-ttu-id="3d36a-140">次のサンプルは、JavaScript コード サンプルで定義された関数に対応する JSON メタデータを示します。</span><span class="sxs-lookup"><span data-stu-id="3d36a-140">The following sample shows the JSON metadata that corresponds to the functions defined in this JavaScript code sample.</span></span>

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

## <a name="additional-considerations"></a><span data-ttu-id="3d36a-141">その他の考慮事項</span><span class="sxs-lookup"><span data-stu-id="3d36a-141">Additional considerations</span></span>

<span data-ttu-id="3d36a-142">複数のプラットフォーム（Officeアドインの主要テナントの一つ）で動作するアドインを作成する際は、カスタム関数でドキュメント オブジェクト モデル (DOM) にアクセスしたり、jQueryのようなDOMに依存するライブラリーを使用してはいけません。</span><span class="sxs-lookup"><span data-stu-id="3d36a-142">In order to create an add-in that will run on multiple platforms (one of the key tenants of Office Add-ins), you should not access the Document Object Model (DOM) in custom functions or use libraries like jQuery that rely on the DOM.</span></span> <span data-ttu-id="3d36a-143">カスタム関数が [JavaScript のランタイム](custom-functions-runtime.md)を使用するExcel for Windows のウィンドウでは、ユーザー定義関数はDOM にアクセスできません。</span><span class="sxs-lookup"><span data-stu-id="3d36a-143">On Excel for Windows, where custom functions use the [JavaScript runtime](custom-functions-runtime.md), custom functions cannot access the DOM.</span></span>

## <a name="see-also"></a><span data-ttu-id="3d36a-144">関連項目</span><span class="sxs-lookup"><span data-stu-id="3d36a-144">See also</span></span>

* [<span data-ttu-id="3d36a-145">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="3d36a-145">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="3d36a-146">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="3d36a-146">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="3d36a-147">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="3d36a-147">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="3d36a-148">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="3d36a-148">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)
