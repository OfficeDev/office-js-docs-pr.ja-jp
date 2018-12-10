---
ms.date: 11/29/2018
description: Excel のカスタム関数を開発する際のベスト プラクティスについて説明します。
title: カスタム関数のベスト プラクティス
ms.openlocfilehash: b1785c7f41af9823cfd135ead29fff4eda4b0b1d
ms.sourcegitcommit: e2ba9d7210c921d068f40d9f689314c73ad5ab4a
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/05/2018
ms.locfileid: "27156587"
---
# <a name="custom-functions-best-practices-preview"></a><span data-ttu-id="41b21-103">カスタム関数のベスト プラクティス (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="41b21-103">Custom functions best practices (preview)</span></span>

<span data-ttu-id="41b21-104">この記事では、Excel でカスタム関数を開発するためのベスト プラクティスについて説明します。</span><span class="sxs-lookup"><span data-stu-id="41b21-104">This article describes best practices for developing custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="error-handling"></a><span data-ttu-id="41b21-105">エラー処理</span><span class="sxs-lookup"><span data-stu-id="41b21-105">Error handling</span></span>

<span data-ttu-id="41b21-106">カスタム関数を定義するアドインをビルドする場合は、実行時エラーを考慮して、エラー処理ロジックを含めるようにします。</span><span class="sxs-lookup"><span data-stu-id="41b21-106">When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="41b21-107">カスタム関数のエラー処理は、[全体的な Excel の JavaScript API のエラー処理](excel-add-ins-error-handling.md)と同じです。</span><span class="sxs-lookup"><span data-stu-id="41b21-107">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="41b21-108">次のコード サンプルでは、`.catch` がコード内で以前に発生したエラーを処理します。</span><span class="sxs-lookup"><span data-stu-id="41b21-108">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

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

## <a name="troubleshooting"></a><span data-ttu-id="41b21-109">トラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="41b21-109">Troubleshooting</span></span>

<span data-ttu-id="41b21-110">Windows 版 Office でアドインを検証する場合は、アドインの XML マニフェスト ファイルおよびインストールと実行時の条件のトラブルシューティングを行うために**[実行時ログ](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)** を有効にすることをおすすめします。</span><span class="sxs-lookup"><span data-stu-id="41b21-110">If you are testing your add-in in Office on Windows, you should enable **[runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)** to troubleshoot issues with your add-in's XML manifest file, as well as several installation and runtime conditions.</span></span> <span data-ttu-id="41b21-111">実行時ログは、ログファイルに `console.log` ステートメントを書き込んで、問題を発見します。</span><span class="sxs-lookup"><span data-stu-id="41b21-111">Runtime logging writes `console.log` statements to a log file to help you uncover issues.</span></span>

<span data-ttu-id="41b21-112">このトラブルシューティングの方法に関するフィードバックを Excel のユーザー設定関数チームに報告するには、チームにフィードバックを送信します。</span><span class="sxs-lookup"><span data-stu-id="41b21-112">To report feedback to the Excel Custom Functions team about this method of troubleshooting, send the team feedback.</span></span> <span data-ttu-id="41b21-113">これを行うには、**[ファイル] > [フィードバック] > [問題点、改善点の報告]** の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="41b21-113">To do this, select **File | Feedback | Send a Frown**.</span></span> <span data-ttu-id="41b21-114">問題点や改善点の報告では、発生した問題を理解するために必要なログが提供されます。</span><span class="sxs-lookup"><span data-stu-id="41b21-114">Sending a frown will provide the necessary logs to understand the issue you are hitting.</span></span>

## <a name="debugging"></a><span data-ttu-id="41b21-115">デバッグ</span><span class="sxs-lookup"><span data-stu-id="41b21-115">Debugging</span></span>

<span data-ttu-id="41b21-116">現時点で Excel カスタム関数をデバッグするための最良の方法は、**Excel Online** 内で最初にアドインを[サイドロード](../testing/sideload-office-add-ins-for-testing.md)する方法です。</span><span class="sxs-lookup"><span data-stu-id="41b21-116">Currently, the best method for debugging Excel custom functions is to first [sideload](../testing/sideload-office-add-ins-for-testing.md) your add-in within **Excel Online**.</span></span> <span data-ttu-id="41b21-117">その後に、次の手法と組み合わせて [お使いのブラウザでネイティブのF12 デバッグ ツール](../testing/debug-add-ins-in-office-online.md)を使用して、カスタム関数をデバッグできます。</span><span class="sxs-lookup"><span data-stu-id="41b21-117">You can then debug your custom functions by using the [F12 debugging tool native to your browser](../testing/debug-add-ins-in-office-online.md) in combination with the following techniques:</span></span>

- <span data-ttu-id="41b21-118">カスタム関数のコード内で `console.log` ステートメントを使用して、コンソールにリアルタイムに出力を送信します。</span><span class="sxs-lookup"><span data-stu-id="41b21-118">Use `console.log` statements within your custom functions code to send output to the console in real time.</span></span>

- <span data-ttu-id="41b21-119">カスタム関数コード内の `debugger;` ステートメントを使用して、F12 ウィンドウが開いているときに実行が一時停止するブレークポイントを指定します。</span><span class="sxs-lookup"><span data-stu-id="41b21-119">Use `debugger;` statements within your custom functions code to specify breakpoints where execution will pause when the F12 window is open.</span></span> <span data-ttu-id="41b21-120">例えば F12 ウィンドウが開いているときに以下の関数が動作している場合には、`debugger;` ステートメント上で実行が停止し、 関数が返される前に、パラメーター値を手動で検査することができます。</span><span class="sxs-lookup"><span data-stu-id="41b21-120">For example, if the following function runs while the F12 window is open, execution will pause on the `debugger;` statement, enabling you to manually inspect parameter values before the function returns.</span></span> <span data-ttu-id="41b21-121">`debugger;` ステートメントは、F12 ウィンドウが開いていない場合、Excel Online には影響しません。</span><span class="sxs-lookup"><span data-stu-id="41b21-121">The `debugger;` statement has no effect in Excel Online when the F12 window is not open.</span></span> <span data-ttu-id="41b21-122">現在、`debugger;` ステートメントは Windows 版 Excel には効果がありません。</span><span class="sxs-lookup"><span data-stu-id="41b21-122">Currently, the `debugger;` statement has no effect in Excel for Windows.</span></span>

    ```js
    function add(first, second){
      debugger;
      return first + second;
    }
    ```

<span data-ttu-id="41b21-123">アドインが登録に失敗した場合は、アドイン アプリケーションをホストしている Web サーバーに、 [SSL 証明書が正しく構成されていることを確認してください](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)。</span><span class="sxs-lookup"><span data-stu-id="41b21-123">If your add-in fails to register, [verify that SSL certificates are correctly configured](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for the web server that's hosting your add-in application.</span></span>

## <a name="mapping-function-names-to-json-metadata"></a><span data-ttu-id="41b21-124">関数名を JSON のメタデータにマップする</span><span class="sxs-lookup"><span data-stu-id="41b21-124">Mapping function names to JSON metadata</span></span>

<span data-ttu-id="41b21-125">[カスタム関数の概要](custom-functions-overview.md)資料で説明したように、カスタム関数プロジェクトには、カスタム関数を登録し、エンド ユーザーが利用できるように Excel が必要とする情報を提供する JSON のメタデータ ファイルを含める必要があります。</span><span class="sxs-lookup"><span data-stu-id="41b21-125">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users.</span></span> <span data-ttu-id="41b21-126">さらに、カスタム関数を定義する JavaScript ファイル内で、JSON のメタデータ ファイルにあるどの関数オブジェクトが JavaScript ファイル内の各カスタム関数に対応するかを指定する情報を提供する必要があります。</span><span class="sxs-lookup"><span data-stu-id="41b21-126">Additionally, within the JavaScript file that defines your custom functions, you must provide information to specify which function object in the JSON metadata file corresponds to each custom function in the JavaScript file.</span></span>

<span data-ttu-id="41b21-127">たとえば、次のコード サンプルは、カスタム関数 `add` を定義し、`id` プロパティの値が **ADD** である JSON のメタデータ ファイル内のオブジェクトに関数 `add` が対応するよう指定します。</span><span class="sxs-lookup"><span data-stu-id="41b21-127">For example, the following code sample defines the custom function `add` and then specifies that the function `add` corresponds to the object in the JSON metadata file where the value of the `id` property is **ADD**.</span></span>

```js
function add(first, second){
  return first + second;
}

CustomFunctionMappings.ADD = add;
```

<span data-ttu-id="41b21-128">JavaScript ファイルでカスタム関数を作成し、JSON のメタデータ ファイルに対応する情報を指定するときは、次のベスト プラクティスに留意してください。</span><span class="sxs-lookup"><span data-stu-id="41b21-128">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

* <span data-ttu-id="41b21-129">JavaScript ファイルでは関数名を キャメルケースで記述します。</span><span class="sxs-lookup"><span data-stu-id="41b21-129">In the JavaScript file, specify function names in camelCase.</span></span> <span data-ttu-id="41b21-130">たとえば、関数名 `addTenToInput` はキャメルケースで記述されています: 名前の最初の単語は小文字で開始し、後続の各単語は大文字で開始します。</span><span class="sxs-lookup"><span data-stu-id="41b21-130">For example, the function name `addTenToInput` is written in camelCase: the first word in the name starts with a lowercase letter and each subsequent word in the name starts with an uppercase letter.</span></span>

* <span data-ttu-id="41b21-131">JSON メタデータ ファイル内で、各 `name` プロパティの値に大文字を指定します。 </span><span class="sxs-lookup"><span data-stu-id="41b21-131">In the JSON metadata file, specify the value of each `name` property in uppercase.</span></span> <span data-ttu-id="41b21-132">`name` プロパティは、Excel でエンド ユーザーに表示される関数の名前を定義します。</span><span class="sxs-lookup"><span data-stu-id="41b21-132">The `name` property defines the function name that end users will see in Excel.</span></span> <span data-ttu-id="41b21-133">各カスタム関数の名前の大文字を使用することで、すべての組み込み関数の名前は大文字である Excel で、一貫性のあるエクスペリエンスをエンド ユーザーに提供します。</span><span class="sxs-lookup"><span data-stu-id="41b21-133">Using uppercase letters for the name of each custom function provides a consistent experience for end users in Excel, where all built-in function names are uppercase.</span></span>

* <span data-ttu-id="41b21-134">JSON メタデータ ファイル内で、各 `id` プロパティの値に大文字を指定します。 </span><span class="sxs-lookup"><span data-stu-id="41b21-134">In the JSON metadata file, specify the value of each `id` property in uppercase.</span></span> <span data-ttu-id="41b21-135">このようにすると、JavaScript コード内の `CustomFunctionMappings` ステートメントのどの部分が、JSON のメタデータの `id` プロパティに対応するかが明らかになります (推奨したように、関数名はキャメルケースを使用している前提で)。</span><span class="sxs-lookup"><span data-stu-id="41b21-135">Doing so makes it obvious which part of the `CustomFunctionMappings` statement in your JavaScript code corresponds to the `id` property in the JSON metadata file (provided that your function name uses camelCase, as recommended earlier).</span></span>

* <span data-ttu-id="41b21-136">JSON のメタデータ ファイルにそれぞれの `id` プロパティには、英数字とピリオドのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="41b21-136">In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.</span></span> 

* <span data-ttu-id="41b21-137">JSON のメタデータ ファイルで、各 `id` プロパティの値が、ファイルのスコープ内で一意であることを確認します。</span><span class="sxs-lookup"><span data-stu-id="41b21-137">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file.</span></span> <span data-ttu-id="41b21-138">すなわち、メタデータ ファイル内の 2 つの関数オブジェクトは同じ `id` 値であってはいけません。</span><span class="sxs-lookup"><span data-stu-id="41b21-138">That is, no two function objects in the metadata file should have the same `id` value.</span></span> <span data-ttu-id="41b21-139">さらに、2 つの大文字と小文字だけが異なるメタデータ ファイル内の `id` 値を指定しないでください。</span><span class="sxs-lookup"><span data-stu-id="41b21-139">Additionally, do not specify two `id` values in the metadata file that only differ by case.</span></span> <span data-ttu-id="41b21-140">たとえば、 **add**の値 `id` の関数オブジェクトを、**ADD**の値 `id` の別の関数オブジェクトと定義しないでください。</span><span class="sxs-lookup"><span data-stu-id="41b21-140">For example, do not define one function object with an `id` value of **add** and another function object with an `id` value of **ADD**.</span></span>

* <span data-ttu-id="41b21-141">対応する JavaScript 関数の名前にマップされた後では、JSON のメタデータ ファイル内の `id` プロパティの値を変更しないでください。</span><span class="sxs-lookup"><span data-stu-id="41b21-141">Do not change the value of an `id` property in the JSON metadata file after it's been mapped to a corresponding JavaScript function name.</span></span> <span data-ttu-id="41b21-142">JSON のメタデータ ファイル内の `name` プロパティを更新することによって Excel でエンド ユーザーに表示される関数の名前を変更することができます。しかし、確立された後は、 `id` プロパティの値を決して変更しないでください。</span><span class="sxs-lookup"><span data-stu-id="41b21-142">You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

* <span data-ttu-id="41b21-143">JavaScript ファイルで同じ場所にすべてのカスタム関数のマッピングを指定します。</span><span class="sxs-lookup"><span data-stu-id="41b21-143">In the JavaScript file, specify all custom function mappings in the same location.</span></span> <span data-ttu-id="41b21-144">例えば次のコード サンプルは、2 つのカスタム関数を定義し、両方の関数のマッピング情報を指定します。</span><span class="sxs-lookup"><span data-stu-id="41b21-144">For example, the following code sample defines two custom functions and then specifies the mapping information for both functions.</span></span>

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

    <span data-ttu-id="41b21-145">次のサンプルは、JavaScript コード サンプルで定義された関数に対応する JSON メタデータを示します。</span><span class="sxs-lookup"><span data-stu-id="41b21-145">The following sample shows the JSON metadata that corresponds to the functions defined in this JavaScript code sample.</span></span>

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

## <a name="declaring-optional-parameters"></a><span data-ttu-id="41b21-146">省略可能なパラメーターの宣言</span><span class="sxs-lookup"><span data-stu-id="41b21-146">Declaring optional parameters</span></span> 
<span data-ttu-id="41b21-147">Windows 版 Excel (バージョン 1812 以降) では、カスタム関数に省略可能なパラメーターを宣言できます。</span><span class="sxs-lookup"><span data-stu-id="41b21-147">In Excel for Windows (version 1812 or later), you can declare optional parameters for your custom functions.</span></span> <span data-ttu-id="41b21-148">ユーザーが Excel で関数を呼び出すと、角かっこで囲まれた省略可能なパラメーターが表示されます。</span><span class="sxs-lookup"><span data-stu-id="41b21-148">When a user invokes a function in Excel, optional parameters appear in brackets.</span></span> <span data-ttu-id="41b21-149">たとえば、関数 `FOO` に 1 つの必須パラメーター `parameter1` と 1 つの省略可能なパラメーター `parameter2` があるとすると、Excel では `=FOO(parameter1, [parameter2])` のように表示されます。</span><span class="sxs-lookup"><span data-stu-id="41b21-149">For example, a function `FOO` with one required parameter called `parameter1` and one optional parameter called `parameter2` would appear as `=FOO(parameter1, [parameter2])` in Excel.</span></span>

<span data-ttu-id="41b21-150">パラメーターを省略可能にするには、関数を定義している JSON メタデータ ファイルでパラメーターに `"optional": true` を追加します。</span><span class="sxs-lookup"><span data-stu-id="41b21-150">To make a parameter optional, add `"optional": true` to the parameter in the JSON metadata file that defines the function.</span></span> <span data-ttu-id="41b21-151">次の例では、関数 `=ADD(first, second, [third])` について、これがどのような内容になるかを示しています。</span><span class="sxs-lookup"><span data-stu-id="41b21-151">The following example shows what this might look like for the function `=ADD(first, second, [third])`.</span></span> <span data-ttu-id="41b21-152">省略可能な `[third]` パラメーターが 2 つの必須パラメーターの後にある点に注目してください。</span><span class="sxs-lookup"><span data-stu-id="41b21-152">Notice that the optional `[third]` parameter follows the two required parameters.</span></span> <span data-ttu-id="41b21-153">Excel の数式 UI では、必須パラメーターが最初に表示されます。</span><span class="sxs-lookup"><span data-stu-id="41b21-153">Required parameters will appear first in Excel’s Formula UI.</span></span>

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

<span data-ttu-id="41b21-154">関数の定義時に 1 つ以上の省略可能なパラメーターを含める場合は、省略可能なパラメーターが未定義のときの処理を指定しておく必要があります。</span><span class="sxs-lookup"><span data-stu-id="41b21-154">When you define a function that contains one or more optional parameters, you should specify what happens when the optional parameters are undefined.</span></span> <span data-ttu-id="41b21-155">次の例の `zipCode` と `dayOfWeek` は、どちらも `getWeatherReport` 関数の省略可能なパラメーターです。</span><span class="sxs-lookup"><span data-stu-id="41b21-155">In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function.</span></span> <span data-ttu-id="41b21-156">`zipCode` パラメーターが未定義の場合は、既定値が 98052 に設定されます。</span><span class="sxs-lookup"><span data-stu-id="41b21-156">If the `zipCode` parameter is undefined, the default value is set to 98052.</span></span> <span data-ttu-id="41b21-157">`dayOfWeek` パラメーターが未定義の場合は、Wednesday が設定されます。</span><span class="sxs-lookup"><span data-stu-id="41b21-157">If the `dayOfWeek` parameter is undefined, it is set to Wednesday.</span></span>

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

## <a name="additional-considerations"></a><span data-ttu-id="41b21-158">その他の考慮事項</span><span class="sxs-lookup"><span data-stu-id="41b21-158">Additional considerations</span></span>

<span data-ttu-id="41b21-159">複数のプラットフォーム (Office アドインの主要テナントの 1 つ) で動作するアドインを作成する際は、カスタム関数でドキュメント オブジェクト モデル (DOM) にアクセスしたり、jQuery のような DOM に依存するライブラリを使用してはいけません。 </span><span class="sxs-lookup"><span data-stu-id="41b21-159">In order to create an add-in that will run on multiple platforms (one of the key tenants of Office Add-ins), you should not access the Document Object Model (DOM) in custom functions or use libraries like jQuery that rely on the DOM.</span></span> <span data-ttu-id="41b21-160">カスタム関数が [JavaScript ランタイム](custom-functions-runtime.md)を使用する Excel for Windows では、カスタム関数はDOM にアクセスできません。</span><span class="sxs-lookup"><span data-stu-id="41b21-160">On Excel for Windows, where custom functions use the [JavaScript runtime](custom-functions-runtime.md), custom functions cannot access the DOM.</span></span>

## <a name="see-also"></a><span data-ttu-id="41b21-161">関連項目</span><span class="sxs-lookup"><span data-stu-id="41b21-161">See also</span></span>

* [<span data-ttu-id="41b21-162">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="41b21-162">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="41b21-163">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="41b21-163">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="41b21-164">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="41b21-164">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="41b21-165">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="41b21-165">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)
