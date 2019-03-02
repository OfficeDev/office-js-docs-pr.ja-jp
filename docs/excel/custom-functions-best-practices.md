---
ms.date: 01/08/2019
description: Excel のカスタム関数を開発する際のベスト プラクティスについて説明します。
title: カスタム関数のベスト プラクティス (プレビュー)
localization_priority: Normal
ms.openlocfilehash: 24c73ec643df073ac97dc399343a7feb0b0b4168
ms.sourcegitcommit: f7f3d38ae4430e2218bf0abe7bb2976108de3579
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/01/2019
ms.locfileid: "30359262"
---
# <a name="custom-functions-best-practices-preview"></a><span data-ttu-id="bd96f-103">カスタム関数のベスト プラクティス (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="bd96f-103">Custom functions best practices (preview)</span></span>

<span data-ttu-id="bd96f-104">この記事では、Excel でカスタム関数を開発するためのベスト プラクティスについて説明します。</span><span class="sxs-lookup"><span data-stu-id="bd96f-104">This article describes best practices for developing custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="error-handling"></a><span data-ttu-id="bd96f-105">エラー処理</span><span class="sxs-lookup"><span data-stu-id="bd96f-105">Error handling</span></span>

<span data-ttu-id="bd96f-106">カスタム関数を定義するアドインをビルドする場合は、実行時エラーを考慮して、エラー処理ロジックを含めるようにします。</span><span class="sxs-lookup"><span data-stu-id="bd96f-106">When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="bd96f-107">カスタム関数のエラー処理は、[全体的な Excel の JavaScript API のエラー処理](excel-add-ins-error-handling.md)と同じです。</span><span class="sxs-lookup"><span data-stu-id="bd96f-107">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="bd96f-108">次のコード サンプルでは、`.catch` がコード内で以前に発生したエラーを処理します。</span><span class="sxs-lookup"><span data-stu-id="bd96f-108">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

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

## <a name="troubleshooting"></a><span data-ttu-id="bd96f-109">トラブルシューティング</span><span class="sxs-lookup"><span data-stu-id="bd96f-109">Troubleshooting</span></span>

1. <span data-ttu-id="bd96f-110">Windows 版 Office でアドインを検証する場合は、アドインの XML マニフェスト ファイルおよびインストールと実行時の条件のトラブルシューティングを行うために**[実行時ログ](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)** を有効にすることをおすすめします。</span><span class="sxs-lookup"><span data-stu-id="bd96f-110">If you are testing your add-in in Office on Windows, you should enable **[runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)** to troubleshoot issues with your add-in's XML manifest file, as well as several installation and runtime conditions.</span></span> <span data-ttu-id="bd96f-111">実行時ログは、ログファイルに `console.log` ステートメントを書き込んで、問題を発見します。</span><span class="sxs-lookup"><span data-stu-id="bd96f-111">Runtime logging writes `console.log` statements to a log file to help you uncover issues.</span></span>

2. <span data-ttu-id="bd96f-112">1つ以上のカスタム関数が以前に登録されたアドインのカスタム関数と競合する場合、アドインは読み込まれません。</span><span class="sxs-lookup"><span data-stu-id="bd96f-112">Your add-in will not load if one or more custom functions conflicts with a previously registered add-in's custom functions.</span></span> <span data-ttu-id="bd96f-113">この場合、既存のアドインを削除するか、アドインの開発時にこのエラーが発生した場合は、マニフェストで別の名前空間名を指定することができます。</span><span class="sxs-lookup"><span data-stu-id="bd96f-113">In this case, you can either remove the existing add-in, or if you encounter this error while developing an add-in, you can specify a different namespace name in your manifest.</span></span>

3. <span data-ttu-id="bd96f-114">このトラブルシューティングの方法に関するフィードバックを Excel のユーザー設定関数チームに報告するには、チームにフィードバックを送信します。</span><span class="sxs-lookup"><span data-stu-id="bd96f-114">To report feedback to the Excel Custom Functions team about this method of troubleshooting, send the team feedback.</span></span> <span data-ttu-id="bd96f-115">これを行うには、**[ファイル] > [フィードバック] > [問題点、改善点の報告]** の順に選択します。</span><span class="sxs-lookup"><span data-stu-id="bd96f-115">To do this, select **File | Feedback | Send a Frown**.</span></span> <span data-ttu-id="bd96f-116">問題点や改善点の報告では、発生した問題を理解するために必要なログが提供されます。</span><span class="sxs-lookup"><span data-stu-id="bd96f-116">Sending a frown will provide the necessary logs to understand the issue you are hitting.</span></span>


## <a name="debugging"></a><span data-ttu-id="bd96f-117">デバッグ</span><span class="sxs-lookup"><span data-stu-id="bd96f-117">Debugging</span></span>

<span data-ttu-id="bd96f-118">現時点で Excel カスタム関数をデバッグするための最良の方法は、**Excel Online** 内で最初にアドインを[サイドロード](../testing/sideload-office-add-ins-for-testing.md)する方法です。</span><span class="sxs-lookup"><span data-stu-id="bd96f-118">Currently, the best method for debugging Excel custom functions is to first [sideload](../testing/sideload-office-add-ins-for-testing.md) your add-in within **Excel Online**.</span></span> <span data-ttu-id="bd96f-119">その後に、次の手法と組み合わせて [お使いのブラウザでネイティブのF12 デバッグ ツール](../testing/debug-add-ins-in-office-online.md)を使用して、カスタム関数をデバッグできます。</span><span class="sxs-lookup"><span data-stu-id="bd96f-119">You can then debug your custom functions by using the [F12 debugging tool native to your browser](../testing/debug-add-ins-in-office-online.md) in combination with the following techniques:</span></span>

- <span data-ttu-id="bd96f-120">カスタム関数のコード内で `console.log` ステートメントを使用して、コンソールにリアルタイムに出力を送信します。</span><span class="sxs-lookup"><span data-stu-id="bd96f-120">Use `console.log` statements within your custom functions code to send output to the console in real time.</span></span>

- <span data-ttu-id="bd96f-121">カスタム関数コード内の `debugger;` ステートメントを使用して、F12 ウィンドウが開いているときに実行が一時停止するブレークポイントを指定します。</span><span class="sxs-lookup"><span data-stu-id="bd96f-121">Use `debugger;` statements within your custom functions code to specify breakpoints where execution will pause when the F12 window is open.</span></span> <span data-ttu-id="bd96f-122">例えば F12 ウィンドウが開いているときに以下の関数が動作している場合には、`debugger;` ステートメント上で実行が停止し、 関数が返される前に、パラメーター値を手動で検査することができます。</span><span class="sxs-lookup"><span data-stu-id="bd96f-122">For example, if the following function runs while the F12 window is open, execution will pause on the `debugger;` statement, enabling you to manually inspect parameter values before the function returns.</span></span> <span data-ttu-id="bd96f-123">`debugger;` ステートメントは、F12 ウィンドウが開いていない場合、Excel Online には影響しません。</span><span class="sxs-lookup"><span data-stu-id="bd96f-123">The `debugger;` statement has no effect in Excel Online when the F12 window is not open.</span></span> <span data-ttu-id="bd96f-124">現在、`debugger;` ステートメントは Windows 版 Excel には効果がありません。</span><span class="sxs-lookup"><span data-stu-id="bd96f-124">Currently, the `debugger;` statement has no effect in Excel for Windows.</span></span>

    ```js
    function add(first, second){
      debugger;
      return first + second;
    }
    ```

<span data-ttu-id="bd96f-125">アドインが登録に失敗した場合は、アドイン アプリケーションをホストしている Web サーバーに、 [SSL 証明書が正しく構成されていることを確認してください](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md)。</span><span class="sxs-lookup"><span data-stu-id="bd96f-125">If your add-in fails to register, [verify that SSL certificates are correctly configured](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for the web server that's hosting your add-in application.</span></span>

## <a name="associating-function-names-with-json-metadata"></a><span data-ttu-id="bd96f-126">関数名を JSON メタデータに関連付ける</span><span class="sxs-lookup"><span data-stu-id="bd96f-126">Associating function names with JSON metadata</span></span>

<span data-ttu-id="bd96f-127">[カスタム関数の概要](custom-functions-overview.md)という記事で取り上げたように、カスタム関数プロジェクトには、カスタム関数を作成するために、JSON メタデータ ファイルとスクリプト (JavaScript または TypeScript) の両方を含める必要があります。</span><span class="sxs-lookup"><span data-stu-id="bd96f-127">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include both a JSON metadata file and a script (either JavaScript or TypeScript) file to form a complete function.</span></span> <span data-ttu-id="bd96f-128">関数が正しく動作するには、スクリプト ファイル内の関数名を、JSON ファイルに記載されている ID にバインドしなければなりません。</span><span class="sxs-lookup"><span data-stu-id="bd96f-128">For a function to work properly, you'll need to bind the name of the function in the script file to the id listed in the JSON file.</span></span> <span data-ttu-id="bd96f-129">このプロセスは関連付けと呼ばれます。</span><span class="sxs-lookup"><span data-stu-id="bd96f-129">This process is called association.</span></span> <span data-ttu-id="bd96f-130">JavaScript コード ファイルの最後に関連付けを含める点に注意してください。そのようにしない限り、関数は動作しません。</span><span class="sxs-lookup"><span data-stu-id="bd96f-130">Make a note to include associations at the end of your JavaScript code files; otherwise, your functions will not work.</span></span>

<span data-ttu-id="bd96f-131">次のコード サンプルは、この関連付けを実行する方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="bd96f-131">The following code sample shows how to do this association.</span></span> <span data-ttu-id="bd96f-132">このサンプルではカスタム関数 `add` を定義し、それを `id` プロパティ値が **ADD** の、JSON メタデータ ファイル内のオブジェクトに関連付けます。</span><span class="sxs-lookup"><span data-stu-id="bd96f-132">The sample defines the custom function `add` and associates it with the object in the JSON metadata file where the value of the `id` property is **ADD**.</span></span>

```js
function add(first, second){
  return first + second;
}

CustomFunctions.associate("ADD", add); 
```

<span data-ttu-id="bd96f-133">JavaScript ファイルでカスタム関数を作成し、JSON のメタデータ ファイルに対応する情報を指定するときは、次のベスト プラクティスに留意してください。</span><span class="sxs-lookup"><span data-stu-id="bd96f-133">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

* <span data-ttu-id="bd96f-134">JSON メタデータ ファイルでは関数の `name` と `id` に大文字のみを使用します。</span><span class="sxs-lookup"><span data-stu-id="bd96f-134">Only use uppercase letters for a function's `name` and `id` in the JSON metadata file.</span></span> <span data-ttu-id="bd96f-135">小文字と大文字を組み合わせたり、小文字のみを使用したりしないでください。</span><span class="sxs-lookup"><span data-stu-id="bd96f-135">Do not use a mix of cases or only lowercase letters.</span></span> <span data-ttu-id="bd96f-136">このような文字を使用すると、大文字小文字だけが異なる 2 つの値が存在するようになり、関数で意図しない上書きが生じる原因となる場合があります。</span><span class="sxs-lookup"><span data-stu-id="bd96f-136">If you do, you may end up with two values that only differ by case which will cause unintentional overwriting of your functions.</span></span> <span data-ttu-id="bd96f-137">たとえば、`id` 値が **add** の関数オブジェクトが、`id` 値 **ADD** の関数オブジェクトのファイルに含まれる宣言によって後ほど上書きされる場合があります。</span><span class="sxs-lookup"><span data-stu-id="bd96f-137">For example, a function object with an `id` value of **add** could be overwritten by declaration later in the file of function object with an `id` value of **ADD**.</span></span> <span data-ttu-id="bd96f-138">また `name` プロパティは、Excel でエンド ユーザーに表示される関数の名前を定義します。</span><span class="sxs-lookup"><span data-stu-id="bd96f-138">Additionally, the `name` property defines the function name that end users will see in Excel.</span></span> <span data-ttu-id="bd96f-139">各カスタム関数の名前の大文字を使用することで、すべての組み込み関数の名前は大文字である Excel で、一貫性のあるエクスペリエンスを提供します。</span><span class="sxs-lookup"><span data-stu-id="bd96f-139">Using uppercase letters for the name of each custom function provides a consistent experience in Excel, where all built-in function names are uppercase.</span></span>

* <span data-ttu-id="bd96f-140">ただし、関連付けを行う場合、関数の `name` を大文字にしなければならないわけではありません。</span><span class="sxs-lookup"><span data-stu-id="bd96f-140">However, it is not necessary to capitalize the function's `name` when associating.</span></span> <span data-ttu-id="bd96f-141">たとえば、`CustomFunctions.associate("add", add)` は `CustomFunctions.associate("ADD", add)` に相当します。</span><span class="sxs-lookup"><span data-stu-id="bd96f-141">For example, `CustomFunctions.associate("add", add)` is equivalent to `CustomFunctions.associate("ADD", add)`.</span></span>

* <span data-ttu-id="bd96f-142">JSON のメタデータ ファイルにそれぞれの `id` プロパティには、英数字とピリオドのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="bd96f-142">In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.</span></span>

* <span data-ttu-id="bd96f-143">JSON のメタデータ ファイルで、各 `id` プロパティの値が、ファイルのスコープ内で一意であることを確認します。</span><span class="sxs-lookup"><span data-stu-id="bd96f-143">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file.</span></span> <span data-ttu-id="bd96f-144">すなわち、メタデータ ファイル内の 2 つの関数オブジェクトは同じ `id` 値であってはいけません。</span><span class="sxs-lookup"><span data-stu-id="bd96f-144">That is, no two function objects in the metadata file should have the same `id` value.</span></span> 

* <span data-ttu-id="bd96f-145">対応する JavaScript 関数の名前に関連付けられた後では、JSON のメタデータ ファイル内の `id` プロパティの値を変更しないでください。</span><span class="sxs-lookup"><span data-stu-id="bd96f-145">Do not change the value of an `id` property in the JSON metadata file after it's been associated with a corresponding JavaScript function name.</span></span> <span data-ttu-id="bd96f-146">JSON のメタデータ ファイル内の `name` プロパティを更新することによって Excel でエンド ユーザーに表示される関数の名前を変更することができます。しかし、確立された後は、 `id` プロパティの値を決して変更しないでください。</span><span class="sxs-lookup"><span data-stu-id="bd96f-146">You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

* <span data-ttu-id="bd96f-147">JavaScript ファイルで同じ場所にすべてのカスタム関数の関連付けを指定します。</span><span class="sxs-lookup"><span data-stu-id="bd96f-147">In the JavaScript file, specify all custom function associations in the same location.</span></span> <span data-ttu-id="bd96f-148">たとえば次のコード サンプルは、2 つのカスタム関数を定義し、両方の関数の関連付け情報を指定します。</span><span class="sxs-lookup"><span data-stu-id="bd96f-148">For example, the following code sample defines two custom functions and then specifies the association information for both functions.</span></span>

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

    <span data-ttu-id="bd96f-149">次のサンプルは、JavaScript コード サンプルで定義された関数に対応する JSON メタデータを示します。</span><span class="sxs-lookup"><span data-stu-id="bd96f-149">The following sample shows the JSON metadata that corresponds to the functions defined in this JavaScript code sample.</span></span> <span data-ttu-id="bd96f-150">`id` と `name` のプロパティがこのファイル内で大文字であることに注意してください。</span><span class="sxs-lookup"><span data-stu-id="bd96f-150">Note that the `id` and `name` properties are in uppercase letters in this file.</span></span> 

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

## <a name="declaring-optional-parameters"></a><span data-ttu-id="bd96f-151">省略可能なパラメーターの宣言</span><span class="sxs-lookup"><span data-stu-id="bd96f-151">Declaring optional parameters</span></span> 
<span data-ttu-id="bd96f-152">Windows 版 Excel (バージョン 1812 以降) では、カスタム関数に省略可能なパラメーターを宣言できます。</span><span class="sxs-lookup"><span data-stu-id="bd96f-152">In Excel for Windows (version 1812 or later), you can declare optional parameters for your custom functions.</span></span> <span data-ttu-id="bd96f-153">ユーザーが Excel で関数を呼び出すと、角かっこで囲まれた省略可能なパラメーターが表示されます。</span><span class="sxs-lookup"><span data-stu-id="bd96f-153">When a user invokes a function in Excel, optional parameters appear in brackets.</span></span> <span data-ttu-id="bd96f-154">たとえば、関数 `FOO` に 1 つの必須パラメーター `parameter1` と 1 つの省略可能なパラメーター `parameter2` があるとすると、Excel では `=FOO(parameter1, [parameter2])` のように表示されます。</span><span class="sxs-lookup"><span data-stu-id="bd96f-154">For example, a function `FOO` with one required parameter called `parameter1` and one optional parameter called `parameter2` would appear as `=FOO(parameter1, [parameter2])` in Excel.</span></span>

<span data-ttu-id="bd96f-155">パラメーターを省略可能にするには、関数を定義している JSON メタデータ ファイルでパラメーターに `"optional": true` を追加します。</span><span class="sxs-lookup"><span data-stu-id="bd96f-155">To make a parameter optional, add `"optional": true` to the parameter in the JSON metadata file that defines the function.</span></span> <span data-ttu-id="bd96f-156">次の例では、関数 `=ADD(first, second, [third])` について、これがどのような内容になるかを示しています。</span><span class="sxs-lookup"><span data-stu-id="bd96f-156">The following example shows what this might look like for the function `=ADD(first, second, [third])`.</span></span> <span data-ttu-id="bd96f-157">省略可能な `[third]` パラメーターが 2 つの必須パラメーターの後にある点に注目してください。</span><span class="sxs-lookup"><span data-stu-id="bd96f-157">Notice that the optional `[third]` parameter follows the two required parameters.</span></span> <span data-ttu-id="bd96f-158">Excel の数式 UI では、必須パラメーターが最初に表示されます。</span><span class="sxs-lookup"><span data-stu-id="bd96f-158">Required parameters will appear first in Excel’s Formula UI.</span></span>

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

<span data-ttu-id="bd96f-159">関数の定義時に 1 つ以上の省略可能なパラメーターを含める場合は、省略可能なパラメーターが未定義のときの処理を指定しておく必要があります。</span><span class="sxs-lookup"><span data-stu-id="bd96f-159">When you define a function that contains one or more optional parameters, you should specify what happens when the optional parameters are undefined.</span></span> <span data-ttu-id="bd96f-160">次の例の `zipCode` と `dayOfWeek` は、どちらも `getWeatherReport` 関数の省略可能なパラメーターです。</span><span class="sxs-lookup"><span data-stu-id="bd96f-160">In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function.</span></span> <span data-ttu-id="bd96f-161">`zipCode` パラメーターが未定義の場合は、既定値が 98052 に設定されます。</span><span class="sxs-lookup"><span data-stu-id="bd96f-161">If the `zipCode` parameter is undefined, the default value is set to 98052.</span></span> <span data-ttu-id="bd96f-162">`dayOfWeek` パラメーターが未定義の場合は、Wednesday が設定されます。</span><span class="sxs-lookup"><span data-stu-id="bd96f-162">If the `dayOfWeek` parameter is undefined, it is set to Wednesday.</span></span>

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

## <a name="additional-considerations"></a><span data-ttu-id="bd96f-163">その他の考慮事項</span><span class="sxs-lookup"><span data-stu-id="bd96f-163">Additional considerations</span></span>

<span data-ttu-id="bd96f-164">複数のプラットフォーム (Office アドインの主要テナントの 1 つ) で動作するアドインを作成する際は、カスタム関数でドキュメント オブジェクト モデル (DOM) にアクセスしたり、jQuery のような DOM に依存するライブラリを使用してはいけません。 </span><span class="sxs-lookup"><span data-stu-id="bd96f-164">In order to create an add-in that will run on multiple platforms (one of the key tenants of Office Add-ins), you should not access the Document Object Model (DOM) in custom functions or use libraries like jQuery that rely on the DOM.</span></span> <span data-ttu-id="bd96f-165">カスタム関数が [JavaScript ランタイム](custom-functions-runtime.md)を使用する Excel for Windows では、カスタム関数はDOM にアクセスできません。</span><span class="sxs-lookup"><span data-stu-id="bd96f-165">On Excel for Windows, where custom functions use the [JavaScript runtime](custom-functions-runtime.md), custom functions cannot access the DOM.</span></span>

## <a name="see-also"></a><span data-ttu-id="bd96f-166">関連項目</span><span class="sxs-lookup"><span data-stu-id="bd96f-166">See also</span></span>

* [<span data-ttu-id="bd96f-167">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="bd96f-167">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="bd96f-168">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="bd96f-168">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="bd96f-169">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="bd96f-169">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="bd96f-170">カスタム関数の変更ログ</span><span class="sxs-lookup"><span data-stu-id="bd96f-170">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="bd96f-171">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="bd96f-171">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
