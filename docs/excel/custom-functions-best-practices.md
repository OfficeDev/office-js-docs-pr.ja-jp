---
ms.date: 09/27/2018
description: Excel のカスタム関数のベスト プラクティスと推奨パターンについて説明します。
title: カスタム関数のベスト プラクティス
ms.openlocfilehash: 4590682a9efa3048605686763f9af28f2fad20a4
ms.sourcegitcommit: fdf7f4d686700edd6e6b04b2ea1bd43e59d4a03a
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/28/2018
ms.locfileid: "25348115"
---
# <a name="custom-functions-best-practices-preview"></a><span data-ttu-id="65f8e-103">カスタム関数のベスト プラクティス (プレビュー)</span><span class="sxs-lookup"><span data-stu-id="65f8e-103">Custom functions best practices</span></span>

<span data-ttu-id="65f8e-104">この記事は、Excel でカスタム関数を開発するためのベスト プラクティスについて説明します。</span><span class="sxs-lookup"><span data-stu-id="65f8e-104">This article describes best practices for developing custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="error-handling"></a><span data-ttu-id="65f8e-105">エラー処理</span><span class="sxs-lookup"><span data-stu-id="65f8e-105">Error handling</span></span>

<span data-ttu-id="65f8e-106">カスタム関数を定義するアドインを作成する場合は、実行時エラーに対処するエラー処理ロジックを含めてください。</span><span class="sxs-lookup"><span data-stu-id="65f8e-106">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="65f8e-107">カスタム関数のエラー処理は、[一般的な Excel JavaScript API のエラー処理](excel-add-ins-error-handling.md) と同じです。</span><span class="sxs-lookup"><span data-stu-id="65f8e-107">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="65f8e-108">以下のコード サンプルでは、`.catch` がコード内で発生するエラーを処理します。</span><span class="sxs-lookup"><span data-stu-id="65f8e-108">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

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

## <a name="debugging"></a><span data-ttu-id="65f8e-109">デバッグ</span><span class="sxs-lookup"><span data-stu-id="65f8e-109">Debugging</span></span>

<span data-ttu-id="65f8e-p102">現時点で Excel カスタム関数をデバッグするための最良の方法は、[ Excel Online ](../testing/sideload-office-add-ins-for-testing.md) 内でアドインを最初に\*\* サイドロード\*\* することです。次の手法と組み合わせて [お使いのブラウザでネイティブのF12 デバッグ ツール](../testing/debug-add-ins-in-office-online.md) を使用して、カスタム関数をデバッグできます。</span><span class="sxs-lookup"><span data-stu-id="65f8e-p102">Currently, the best method for debugging Excel custom functions is to first [sideload](../testing/sideload-office-add-ins-for-testing.md) your add-in within **Excel Online**. You can then debug your custom functions by using the [F12 debugging tool native to your browser](../testing/debug-add-ins-in-office-online.md). Use  statements within your custom functions code to send output to the console in real time.</span></span>

- <span data-ttu-id="65f8e-112">カスタム関数内の `console.log` 文を使用し、コンソールにリアルタイムに出力を送信する。</span><span class="sxs-lookup"><span data-stu-id="65f8e-112">Use `console.log` statements within your custom functions code to send output to the console in real time.</span></span>

- <span data-ttu-id="65f8e-113">カスタム関数のコード内で`debugger;` 文を使用して、F12 ウィンドウが開いているときに実行が一時停止するブレークポイントを指定します。</span><span class="sxs-lookup"><span data-stu-id="65f8e-113">Use `debugger;` statements within your custom functions code to specify breakpoints where execution will pause when the F12 window is open.</span></span> <span data-ttu-id="65f8e-114">たとえば、F12 ウィンドウが開いているときに次の関数を実行する場合は、実行が `debugger;` 文で一時停止し、関数が返される前にパラメーターの値を手動で検査することを有効にします。</span><span class="sxs-lookup"><span data-stu-id="65f8e-114">For example, if the following function runs while the F12 window is open, execution will pause on the `debugger;` statement, enabling you to manually inspect parameter values before the function returns.</span></span> <span data-ttu-id="65f8e-115">F12 ウィンドウが開いていない場合、`debugger;` 文は Excel Online に影響を及ぼしません。</span><span class="sxs-lookup"><span data-stu-id="65f8e-115">The `debugger;` statement has no effect in Excel Online when the F12 window is not open.</span></span> <span data-ttu-id="65f8e-116">現在、 `debugger;` 文は Windows 版 Excel で効果がありません。</span><span class="sxs-lookup"><span data-stu-id="65f8e-116">Currently, the `debugger;` statement has no effect in Excel for Windows.</span></span>

    ```js
    function add(first, second){
      debugger;
      return first + second;
    }
    ```

<span data-ttu-id="65f8e-117">アドインの登録に失敗した場合は、アドイン アプリケーションをホストしている Web サーバーで [SSL 証明書が正しく構成されていることを確認してください](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) 。</span><span class="sxs-lookup"><span data-stu-id="65f8e-117">If your add-in fails to register, [verify that SSL certificates are correctly configured](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for the web server that's hosting your add-in application.</span></span>

<span data-ttu-id="65f8e-118">Office 2016 のデスクトップでアドインをテストする場合、いくつかのインストールと実行時の条件と同様に、アドインの XML マニフェスト ファイルの問題をデバッグする [実行時のログ](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in) を有効にできます。</span><span class="sxs-lookup"><span data-stu-id="65f8e-118">If you are testing your add-in in Office 2016 desktop, you can enable [runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in) to debug issues with your add-in's XML manifest file as well as several installation and runtime conditions.</span></span>


## <a name="mapping-function-names-to-json-metadata"></a><span data-ttu-id="65f8e-119">関数名を JSON のメタデータにマップする</span><span class="sxs-lookup"><span data-stu-id="65f8e-119">Mapping function names to JSON metadata</span></span>

<span data-ttu-id="65f8e-120"> [カスタム関数の概要](custom-functions-overview.md) で説明したとおり、カスタム関数プロジェクトは、カスタム関数を登録してエンド ユーザーが利用できるように Excel が必要とする情報を提供する、JSON のメタデータ ファイルを含める必要があります。</span><span class="sxs-lookup"><span data-stu-id="65f8e-120">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users.</span></span> <span data-ttu-id="65f8e-121">さらに、カスタム関数を定義する JavaScript ファイル内で、JSON メタデータ ファイルにあるどの関数オブジェクトが、 JavaScript ファイル内の各カスタム関数に対応するかを指定するための情報を提供する必要があります。</span><span class="sxs-lookup"><span data-stu-id="65f8e-121">Additionally, within the JavaScript file that defines your custom functions, you must provide information to specify which function object in the JSON metadata file corresponds to each custom function in the JavaScript file.</span></span>

<span data-ttu-id="65f8e-122">たとえば、次のコード サンプルがカスタム関数 `add` を定義し、`id` プロパティの値が **追加**される JSON のメタデータ ファイル内のオブジェクトに関数 `add` が対応するよう指定します。</span><span class="sxs-lookup"><span data-stu-id="65f8e-122">For example, the following code sample defines the custom function `add` and then specifies that the function `add` corresponds to the object in the JSON metadata file where the value of the `id` property is **ADD**.</span></span>

```js
function add(first, second){
  return first + second;
}

CustomFunctionMappings.ADD = add;
```

<span data-ttu-id="65f8e-123">JavaScript ファイルでカスタム関数を作成し、JSON メタデータ ファイルに対応する情報を指定するときは、次のベスト プラクティスに留意してください。</span><span class="sxs-lookup"><span data-stu-id="65f8e-123">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

* <span data-ttu-id="65f8e-124">JavaScript ファイルでは、キャメル ケースで関数名を指定します。</span><span class="sxs-lookup"><span data-stu-id="65f8e-124">In the JavaScript file, specify function names in camelCase.</span></span> <span data-ttu-id="65f8e-125">たとえば、関数名 `addTenToInput` はキャメル ケースで記述します: 名前の最初の単語は小文字で開始し、後続の各単語は大文字で開始します。</span><span class="sxs-lookup"><span data-stu-id="65f8e-125">For example, the function name `addTenToInput` is written in camelCase: the first word in the name starts with a lowercase letter and each subsequent word in the name starts with an uppercase letter.</span></span>

* <span data-ttu-id="65f8e-126">JSON メタデータ ファイル内で、大文字で各 `name` プロパティの値を指定します。</span><span class="sxs-lookup"><span data-stu-id="65f8e-126">In the JSON metadata file, specify the value of each `name` property in uppercase.</span></span> <span data-ttu-id="65f8e-127"> `name` プロパティは、Excel でエンド ユーザーに表示される関数名を定義します。</span><span class="sxs-lookup"><span data-stu-id="65f8e-127">The `name` property defines the function name that end users will see in Excel.</span></span> <span data-ttu-id="65f8e-128">各カスタム関数の名前の大文字を使用して、すべての組み込み関数の名前が大文字の Excel で、エンド ユーザーに一貫性のあるエクスペリエンスを提供します。</span><span class="sxs-lookup"><span data-stu-id="65f8e-128">Using uppercase letters for the name of each custom function provides a consistent experience for end users in Excel, where all built-in function names are uppercase.</span></span>

* <span data-ttu-id="65f8e-129">JSON メタデータ ファイル内で、大文字で各 `id` プロパティの値を指定します。</span><span class="sxs-lookup"><span data-stu-id="65f8e-129">In the JSON metadata file, specify the value of each `id` property in uppercase.</span></span> <span data-ttu-id="65f8e-130">これにより、 JavaScript コード内の`CustomFunctionMappings` 文のどの部分が、JSON のメタデータ ファイルの`id` プロパティに対応しているかを明らかにします(推奨したように、関数名がキャメル ケースを使用している場合)。</span><span class="sxs-lookup"><span data-stu-id="65f8e-130">Doing so makes it obvious which part of the `CustomFunctionMappings` statement in your JavaScript code corresponds to the `id` property in the JSON metadata file (provided that your function name uses camelCase, as recommended earlier).</span></span>

* <span data-ttu-id="65f8e-131">JSON のメタデータ ファイルで、各 `id` プロパティの値が、ファイルのスコープ内で一意であることを確認します。</span><span class="sxs-lookup"><span data-stu-id="65f8e-131">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file.</span></span> <span data-ttu-id="65f8e-132">すなわち、メタデータ ファイル内で 2 つの関数オブジェクトが同じ `id` の値を持つことはありません。</span><span class="sxs-lookup"><span data-stu-id="65f8e-132">That is, no two function objects in the metadata file should have the same `id` value.</span></span> <span data-ttu-id="65f8e-133">さらに、大文字と小文字が異なるだけの 2 つの `id` の値をメタデータ ファイル内で指定しないでください。</span><span class="sxs-lookup"><span data-stu-id="65f8e-133">Additionally, do not specify two `id` values in the metadata file that only differ by case.</span></span> <span data-ttu-id="65f8e-134">たとえば、 **追加**の値`id` の関数オブジェクトを、**追加**の値`id` と別の関数オブジェクトと定義しないでください。</span><span class="sxs-lookup"><span data-stu-id="65f8e-134">For example, do not define one function object with an `id` value of **add** and another function object with an `id` value of **ADD**.</span></span>

* <span data-ttu-id="65f8e-135">対応する JavaScript 関数名にマップした後、JSON メタデータ ファイルの `id` プロパティの値を変更しないでください。</span><span class="sxs-lookup"><span data-stu-id="65f8e-135">Do not change the value of an `id` property in the JSON metadata file after it's been mapped to a corresponding JavaScript function name.</span></span> <span data-ttu-id="65f8e-136">JSON のメタデータ ファイル内の`name` プロパティを更新して、Excel でエンド ユーザーに表示される関数名を変更することができます。ただし、確立された後、 `id` プロパティの値は変更しないでください。</span><span class="sxs-lookup"><span data-stu-id="65f8e-136">You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

* <span data-ttu-id="65f8e-137">JavaScript ファイルで、すべてのカスタム関数のマッピングを同じ場所で指定します。</span><span class="sxs-lookup"><span data-stu-id="65f8e-137">In the JavaScript file, specify all custom function mappings in the same location.</span></span> <span data-ttu-id="65f8e-138">たとえば、次のコード サンプルは 2 つのカスタム関数を定義し、両方の関数のマッピング情報を指定します。</span><span class="sxs-lookup"><span data-stu-id="65f8e-138">For example, the following code sample defines two custom functions and then specifies the mapping information for both functions.</span></span>

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

    <span data-ttu-id="65f8e-139">次のサンプルは、JavaScript コード サンプルで定義された関数に対応する JSON メタデータを示します。</span><span class="sxs-lookup"><span data-stu-id="65f8e-139">The following sample shows the JSON metadata that corresponds to the functions defined in this JavaScript code sample.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="65f8e-140">関連項目</span><span class="sxs-lookup"><span data-stu-id="65f8e-140">See also</span></span>

- [<span data-ttu-id="65f8e-141">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="65f8e-141">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
- [<span data-ttu-id="65f8e-142">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="65f8e-142">Custom functions metadata</span></span>](custom-functions-json.md)
- [<span data-ttu-id="65f8e-143">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="65f8e-143">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
