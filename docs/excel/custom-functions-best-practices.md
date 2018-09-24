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
# <a name="custom-functions-best-practices"></a><span data-ttu-id="40ac0-103">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="40ac0-103">Custom functions best practices</span></span>

<span data-ttu-id="40ac0-104">この記事は、Excel でカスタム関数を開発するためのベスト プラクティスについて説明します。</span><span class="sxs-lookup"><span data-stu-id="40ac0-104">This article describes best practices for developing custom functions in Excel.</span></span>

## <a name="error-handling"></a><span data-ttu-id="40ac0-105">エラー処理</span><span class="sxs-lookup"><span data-stu-id="40ac0-105">Error handling</span></span>

<span data-ttu-id="40ac0-106">カスタム関数を定義するアドインを作成する場合は、実行時エラーを含むエラー処理ロジックを含めてください。</span><span class="sxs-lookup"><span data-stu-id="40ac0-106">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="40ac0-107">カスタム関数のエラー処理は、「[一般的な Excel JavaScript API のエラー処理](excel-add-ins-error-handling.md) 」と同じです。</span><span class="sxs-lookup"><span data-stu-id="40ac0-107">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="40ac0-108">次のコード サンプルでは、`.catch` がコード内で発生するエラーを処理します。</span><span class="sxs-lookup"><span data-stu-id="40ac0-108">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

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

## <a name="error-logging"></a><span data-ttu-id="40ac0-109">エラー ログ</span><span class="sxs-lookup"><span data-stu-id="40ac0-109">Error logging</span></span>

<span data-ttu-id="40ac0-110">カスタム関数のエラーログは、次のような複数の方法で有効にすることができます。</span><span class="sxs-lookup"><span data-stu-id="40ac0-110">You can enable error logging for your custom functions add-in in multiple ways, such as:</span></span> 

- <span data-ttu-id="40ac0-111">アドインの XML マニフェスト ファイルをデバッグするために、[ 実行時ログを使用する](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in-manifest)。</span><span class="sxs-lookup"><span data-stu-id="40ac0-111">[Use runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in-manifest) to debug your add-in's XML manifest file.</span></span> 

- <span data-ttu-id="40ac0-112">カスタム関数内の `console.log` 文を使用し、コンソールにリアルタイムに出力を送信する。</span><span class="sxs-lookup"><span data-stu-id="40ac0-112">Use `console.log` statements within your custom functions code to send output to the console in real time.</span></span>

> [!NOTE]
> <span data-ttu-id="40ac0-113">現時点では、実行時ログ機能は Office 2016 デスクトップでのみ利用可能です。</span><span class="sxs-lookup"><span data-stu-id="40ac0-113">The runtime logging feature is currently available for Office 2016 desktop.</span></span>

## <a name="debugging"></a><span data-ttu-id="40ac0-114">デバッグ</span><span class="sxs-lookup"><span data-stu-id="40ac0-114">Debugging</span></span>

<span data-ttu-id="40ac0-115">現時点で、Excel のカスタム関数をデバッグする最良の方法は、 [Excel Online](https://www.office.com/launch/excel) を使用し、使用するブラウザに対応する F12 デバッグ ツールを使用する方法です。</span><span class="sxs-lookup"><span data-stu-id="40ac0-115">Currently, the best method for debugging Excel custom functions is to use [Excel Online](https://www.office.com/launch/excel) and use the F12 debugging tool native to your browser.</span></span> <span data-ttu-id="40ac0-116">将来的には、カスタム関数用の他のデバッグ ツールも利用できる可能性があります。</span><span class="sxs-lookup"><span data-stu-id="40ac0-116">Additional debugging tools for custom functions may be available in the future.</span></span>

## <a name="mapping-names"></a><span data-ttu-id="40ac0-117">名前のマッピング</span><span class="sxs-lookup"><span data-stu-id="40ac0-117">Mapping names</span></span>

<span data-ttu-id="40ac0-118">デフォルトでは、JavaScript ファイル内のカスタム関数の名前は通常すべて大文字を使用して宣言し、エンド ユーザーに Excel で表示される関数の名前と正確に対応します。</span><span class="sxs-lookup"><span data-stu-id="40ac0-118">By default, the name of a custom function in your JavaScript file is typically declared using entirely uppercase letters, and corresponds exactly to the function name that end users see in Excel.</span></span> <span data-ttu-id="40ac0-119">ただし、`CustomFunctionsMappings` オブジェクトを使用して、JavaScript ファイルの 1 つ以上の関数名を、Excel でエンド ユーザーに関数名として表示する他の値にマップするように変更できます。</span><span class="sxs-lookup"><span data-stu-id="40ac0-119">However, you can change this by using the `CustomFunctionsMappings` object to map one or more function names from the JavaScript file to different values that end users will see as function names in Excel.</span></span> <span data-ttu-id="40ac0-120">`CustomFunctionsMapping` を使用する必要はありませんが、大文字の関数名では問題が生じる uglifier、webpack、import 構文などを使用している場合に役立ちます。</span><span class="sxs-lookup"><span data-stu-id="40ac0-120">Although you're not required to use `CustomFunctionsMapping`, it can be helpful if you are using an uglifier, webpack, or import syntax - all of which have difficulty with uppercase function names.</span></span>
  
<span data-ttu-id="40ac0-121">次のコード サンプルは、JavaScript 関数名 `plusFortyTwo` を、Excel UI の `ADD42` 関数名にマップする単一のキーと値のペアを定義しています。</span><span class="sxs-lookup"><span data-stu-id="40ac0-121">The following code sample defines a single key-value pair that maps the JavaScript function name `plusFortyTwo` to the `ADD42` function name in the Excel UI.</span></span> <span data-ttu-id="40ac0-122">エンド ユーザーが Excel で `ADD42` 関数を選択すると、`plusFortyTwo` JavaScript 関数が実行されます。</span><span class="sxs-lookup"><span data-stu-id="40ac0-122">When the end user chooses the `ADD42` function in Excel, the `plusFortyTwo` JavaScript function will run.</span></span>

```js
function plusFortyTwo(num) {
    return num + 42;  
}  
  
CustomFunctionsMappings = {
    "plusFortyTwo" : ADD42
}
```

<span data-ttu-id="40ac0-123">次のコード サンプルは、2 つのキーと値のペアを定義しています。</span><span class="sxs-lookup"><span data-stu-id="40ac0-123">The following code sample defines a two key-value pairs.</span></span> <span data-ttu-id="40ac0-124">最初のペアは、JavaScript 関数名 `plusFifty` を Excel UI の `ADD50` 関数名にマップし、2 番目のペアは、JavaScript 関数名 `plusOneHundred` を Excel UI の `ADD100` 関数名にマップします。</span><span class="sxs-lookup"><span data-stu-id="40ac0-124">The first pair maps the JavaScript function name `plusFifty` to the `ADD50` function name in the Excel UI, and the second pair maps the JavaScript function name `plusOneHundred` to the `ADD100` function name in the Excel UI.</span></span> <span data-ttu-id="40ac0-125">エンド ユーザーが Excel で `ADD50` 関数を選択すると、`plusFifty` JavaScript 関数が実行されます。</span><span class="sxs-lookup"><span data-stu-id="40ac0-125">When the end user chooses the `ADD50` function in Excel, the `plusFifty` JavaScript function will run.</span></span> <span data-ttu-id="40ac0-126">エンド ユーザーが Excel で `ADD100` 関数を選択すると、`plusOneHundred` JavaScript 関数が実行されます。</span><span class="sxs-lookup"><span data-stu-id="40ac0-126">When the end user chooses the `ADD100` function in Excel, the `plusOneHundred` JavaScript function will run.</span></span>

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

 ## <a name="see-also"></a><span data-ttu-id="40ac0-127">関連項目</span><span class="sxs-lookup"><span data-stu-id="40ac0-127">See also</span></span>

* [<span data-ttu-id="40ac0-128">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="40ac0-128">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="40ac0-129">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="40ac0-129">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="40ac0-130">Excel のカスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="40ac0-130">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)