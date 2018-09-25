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
# <a name="custom-functions-best-practices"></a><span data-ttu-id="13066-103">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="13066-103">Custom functions best practices</span></span>

<span data-ttu-id="13066-104">この記事は、Excel でカスタム関数を開発するためのベスト プラクティスについて説明します。</span><span class="sxs-lookup"><span data-stu-id="13066-104">This article describes best practices for developing custom functions in Excel.</span></span>

## <a name="error-handling"></a><span data-ttu-id="13066-105">エラー処理</span><span class="sxs-lookup"><span data-stu-id="13066-105">Error handling</span></span>

<span data-ttu-id="13066-106">カスタム関数を定義するアドインを作成する場合は、実行時エラーに対処するエラー処理ロジックを含めてください。</span><span class="sxs-lookup"><span data-stu-id="13066-106">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="13066-107">カスタム関数のエラー処理は、[一般的な Excel JavaScript API のエラー処理](excel-add-ins-error-handling.md) と同じです。</span><span class="sxs-lookup"><span data-stu-id="13066-107">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="13066-108">次のコード サンプルでは、`.catch` がコード内で発生するエラーを処理します。</span><span class="sxs-lookup"><span data-stu-id="13066-108">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

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

## <a name="error-logging"></a><span data-ttu-id="13066-109">エラー ログ</span><span class="sxs-lookup"><span data-stu-id="13066-109">Error logging</span></span>

<span data-ttu-id="13066-110">カスタム関数のエラーログは、次のような複数の方法で有効にすることができます。</span><span class="sxs-lookup"><span data-stu-id="13066-110">You can enable error logging for your custom functions add-in in multiple ways, such as:</span></span> 

- <span data-ttu-id="13066-111">アドインの XML マニフェスト ファイルをデバッグするために、[ 実行時ログを使用する](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in-manifest)。</span><span class="sxs-lookup"><span data-stu-id="13066-111">[Use runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in-manifest) to debug your add-in's XML manifest file.</span></span> 

- <span data-ttu-id="13066-112">カスタム関数内の `console.log` 文を使用し、コンソールにリアルタイムに出力を送信する。</span><span class="sxs-lookup"><span data-stu-id="13066-112">Use `console.log` statements within your custom functions code to send output to the console in real time.</span></span>

> [!NOTE]
> <span data-ttu-id="13066-113">現時点では、実行時ログ機能は Office 2016 デスクトップでのみ利用可能です。</span><span class="sxs-lookup"><span data-stu-id="13066-113">The runtime logging feature is currently available for Office 2016 desktop.</span></span>

## <a name="debugging"></a><span data-ttu-id="13066-114">デバッグ</span><span class="sxs-lookup"><span data-stu-id="13066-114">Debugging</span></span>

<span data-ttu-id="13066-115">現時点で Excel カスタム関数をデバッグするための最良の方法は、Excel Online 内でアドインを最初に[サイドロード](../testing/sideload-office-add-ins-for-testing.md)することです。</span><span class="sxs-lookup"><span data-stu-id="13066-115">Currently, the best method for debugging Excel custom functions is to first [sideload](../testing/sideload-office-add-ins-for-testing.md) your add-in within Excel Online.</span></span> <span data-ttu-id="13066-116"> [お使いのブラウザーにネイティブの F12 デバッグ ツール](../testing/debug-add-ins-in-office-online.md)を使用して、カスタム関数をデバッグできます。</span><span class="sxs-lookup"><span data-stu-id="13066-116">Then you can debug your custom functions by using the [F12 debugging tool native to your browser](../testing/debug-add-ins-in-office-online.md).</span></span>

<span data-ttu-id="13066-117">アドインの登録に失敗した場合は、アドイン アプリケーションをホストしている Web サーバーに、 [SSL 証明書が正しく構成されていることを確認してください](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) 。</span><span class="sxs-lookup"><span data-stu-id="13066-117">If your add-in fails to register, [verify that SSL certificates are correctly configured](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for the web server that's hosting your add-in application.</span></span>

## <a name="mapping-names"></a><span data-ttu-id="13066-118">名前のマッピング</span><span class="sxs-lookup"><span data-stu-id="13066-118">Mapping names</span></span>

<span data-ttu-id="13066-119">デフォルトでは、JavaScript ファイル内のカスタム関数の名前は通常すべて大文字を使用して宣言し、エンド ユーザーに Excel で表示される関数の名前と正確に対応します。</span><span class="sxs-lookup"><span data-stu-id="13066-119">By default, the name of a custom function in your JavaScript file is typically declared using entirely uppercase letters, and corresponds exactly to the function name that end users see in Excel.</span></span> <span data-ttu-id="13066-120">ただし、`CustomFunctionsMappings` オブジェクトを使用して、JavaScript ファイルの 1 つ以上の関数名を、Excel でエンド ユーザーに関数名として表示する他の値にマップするように変更できます。</span><span class="sxs-lookup"><span data-stu-id="13066-120">However, you can change this by using the `CustomFunctionsMappings` object to map one or more function names from the JavaScript file to different values that end users will see as function names in Excel.</span></span> <span data-ttu-id="13066-121">Uglifier、webpack、または大文字の関数名が困難なすべてのインポートの構文を使用している場合に便利です。</span><span class="sxs-lookup"><span data-stu-id="13066-121">Although you're not required to use , it can be helpful if you are using an uglifier, webpack, or import syntax - all of which have difficulty with uppercase function names.</span></span> <span data-ttu-id="13066-122">`CustomFunctionsMappings` プロジェクトが JavaScript を使用するのは恐らくオプションですが、プロジェクトが TypeScript を使用している場合は、使用する必要があります。</span><span class="sxs-lookup"><span data-stu-id="13066-122">`CustomFunctionsMappings` is possibly optional for projects using JavaScript but must be used if your project uses TypeScript.</span></span>  
  
<span data-ttu-id="13066-123">次のコード サンプルは、JavaScript 関数名 `plusFortyTwo` を、Excel UI の `ADD42` 関数名にマップする単一のキーと値のペアを定義しています。</span><span class="sxs-lookup"><span data-stu-id="13066-123">The following code sample defines a single key-value pair that maps the JavaScript function name `plusFortyTwo` to the `ADD42` function name in the Excel UI.</span></span> <span data-ttu-id="13066-124">エンド ユーザーが Excel で `ADD42` 関数を選択すると、`plusFortyTwo` JavaScript 関数が実行されます。</span><span class="sxs-lookup"><span data-stu-id="13066-124">When the end user chooses the `ADD42` function in Excel, the `plusFortyTwo` JavaScript function will run.</span></span>

```js
function plusFortyTwo(num) {
    return num + 42;  
}  
  
CustomFunctionsMappings = {
    "plusFortyTwo" : ADD42
}
```

<span data-ttu-id="13066-125">次のコード サンプルは、2 つのキーと値のペアを定義しています。</span><span class="sxs-lookup"><span data-stu-id="13066-125">The following code sample defines a two key-value pairs.</span></span> <span data-ttu-id="13066-126">最初のペアは、JavaScript 関数名 `plusFifty` を Excel UI の `ADD50` 関数名にマップし、2 番目のペアは、JavaScript 関数名 `plusOneHundred` を Excel UI の `ADD100` 関数名にマップします。</span><span class="sxs-lookup"><span data-stu-id="13066-126">The first pair maps the JavaScript function name `plusFifty` to the `ADD50` function name in the Excel UI, and the second pair maps the JavaScript function name `plusOneHundred` to the `ADD100` function name in the Excel UI.</span></span> <span data-ttu-id="13066-127">エンド ユーザーが Excel で `ADD50` 関数を選択すると、`plusFifty` JavaScript 関数が実行されます。</span><span class="sxs-lookup"><span data-stu-id="13066-127">When the end user chooses the `ADD50` function in Excel, the `plusFifty` JavaScript function will run.</span></span> <span data-ttu-id="13066-128">エンド ユーザーが Excel で `ADD100` 関数を選択すると、`plusOneHundred` JavaScript 関数が実行されます。</span><span class="sxs-lookup"><span data-stu-id="13066-128">When the end user chooses the `ADD100` function in Excel, the `plusOneHundred` JavaScript function will run.</span></span>

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

 ## <a name="see-also"></a><span data-ttu-id="13066-129">関連項目</span><span class="sxs-lookup"><span data-stu-id="13066-129">See also</span></span>

* [<span data-ttu-id="13066-130">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="13066-130">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="13066-131">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="13066-131">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="13066-132">Excel のカスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="13066-132">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
