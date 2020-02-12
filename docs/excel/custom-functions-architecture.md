---
ms.date: 07/10/2019
description: Excelのカスタム関数のランタイムについて解説します。
title: カスタム関数のアーキテクチャ
localization_priority: Normal
ms.openlocfilehash: ced62f7efb826862eee8079a66fa657ea466e4b3
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950355"
---
# <a name="custom-functions-architecture"></a><span data-ttu-id="0ef0f-103">カスタム関数のアーキテクチャ</span><span class="sxs-lookup"><span data-stu-id="0ef0f-103">Custom functions architecture</span></span>

 <span data-ttu-id="0ef0f-104">カスタム関数は、計算の実行の優先付けを行う独自のランタイムを持っています。</span><span class="sxs-lookup"><span data-stu-id="0ef0f-104">Custom functions are with their own unique runtime that prioritizes execution of calculations.</span></span> <span data-ttu-id="0ef0f-105">この記事では、カスタム関数ランタイムと、アドインの他の部分を駆動するブラウザベースのJavaScriptエンジンの違いについて説明します。</span><span class="sxs-lookup"><span data-stu-id="0ef0f-105">This article will cover the differences between the custom functions runtime and the browser-based JavaScript engine which powers most other parts of your add-in.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="custom-functions-runtime"></a><span data-ttu-id="0ef0f-106">カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="0ef0f-106">Custom functions runtime</span></span>

<span data-ttu-id="0ef0f-107">Office Webアドインは、作業ウィンドウまたはコンテンツウィンドウとしてユーザーと対話したり、コマンドやカスタム機能を含めることができます。</span><span class="sxs-lookup"><span data-stu-id="0ef0f-107">An Office Web Add-in can interact with the user as a task pane, or a content pane, and can include commands and custom functions.</span></span> <span data-ttu-id="0ef0f-108">カスタム関数を除いて、これらすべての部分はブラウザエンジンランタイムで動作します。</span><span class="sxs-lookup"><span data-stu-id="0ef0f-108">All of these parts run in a browser engine runtime except for custom functions.</span></span> <span data-ttu-id="0ef0f-109">カスタム関数は、計算速度を最適化する別のカスタム関数の実行時に実行します。</span><span class="sxs-lookup"><span data-stu-id="0ef0f-109">Custom functions run in a separate custom functions runtime to optimize for calculation speed.</span></span>

<span data-ttu-id="0ef0f-110">プロジェクトの生成に [Office アドイン用の Yeoman ジェネレーター](https://www.npmjs.com/package/generator-office)を使用している場合は、カスタム関数ランタイムは **functions.html** ファイルで参照されている custom-functions.js スクリプト ファイルを介して読み込みます。</span><span class="sxs-lookup"><span data-stu-id="0ef0f-110">Note that if you're using the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) to generate your project, the custom functions runtime will load through the custom-functions.js script file referenced in the **functions.html** file.</span></span> <span data-ttu-id="0ef0f-111">**functions.html** は、ランタイムを読み込む場合にのみ機能し、アドイン用の作業ウィンドウとして使用することはできません。</span><span class="sxs-lookup"><span data-stu-id="0ef0f-111">The **functions.html** serves only to load the runtime and shouldn't be used as the task pane for your add-in.</span></span>

<span data-ttu-id="0ef0f-112">次の表は、カスタム関数の実行時とブラウザーのエンジンの実行時の違いを示しています。</span><span class="sxs-lookup"><span data-stu-id="0ef0f-112">The following table highlights the differences between the custom functions runtime and the browser engine runtime:</span></span>

| <span data-ttu-id="0ef0f-113">カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="0ef0f-113">Custom functions runtime</span></span>  | <span data-ttu-id="0ef0f-114">ブラウザエンジン ランタイム</span><span class="sxs-lookup"><span data-stu-id="0ef0f-114">Browser engine runtime</span></span>    |
|------------------------------------------------------------------ |-------------------------------------------------------------------------------------------------------------- |
| <span data-ttu-id="0ef0f-115">セルの値を返すことをサポートしています</span><span class="sxs-lookup"><span data-stu-id="0ef0f-115">Supports returning a value from a cell</span></span>    | <span data-ttu-id="0ef0f-116">Office.js Api と UI 要素をサポートしています。</span><span class="sxs-lookup"><span data-stu-id="0ef0f-116">Supports Office.js APIs and UI elements</span></span>   |
| <span data-ttu-id="0ef0f-117">`localStorage` オブジェクトを持たず、代わりに `OfficeRuntime.storage` オブジェクトを使用します。</span><span class="sxs-lookup"><span data-stu-id="0ef0f-117">Does not have `localStorage` object, instead uses the `OfficeRuntime.storage` object.</span></span>     | <span data-ttu-id="0ef0f-118">`localStorage` オブジェクトを持ち, オプションで `OfficeRuntime.storage` オブジェクトを使用することもできます。</span><span class="sxs-lookup"><span data-stu-id="0ef0f-118">Has `localStorage` object, can optionally use the `OfficeRuntime.storage` object.</span></span>     |
| <span data-ttu-id="0ef0f-119">DOM の関連操作や、jQuery など DOM に依存するライブラリの読み込みはサポートされていません。</span><span class="sxs-lookup"><span data-stu-id="0ef0f-119">Does not support interacting with the DOM, or loading libraries that depend on the DOM such as jQuery.</span></span>    | <span data-ttu-id="0ef0f-120">DOM の関連操作や、DOM に依存するライブラリの読み込みがサポートされています。</span><span class="sxs-lookup"><span data-stu-id="0ef0f-120">Supports interacting with the DOM and loading libraries that depend on the DOM.</span></span> |

## <a name="browser-engine-runtime"></a><span data-ttu-id="0ef0f-121">ブラウザエンジン ランタイム</span><span class="sxs-lookup"><span data-stu-id="0ef0f-121">Browser engine runtime</span></span>

<span data-ttu-id="0ef0f-122">作業ウィンドウ、コンテンツアドイン、およびコマンドは、ブラウザエンジンランタイムで実行されます。</span><span class="sxs-lookup"><span data-stu-id="0ef0f-122">The task pane, content add-in, and commands run in a browser engine runtime.</span></span>

<span data-ttu-id="0ef0f-123">ブラウザエンジン ランタイムは、Office.js Api をサポートしています。</span><span class="sxs-lookup"><span data-stu-id="0ef0f-123">The browser engine runtime supports the Office.js APIs.</span></span> <span data-ttu-id="0ef0f-124">Excelのテーブルを操作できるAPIなどのExcel APIは、ブラウザエンジンランタイムで実行されますが、カスタム関数ランタイムから直接アクセスすることはできません。</span><span class="sxs-lookup"><span data-stu-id="0ef0f-124">Keep in mind that any of the Excel APIs, such as APIs which allow you to manipulate Excel tables, run on the browser engine runtime, but aren't directly accessible from the custom functions runtime.</span></span>

## <a name="communicate-between-runtimes"></a><span data-ttu-id="0ef0f-125">ランタイム間のコミュニケーション</span><span class="sxs-lookup"><span data-stu-id="0ef0f-125">Communicate between runtimes</span></span>

<span data-ttu-id="0ef0f-126">カスタム関数のコードは、実行時間が異なるため、作業ウィンドウのようにWebアドインの他の部分のコードと直接対話することはできません。</span><span class="sxs-lookup"><span data-stu-id="0ef0f-126">Your custom functions code cannot directly interact with code in other parts of your web add-in, like the task pane because they are in different runtimes.</span></span> <span data-ttu-id="0ef0f-127">ただし、一部のシナリオでは、トークンを渡すなどのデータを共有する必要があります。</span><span class="sxs-lookup"><span data-stu-id="0ef0f-127">But in some scenarios you may need to share data, such as passing a token.</span></span>

<span data-ttu-id="0ef0f-128">`OfficeRuntime.storage` オブジェクトを、カスタム関数からのデータを保存したり、作業ウィンドウのコードからデータを取得したりするために使用できます。</span><span class="sxs-lookup"><span data-stu-id="0ef0f-128">The `OfficeRuntime.storage` object can be used to store data from your custom functions and get data from your task pane code.</span></span> <span data-ttu-id="0ef0f-129">データの保管と共有の詳細については、「[状態の保存と共有](custom-functions-save-state.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0ef0f-129">For more information about storing and sharing data, see [Save and share state](custom-functions-save-state.md).</span></span>

<span data-ttu-id="0ef0f-130">パターンとプラクティス専用の [Githubリポジトリ](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) で `storage` オブジェクトを使用してコード サンプルを見ることができます。</span><span class="sxs-lookup"><span data-stu-id="0ef0f-130">You can see a code sample using the `storage` object in this [Github repository](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) dedicated to patterns and practices.</span></span>
<span data-ttu-id="0ef0f-131">`storage` オブジェクトに関する一般的な情報の詳細については、「[カスタム関数ランタイム](./custom-functions-runtime.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0ef0f-131">For more general information about the `storage` object, see [Custom functions runtime](./custom-functions-runtime.md).</span></span>

<span data-ttu-id="0ef0f-132">`storage` オブジェクトは認証にも役立つ場合があります。</span><span class="sxs-lookup"><span data-stu-id="0ef0f-132">The `storage` object can also be useful for authentication.</span></span> <span data-ttu-id="0ef0f-133">詳細については、[カスタム関数の認証](custom-functions-authentication.md)を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0ef0f-133">For more information, see [Custom functions authentication](custom-functions-authentication.md).</span></span>

## <a name="next-steps"></a><span data-ttu-id="0ef0f-134">次の手順</span><span class="sxs-lookup"><span data-stu-id="0ef0f-134">Next steps</span></span>
<span data-ttu-id="0ef0f-135">詳細については、「[カスタム関数ランタイムの使用](custom-functions-runtime.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="0ef0f-135">Learn more about how to [use the custom functions runtime](custom-functions-runtime.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="0ef0f-136">関連項目</span><span class="sxs-lookup"><span data-stu-id="0ef0f-136">See also</span></span>

* [<span data-ttu-id="0ef0f-137">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="0ef0f-137">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="0ef0f-138">カスタム関数でデータを受信して​​処理する</span><span class="sxs-lookup"><span data-stu-id="0ef0f-138">Receive and handle data with custom functions</span></span>](custom-functions-web-reqs.md)
* [<span data-ttu-id="0ef0f-139">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="0ef0f-139">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="0ef0f-140">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="0ef0f-140">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
