---
ms.date: 06/18/2019
description: Excel のカスタム関数を開発する際のベスト プラクティスについて説明します。
title: カスタム関数のベスト プラクティス
localization_priority: Normal
ms.openlocfilehash: 7c836119a783f5cc7e1e7f4f52f1d21b86091bfe
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127934"
---
# <a name="custom-functions-best-practices"></a><span data-ttu-id="ad5fd-103">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="ad5fd-103">Custom functions best practices</span></span>

<span data-ttu-id="ad5fd-104">この記事では、Excel でカスタム関数を開発するためのベスト プラクティスについて説明します。</span><span class="sxs-lookup"><span data-stu-id="ad5fd-104">This article describes best practices for developing custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="associating-function-names-with-json-metadata"></a><span data-ttu-id="ad5fd-105">関数名を JSON メタデータに関連付ける</span><span class="sxs-lookup"><span data-stu-id="ad5fd-105">Associating function names with JSON metadata</span></span>

<span data-ttu-id="ad5fd-106">[カスタム関数の概要](custom-functions-overview.md)という記事で取り上げたように、カスタム関数プロジェクトには、カスタム関数を作成するために、JSON メタデータ ファイルとスクリプト (JavaScript または TypeScript) の両方を含める必要があります。</span><span class="sxs-lookup"><span data-stu-id="ad5fd-106">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include both a JSON metadata file and a script (either JavaScript or TypeScript) file to form a complete function.</span></span> <span data-ttu-id="ad5fd-107">JSON メタデータを`yo office`使用している場合は、コードコメントから生成することができます。</span><span class="sxs-lookup"><span data-stu-id="ad5fd-107">If you are using `yo office` the JSON metadata can be generated from the code comments.</span></span> <span data-ttu-id="ad5fd-108">それ以外の場合は、JSON メタデータファイルを手動でビルドする必要があります。</span><span class="sxs-lookup"><span data-stu-id="ad5fd-108">Otherwise you need to build the JSON metadata file manually.</span></span>

<span data-ttu-id="ad5fd-109">関数が正しく動作するには、関数の`id`プロパティを JavaScript 実装に関連付ける必要があります。</span><span class="sxs-lookup"><span data-stu-id="ad5fd-109">For a function to work properly, you need to associate the function's `id` property with the JavaScript implementation.</span></span> <span data-ttu-id="ad5fd-110">関連付けがあることを確認してください。それ以外の場合は、関数は呼び出されません。</span><span class="sxs-lookup"><span data-stu-id="ad5fd-110">Make sure there is an association, otherwise the function will not be called.</span></span> <span data-ttu-id="ad5fd-111">次のコードサンプルは、 `CustomFunctions.associate()`メソッドを使用して関連付けを行う方法を示しています。</span><span class="sxs-lookup"><span data-stu-id="ad5fd-111">The following code sample shows how to make the association using the `CustomFunctions.associate()` method.</span></span> <span data-ttu-id="ad5fd-112">このサンプルではカスタム関数 `add` を定義し、それを `id` プロパティ値が **ADD** の、JSON メタデータ ファイル内のオブジェクトに関連付けます。</span><span class="sxs-lookup"><span data-stu-id="ad5fd-112">The sample defines the custom function `add` and associates it with the object in the JSON metadata file where the value of the `id` property is **ADD**.</span></span>

```js
/**
 * Add two numbers
 * @customfunction
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} The sum of the two numbers.
 */
function add(first, second) {
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

<span data-ttu-id="ad5fd-113">次の JSON は、以前のカスタム関数 JavaScript コードに関連付けられている JSON メタデータを示しています。</span><span class="sxs-lookup"><span data-stu-id="ad5fd-113">The following JSON shows the JSON metadata that is associated with the previous custom function JavaScript code.</span></span>

```json
{
  "functions": [
    {
        "description": "Add two numbers",
        "id": "ADD",
        "name": "ADD",
        "parameters": [
            {
                "description": "First number",
                "name": "first",
                "type": "number"
            },
            {
                "description": "Second number",
                "name": "second",
                "type": "number"
            }
        ],
        "result": {
            "type": "number"
        }
    },
  ]
}
```


<span data-ttu-id="ad5fd-114">JavaScript ファイルでカスタム関数を作成し、JSON のメタデータ ファイルに対応する情報を指定するときは、次のベスト プラクティスに留意してください。</span><span class="sxs-lookup"><span data-stu-id="ad5fd-114">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

* <span data-ttu-id="ad5fd-115">JSON のメタデータ ファイルにそれぞれの `id` プロパティには、英数字とピリオドのみが含まれています。</span><span class="sxs-lookup"><span data-stu-id="ad5fd-115">In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.</span></span>

* <span data-ttu-id="ad5fd-116">JSON のメタデータ ファイルで、各 `id` プロパティの値が、ファイルのスコープ内で一意であることを確認します。</span><span class="sxs-lookup"><span data-stu-id="ad5fd-116">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file.</span></span> <span data-ttu-id="ad5fd-117">すなわち、メタデータ ファイル内の 2 つの関数オブジェクトは同じ `id` 値であってはいけません。</span><span class="sxs-lookup"><span data-stu-id="ad5fd-117">That is, no two function objects in the metadata file should have the same `id` value.</span></span>

* <span data-ttu-id="ad5fd-118">対応する JavaScript 関数の名前に関連付けられた後では、JSON のメタデータ ファイル内の `id` プロパティの値を変更しないでください。</span><span class="sxs-lookup"><span data-stu-id="ad5fd-118">Do not change the value of an `id` property in the JSON metadata file after it's been associated with a corresponding JavaScript function name.</span></span> <span data-ttu-id="ad5fd-119">JSON のメタデータ ファイル内の `name` プロパティを更新することによって Excel でエンド ユーザーに表示される関数の名前を変更することができます。しかし、確立された後は、 `id` プロパティの値を決して変更しないでください。</span><span class="sxs-lookup"><span data-stu-id="ad5fd-119">You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

* <span data-ttu-id="ad5fd-120">JavaScript ファイルで、各関数の`CustomFunctions.associate`後に、カスタム関数の関連付けを指定します。</span><span class="sxs-lookup"><span data-stu-id="ad5fd-120">In the JavaScript file, specify a custom function association using `CustomFunctions.associate` after each function.</span></span>

<span data-ttu-id="ad5fd-121">次のサンプルは、JavaScript コード サンプルで定義された関数に対応する JSON メタデータを示します。</span><span class="sxs-lookup"><span data-stu-id="ad5fd-121">The following sample shows the JSON metadata that corresponds to the functions defined in this JavaScript code sample.</span></span> <span data-ttu-id="ad5fd-122">プロパティ`id`と`name`プロパティの値は、大文字で記述します。これは、カスタム関数を記述するときのベストプラクティスです。</span><span class="sxs-lookup"><span data-stu-id="ad5fd-122">The `id` and `name` property values are in uppercase, which is a best practice when describing your custom functions.</span></span> <span data-ttu-id="ad5fd-123">この JSON を追加する必要があるのは、自動生成を使用せずに、手動で独自の JSON ファイルを準備する場合だけです。</span><span class="sxs-lookup"><span data-stu-id="ad5fd-123">You only need to add this JSON if you are preparing your own JSON file manually and not using autogeneration.</span></span> <span data-ttu-id="ad5fd-124">Autogeneration の詳細については、「[カスタム関数の JSON メタデータを作成](custom-functions-json-autogeneration.md)する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="ad5fd-124">For more information on autogeneration, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

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

## <a name="additional-considerations"></a><span data-ttu-id="ad5fd-125">その他の考慮事項</span><span class="sxs-lookup"><span data-stu-id="ad5fd-125">Additional considerations</span></span>

<span data-ttu-id="ad5fd-126">カスタム関数から直接または間接的に (たとえば、jQuery を使用して) ドキュメントオブジェクトモデル (DOM) にアクセスしないようにします。</span><span class="sxs-lookup"><span data-stu-id="ad5fd-126">Avoid accessing the Document Object Model (DOM) directly or indirectly (for example, using jQuery) from your custom function.</span></span> <span data-ttu-id="ad5fd-127">カスタム関数が[JavaScript ランタイム](custom-functions-runtime.md)を使用する Windows 上の Excel では、カスタム関数は DOM にアクセスできません。</span><span class="sxs-lookup"><span data-stu-id="ad5fd-127">In Excel on Windows, where custom functions use the [JavaScript runtime](custom-functions-runtime.md), custom functions cannot access the DOM.</span></span>

## <a name="next-steps"></a><span data-ttu-id="ad5fd-128">次のステップ</span><span class="sxs-lookup"><span data-stu-id="ad5fd-128">Next steps</span></span>
<span data-ttu-id="ad5fd-129">[カスタム関数を使用して web 要求を実行](custom-functions-web-reqs.md)する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="ad5fd-129">Learn how to [perform web requests with custom functions](custom-functions-web-reqs.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="ad5fd-130">関連項目</span><span class="sxs-lookup"><span data-stu-id="ad5fd-130">See also</span></span>

* [<span data-ttu-id="ad5fd-131">カスタム関数用の JSON メタデータの自動生成</span><span class="sxs-lookup"><span data-stu-id="ad5fd-131">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="ad5fd-132">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="ad5fd-132">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="ad5fd-133">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="ad5fd-133">Create custom functions in Excel</span></span>](custom-functions-overview.md)
