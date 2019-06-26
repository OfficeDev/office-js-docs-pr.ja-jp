---
ms.date: 06/20/2019
description: '`OfficeRuntime.storage`を使用し、カスタム関数で状態を保存します。'
title: カスタム関数で状態を保存して共有する
localization_priority: Priority
ms.openlocfilehash: c6689393e5d118c779b7b261b0de04ead56aff83
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127836"
---
# <a name="save-and-share-state-in-custom-functions"></a><span data-ttu-id="0d8cb-103">カスタム関数で状態を保存して共有する</span><span class="sxs-lookup"><span data-stu-id="0d8cb-103">Save and share state in custom functions</span></span>

<span data-ttu-id="0d8cb-104">`OfficeRuntime.storage`オブジェクトを使用し、カスタム関数またはアドインの作業ウィンドウに関連した状態を保存します。</span><span class="sxs-lookup"><span data-stu-id="0d8cb-104">Use the `OfficeRuntime.storage` object to save state related to custom functions or the task pane in your add-in.</span></span> <span data-ttu-id="0d8cb-105">ストレージはドメイン 1 つにつき 10 MB に制限されています (複数のアドインで共有される可能性があります)。</span><span class="sxs-lookup"><span data-stu-id="0d8cb-105">Storage is limited to 10 MB per domain (which may be shared across multiple add-ins).</span></span> <span data-ttu-id="0d8cb-106">Excel on Windows では、`storage` オブジェクトはカスタム関数ランタイムの範囲内の別の場所になりますが、Excel on the web と Mac では、`storage` オブジェクトはブラウザーの `localStorage` と同じです。</span><span class="sxs-lookup"><span data-stu-id="0d8cb-106">In Excel on Windows, the `storage` object is a separate location within the custom functions runtime, but for Excel Online and Excel for Mac, the `storage` object is the same as the browser's `localStorage`.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="0d8cb-107">状態管理に`storage`を使用する方法は複数あります。</span><span class="sxs-lookup"><span data-stu-id="0d8cb-107">There are multiple ways to use `storage` for state management:</span></span>

- <span data-ttu-id="0d8cb-108">オフラインで Web リソースにアクセスできない時でも、カスタム関数を使用するための既定値を格納できます。</span><span class="sxs-lookup"><span data-stu-id="0d8cb-108">You can store default values for custom functions to use when you are offline and unable to reach a web resource.</span></span>
- <span data-ttu-id="0d8cb-109">Web リソースへの追加の呼び出しを回避するために使用するカスタム関数の値を保存できます。</span><span class="sxs-lookup"><span data-stu-id="0d8cb-109">You can save values for custom functions to use to avoid making additional calls to a web resource.</span></span>
- <span data-ttu-id="0d8cb-110">カスタム関数の値を保存できます。</span><span class="sxs-lookup"><span data-stu-id="0d8cb-110">You can save values from your custom function.</span></span>
- <span data-ttu-id="0d8cb-111">作業ウィンドウの値を格納できます。</span><span class="sxs-lookup"><span data-stu-id="0d8cb-111">You can store values from your task pane.</span></span>

<span data-ttu-id="0d8cb-112">次のコード サンプルでは、`storage`に項目を格納してそれを取得する方法を示します。</span><span class="sxs-lookup"><span data-stu-id="0d8cb-112">The following code sample illustrates how to store an item into `storage` and retrieve it.</span></span>

```js
function storeValue(key, value) {
  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}

function GetValue(key) {
  return OfficeRuntime.storage.getItem(key);
}

CustomFunctions.associate("STOREVALUE", StoreValue);
CustomFunctions.associate("GETVALUE", GetValue);
```

<span data-ttu-id="0d8cb-113">[GitHub 上の詳細なコードサンプル](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage)では、作業ウィンドウに上記の情報を受け渡す例を紹介しています。</span><span class="sxs-lookup"><span data-stu-id="0d8cb-113">[A more detailed code sample on GitHub](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage) gives an example of passing this information to the task pane.</span></span>

>[!NOTE]
> <span data-ttu-id="0d8cb-114">`storage`オブジェクトは、現在は推奨されていないところの`AsyncStorage`と名付けられた以前のストレージ オブジェクトの代わりとなります。</span><span class="sxs-lookup"><span data-stu-id="0d8cb-114">The `storage` object replaces the previous storage object named `AsyncStorage` which is now deprecated.</span></span> <span data-ttu-id="0d8cb-115">現行のカスタム関数コードで`AsyncStorage`オブジェクトを使用している場合は、それを更新して`storage`オブジェクトを使用してください。</span><span class="sxs-lookup"><span data-stu-id="0d8cb-115">If using the `AsyncStorage` object in your current custom functions code, please update it to use the `storage` object.</span></span>

## <a name="next-steps"></a><span data-ttu-id="0d8cb-116">次の手順</span><span class="sxs-lookup"><span data-stu-id="0d8cb-116">Next steps</span></span>
<span data-ttu-id="0d8cb-117">[カスタム関数の JSON メタデータを自動生成する](custom-functions-json-autogeneration.md)方法を学びます。</span><span class="sxs-lookup"><span data-stu-id="0d8cb-117">Learn how to [autogenerate the JSON metadata for your custom functions](custom-functions-json-autogeneration.md).</span></span> 

## <a name="see-also"></a><span data-ttu-id="0d8cb-118">関連項目</span><span class="sxs-lookup"><span data-stu-id="0d8cb-118">See also</span></span>

* [<span data-ttu-id="0d8cb-119">カスタム関数のメタデータ</span><span class="sxs-lookup"><span data-stu-id="0d8cb-119">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="0d8cb-120">Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="0d8cb-120">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="0d8cb-121">カスタム関数のベスト プラクティス</span><span class="sxs-lookup"><span data-stu-id="0d8cb-121">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="0d8cb-122">Excel カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="0d8cb-122">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="0d8cb-123">カスタム関数のデバッグ</span><span class="sxs-lookup"><span data-stu-id="0d8cb-123">Custom functions debugging</span></span>](custom-functions-debugging.md)
