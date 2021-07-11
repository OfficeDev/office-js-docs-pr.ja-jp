---
ms.date: 09/25/2020
description: 作業Excel特定の JavaScript ランタイムを使用しないカスタム関数について説明します。
title: UI レス のカスタム関数Excelランタイム
localization_priority: Normal
ms.openlocfilehash: aa2cf2632ddf9eb1ad1eb202b031ee2ca686af01
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349624"
---
# <a name="runtime-for-ui-less-excel-custom-functions"></a><span data-ttu-id="39ead-103">UI レス のカスタム関数Excelランタイム</span><span class="sxs-lookup"><span data-stu-id="39ead-103">Runtime for UI-less Excel custom functions</span></span>

<span data-ttu-id="39ead-104">作業ウィンドウを使用しないカスタム関数 (UI レスのカスタム関数) は、計算のパフォーマンスを最適化するように設計された JavaScript ランタイムを使用します。</span><span class="sxs-lookup"><span data-stu-id="39ead-104">Custom functions that don't use a task pane (UI-less custom functions) use a JavaScript runtime that is designed to optimize performance of calculations.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

<span data-ttu-id="39ead-105">この JavaScript ランタイムは、UI レスのカスタム関数と作業ウィンドウでデータを格納するために使用できる名前空間内の `OfficeRuntime` API へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="39ead-105">This JavaScript runtime provides access to APIs in the `OfficeRuntime` namespace that can be used by UI-less custom functions and the task pane to store data.</span></span>

## <a name="requesting-external-data"></a><span data-ttu-id="39ead-106">外部データの要求</span><span class="sxs-lookup"><span data-stu-id="39ead-106">Requesting external data</span></span>

<span data-ttu-id="39ead-107">UI レスのカスタム関数内では [、Fetch](https://developer.mozilla.org/docs/Web/API/Fetch_API) のような API を使用するか、サーバーとやり取りするための HTTP 要求を発行する標準 Web API [である XmlHttpRequest (XHR)](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest)を使用して外部データを要求できます。</span><span class="sxs-lookup"><span data-stu-id="39ead-107">Within a UI-less custom function, you can request external data by using an API like [Fetch](https://developer.mozilla.org/docs/Web/API/Fetch_API) or by using [XmlHttpRequest (XHR)](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span>

<span data-ttu-id="39ead-108">XMLHttpRequests を作成する場合は、UI レス関数で追加のセキュリティ対策を使用[](https://developer.mozilla.org/docs/Web/Security/Same-origin_policy)する必要があります。同じオリジン ポリシーと単純[な CORS](https://www.w3.org/TR/cors/)が必要です。</span><span class="sxs-lookup"><span data-stu-id="39ead-108">Be aware that UI-less functions must use additional security measures when making XmlHttpRequests, requiring [Same Origin Policy](https://developer.mozilla.org/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>

<span data-ttu-id="39ead-109">単純な CORS 実装では Cookie を使用できません。単純なメソッド (GET、HEAD、POST) のみをサポートします。</span><span class="sxs-lookup"><span data-stu-id="39ead-109">A simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST).</span></span> <span data-ttu-id="39ead-110">単純な CORS はフィールド名`Accept`、 `Accept-Language`、`Content-Language`の簡単なヘッダーを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="39ead-110">Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`.</span></span> <span data-ttu-id="39ead-111">コンテンツ タイプが 、 である場合は、単純な CORS でヘッダー `Content-Type` `application/x-www-form-urlencoded` `text/plain` を使用できます `multipart/form-data` 。</span><span class="sxs-lookup"><span data-stu-id="39ead-111">You can also use a `Content-Type` header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.</span></span>

## <a name="storing-and-accessing-data"></a><span data-ttu-id="39ead-112">データの格納およびアクセス</span><span class="sxs-lookup"><span data-stu-id="39ead-112">Storing and accessing data</span></span>

<span data-ttu-id="39ead-113">UI レスのカスタム関数内では、オブジェクトを使用してデータを格納およびアクセス `OfficeRuntime.storage` できます。</span><span class="sxs-lookup"><span data-stu-id="39ead-113">Within a UI-less custom function, you can store and access data by using the `OfficeRuntime.storage` object.</span></span> <span data-ttu-id="39ead-114">`Storage` は、UI レスのカスタム関数では使用できない [localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage)の代替手段を提供する、暗号化されていない永続的なキー値ストレージ システムです。</span><span class="sxs-lookup"><span data-stu-id="39ead-114">`Storage` is a persistent, unencrypted, key-value storage system that provides an alternative to [localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage), which cannot be used by UI-less custom functions.</span></span> <span data-ttu-id="39ead-115">`Storage` ドメインごとに 10 MB のデータを提供します。</span><span class="sxs-lookup"><span data-stu-id="39ead-115">`Storage` offers 10 MB of data per domain.</span></span> <span data-ttu-id="39ead-116">ドメインは、複数のアドインで共有できます。</span><span class="sxs-lookup"><span data-stu-id="39ead-116">Domains can be shared by more than one add-in.</span></span>

<span data-ttu-id="39ead-117">`Storage` は共有ストレージ ソリューションとして機能することを意図しています。つまり、アドインの複数の部分が同じデータにアクセスできるようになります。</span><span class="sxs-lookup"><span data-stu-id="39ead-117">`Storage` is intended as a shared storage solution, meaning multiple parts of an add-in are able to access the same data.</span></span> <span data-ttu-id="39ead-118">たとえば、ユーザー認証のトークンは、UI レスのカスタム関数と作業ウィンドウなどのアドイン UI 要素の両方からアクセスできるので、格納 `storage` できます。</span><span class="sxs-lookup"><span data-stu-id="39ead-118">For example, tokens for user authentication may be stored in `storage` because it can be accessed by both a UI-less custom function and add-in UI elements such as a task pane.</span></span> <span data-ttu-id="39ead-119">同様に、2 つのアドインが同じドメイン (たとえば、 ) を共有している場合、情報の前後 `www.contoso.com/addin1` `www.contoso.com/addin2` の共有も許可されます `storage` 。</span><span class="sxs-lookup"><span data-stu-id="39ead-119">Similarly, if two add-ins share the same domain (for example, `www.contoso.com/addin1`, `www.contoso.com/addin2`), they are also permitted to share information back and forth through `storage`.</span></span> <span data-ttu-id="39ead-120">異なるサブドメインを持つアドインは、(たとえば、 ) の異なるインスタンスを持つ点に `storage` `subdomain.contoso.com/addin1` 注意 `differentsubdomain.contoso.com/addin2` してください。</span><span class="sxs-lookup"><span data-stu-id="39ead-120">Note that add-ins which have different subdomains will have different instances of `storage` (for example, `subdomain.contoso.com/addin1`, `differentsubdomain.contoso.com/addin2`).</span></span>

<span data-ttu-id="39ead-121">`storage` は共有の場所として機能することから、キー値の組み合わせが書き換えられる可能性があることにご注意ください。</span><span class="sxs-lookup"><span data-stu-id="39ead-121">Because `storage` can be a shared location, it is important to realize that it is possible to override key-value pairs.</span></span>

<span data-ttu-id="39ead-122">オブジェクトでは、次のメソッドを使用 `storage` できます。</span><span class="sxs-lookup"><span data-stu-id="39ead-122">The following methods are available on the `storage` object.</span></span>

- `getItem`
- `getItems`
- `setItem`
- `setItems`
- `removeItem`
- `removeItems`
- `getKeys`

> [!NOTE]
> <span data-ttu-id="39ead-123">すべての情報 (など) をクリアする方法はありません `clear` 。</span><span class="sxs-lookup"><span data-stu-id="39ead-123">There's no method for clearing all information (such as `clear`).</span></span> <span data-ttu-id="39ead-124">代わりに、一度に複数のエントリを削除できる `removeItems` を使用してください。</span><span class="sxs-lookup"><span data-stu-id="39ead-124">Instead, you should instead use `removeItems` to remove multiple entries at a time.</span></span>

### <a name="officeruntimestorage-example"></a><span data-ttu-id="39ead-125">OfficeRuntime.storage の例</span><span class="sxs-lookup"><span data-stu-id="39ead-125">OfficeRuntime.storage example</span></span>

<span data-ttu-id="39ead-126">次のコード サンプルでは、キー `OfficeRuntime.storage.setItem` と値をに設定する関数を呼び出します `storage` 。</span><span class="sxs-lookup"><span data-stu-id="39ead-126">The following code sample calls the `OfficeRuntime.storage.setItem` function to set a key and value into `storage`.</span></span>

```js
function StoreValue(key, value) {

  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}
```

## <a name="additional-considerations"></a><span data-ttu-id="39ead-127">その他の考慮事項</span><span class="sxs-lookup"><span data-stu-id="39ead-127">Additional considerations</span></span>

<span data-ttu-id="39ead-128">アドインで UI レスのカスタム関数のみを使用する場合は、UI レスのカスタム関数を使用してドキュメント オブジェクト モデル (DOM) にアクセスしたり、DOM に依存する jQuery のようなライブラリを使用したりすることはできません。</span><span class="sxs-lookup"><span data-stu-id="39ead-128">If your add-in only uses UI-less custom functions, note that you can't access the Document Object Model (DOM) with UI-less custom functions or use libraries like jQuery that rely on the DOM.</span></span>

## <a name="next-steps"></a><span data-ttu-id="39ead-129">次の手順</span><span class="sxs-lookup"><span data-stu-id="39ead-129">Next steps</span></span>
<span data-ttu-id="39ead-130">UI レスの [カスタム関数をデバッグする方法について説明します](custom-functions-debugging.md)。</span><span class="sxs-lookup"><span data-stu-id="39ead-130">Learn how to [debug UI-less custom functions](custom-functions-debugging.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="39ead-131">関連項目</span><span class="sxs-lookup"><span data-stu-id="39ead-131">See also</span></span>

* [<span data-ttu-id="39ead-132">UI レスのカスタム関数を認証する</span><span class="sxs-lookup"><span data-stu-id="39ead-132">Authenticate UI-less custom functions</span></span>](custom-functions-authentication.md)
* [<span data-ttu-id="39ead-133">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="39ead-133">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="39ead-134">カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="39ead-134">Custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
