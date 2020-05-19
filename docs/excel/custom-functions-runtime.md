---
ms.date: 05/17/2020
description: 作業ウィンドウおよび特定の JavaScript ランタイムを使用しない Excel カスタム関数について説明します。
title: UI レス Excel カスタム関数のランタイム
localization_priority: Normal
ms.openlocfilehash: 31044d4569d230e252c05a39785fc7d47b802e37
ms.sourcegitcommit: f62d9630de69c5c070e3d4048205f5cc654db7e4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/18/2020
ms.locfileid: "44278358"
---
# <a name="runtime-for-ui-less-excel-custom-functions"></a><span data-ttu-id="f194b-103">UI レス Excel カスタム関数のランタイム</span><span class="sxs-lookup"><span data-stu-id="f194b-103">Runtime for UI-less Excel custom functions</span></span>

<span data-ttu-id="f194b-104">作業ウィンドウを使用しないカスタム関数 (UI レスカスタム関数) は、計算のパフォーマンスを最適化するように設計された JavaScript ランタイムを使用します。</span><span class="sxs-lookup"><span data-stu-id="f194b-104">Custom functions that don't use a task pane (UI-less custom functions) use a JavaScript runtime that is designed to optimize performance of calculations.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

<span data-ttu-id="f194b-105">この JavaScript ランタイムは、UI を使用しない `OfficeRuntime` カスタム関数と作業ウィンドウでデータを格納するために使用できる名前空間の api へのアクセスを提供します。</span><span class="sxs-lookup"><span data-stu-id="f194b-105">This JavaScript runtime provides access to APIs in the `OfficeRuntime` namespace that can be used by UI-less custom functions and the task pane to store data.</span></span>

## <a name="requesting-external-data"></a><span data-ttu-id="f194b-106">外部データの要求</span><span class="sxs-lookup"><span data-stu-id="f194b-106">Requesting external data</span></span>

<span data-ttu-id="f194b-107">UI を使用しないカスタム関数内では、サーバーと対話するために HTTP 要求を発行する標準の web API である[Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API)や、 [XmlHttpRequest (xhr)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest)などの API を使用して外部データを要求できます。</span><span class="sxs-lookup"><span data-stu-id="f194b-107">Within a UI-less custom function, you can request external data by using an API like [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) or by using [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span>

<span data-ttu-id="f194b-108">UI を使用しない関数では、XmlHttpRequests を作成するときに追加のセキュリティ対策を使用する必要があることに注意してください。[元のポリシー](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy)と単純な[CORS](https://www.w3.org/TR/cors/)が必要です。</span><span class="sxs-lookup"><span data-stu-id="f194b-108">Be aware that UI-less functions must use additional security measures when making XmlHttpRequests, requiring [Same Origin Policy](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>

<span data-ttu-id="f194b-109">単純な CORS 実装は cookie を使用できず、simple メソッド (GET、HEAD、POST) のみをサポートしています。</span><span class="sxs-lookup"><span data-stu-id="f194b-109">A simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST).</span></span> <span data-ttu-id="f194b-110">単純な CORS はフィールド名`Accept`、 `Accept-Language`、`Content-Language`の簡単なヘッダーを受け入れます。</span><span class="sxs-lookup"><span data-stu-id="f194b-110">Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`.</span></span> <span data-ttu-id="f194b-111">`Content-Type`コンテンツタイプが、、またはの場合は、単純な CORS のヘッダーを使用することもでき `application/x-www-form-urlencoded` `text/plain` `multipart/form-data` ます。</span><span class="sxs-lookup"><span data-stu-id="f194b-111">You can also use a `Content-Type` header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.</span></span>

## <a name="storing-and-accessing-data"></a><span data-ttu-id="f194b-112">データの格納およびアクセス</span><span class="sxs-lookup"><span data-stu-id="f194b-112">Storing and accessing data</span></span>

<span data-ttu-id="f194b-113">UI を使用しないカスタム関数では、オブジェクトを使用してデータを格納したり、データにアクセスしたりでき `OfficeRuntime.storage` ます。</span><span class="sxs-lookup"><span data-stu-id="f194b-113">Within a UI-less custom function, you can store and access data by using the `OfficeRuntime.storage` object.</span></span> <span data-ttu-id="f194b-114">`Storage`は、暗号化されていない、暗号化されていないキー値を持つ、永続的なストレージシステムです。これ[は、UI には](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage)ないカスタム関数では使用できません。</span><span class="sxs-lookup"><span data-stu-id="f194b-114">`Storage` is a persistent, unencrypted, key-value storage system that provides an alternative to [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), which cannot be used by UI-less custom functions.</span></span> <span data-ttu-id="f194b-115">`Storage`ドメインごとに 10 MB のデータを提供します。</span><span class="sxs-lookup"><span data-stu-id="f194b-115">`Storage` offers 10 MB of data per domain.</span></span> <span data-ttu-id="f194b-116">ドメインは複数のアドインで共有できます。</span><span class="sxs-lookup"><span data-stu-id="f194b-116">Domains can be shared by more than one add-in.</span></span>

<span data-ttu-id="f194b-117">`Storage` は共有ストレージ ソリューションとして機能することを意図しています。つまり、アドインの複数の部分が同じデータにアクセスできるようになります。</span><span class="sxs-lookup"><span data-stu-id="f194b-117">`Storage` is intended as a shared storage solution, meaning multiple parts of an add-in are able to access the same data.</span></span> <span data-ttu-id="f194b-118">たとえば、ユーザー認証のトークンは、 `storage` UI なしのカスタム関数と、作業ウィンドウなどのアドインの ui 要素の両方からアクセスできるため、に格納されます。</span><span class="sxs-lookup"><span data-stu-id="f194b-118">For example, tokens for user authentication may be stored in `storage` because it can be accessed by both a UI-less custom function and add-in UI elements such as a task pane.</span></span> <span data-ttu-id="f194b-119">同様に、2つのアドインが同じドメイン (たとえば、など) を共有している場合は、 `www.contoso.com/addin1` `www.contoso.com/addin2` 情報を相互間で共有することもでき `storage` ます。</span><span class="sxs-lookup"><span data-stu-id="f194b-119">Similarly, if two add-ins share the same domain (for example, `www.contoso.com/addin1`, `www.contoso.com/addin2`), they are also permitted to share information back and forth through `storage`.</span></span> <span data-ttu-id="f194b-120">サブドメインが異なるアドインは、のインスタンスが異なることに注意 `storage` してください (例: `subdomain.contoso.com/addin1` `differentsubdomain.contoso.com/addin2` )。</span><span class="sxs-lookup"><span data-stu-id="f194b-120">Note that add-ins which have different subdomains will have different instances of `storage` (for example, `subdomain.contoso.com/addin1`, `differentsubdomain.contoso.com/addin2`).</span></span>

<span data-ttu-id="f194b-121">`storage` は共有の場所として機能することから、キー値の組み合わせが書き換えられる可能性があることにご注意ください。</span><span class="sxs-lookup"><span data-stu-id="f194b-121">Because `storage` can be a shared location, it is important to realize that it is possible to override key-value pairs.</span></span>

<span data-ttu-id="f194b-122">`storage` オブジェクトでは、以下の方法が利用可能です。</span><span class="sxs-lookup"><span data-stu-id="f194b-122">The following methods are available on the `storage` object:</span></span>

 - `getItem`
 - `getItems`
 - `setItem`
 - `setItems`
 - `removeItem`
 - `removeItems`
 - `getKeys`

<span data-ttu-id="f194b-123">.</span><span class="sxs-lookup"><span data-stu-id="f194b-123">.</span></span>[!NOTE]
> <span data-ttu-id="f194b-124">すべての情報 (など) を消去する方法はありません `clear` 。</span><span class="sxs-lookup"><span data-stu-id="f194b-124">There's no method for clearing all information (such as `clear`).</span></span> <span data-ttu-id="f194b-125">代わりに、一度に複数のエントリを削除できる `removeItems` を使用してください。</span><span class="sxs-lookup"><span data-stu-id="f194b-125">Instead, you should instead use `removeItems` to remove multiple entries at a time.</span></span>

### <a name="officeruntimestorage-example"></a><span data-ttu-id="f194b-126">一例</span><span class="sxs-lookup"><span data-stu-id="f194b-126">OfficeRuntime.storage example</span></span>

<span data-ttu-id="f194b-127">次のコードサンプルでは、関数を呼び出して `OfficeRuntime.storage.setItem` キーと値をに設定 `storage` します。</span><span class="sxs-lookup"><span data-stu-id="f194b-127">The following code sample calls the `OfficeRuntime.storage.setItem` function to set a key and value into `storage`.</span></span>

```js
function StoreValue(key, value) {

  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}
```

## <a name="additional-considerations"></a><span data-ttu-id="f194b-128">その他の考慮事項</span><span class="sxs-lookup"><span data-stu-id="f194b-128">Additional considerations</span></span>

<span data-ttu-id="f194b-129">アドインで UI を使用しないカスタム関数のみが使用されている場合は、UI を使用しないカスタム関数を使用してドキュメントオブジェクトモデル (DOM) にアクセスしたり、DOM に依存している jQuery などのライブラリを使用したりすることができないことに注意してください。</span><span class="sxs-lookup"><span data-stu-id="f194b-129">If your add-in only uses UI-less custom functions, note that you can't access the Document Object Model (DOM) with UI-less custom functions or use libraries like jQuery that rely on the DOM.</span></span>

## <a name="next-steps"></a><span data-ttu-id="f194b-130">次の手順</span><span class="sxs-lookup"><span data-stu-id="f194b-130">Next steps</span></span>
<span data-ttu-id="f194b-131">[UI のないカスタム関数をデバッグ](custom-functions-debugging.md)する方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="f194b-131">Learn how to [debug UI-less custom functions](custom-functions-debugging.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="f194b-132">関連項目</span><span class="sxs-lookup"><span data-stu-id="f194b-132">See also</span></span>

* [<span data-ttu-id="f194b-133">UI レスのカスタム関数を認証する</span><span class="sxs-lookup"><span data-stu-id="f194b-133">Authenticate UI-less custom functions</span></span>](custom-functions-authentication.md)
* [<span data-ttu-id="f194b-134">Excel でカスタム関数を作成する</span><span class="sxs-lookup"><span data-stu-id="f194b-134">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="f194b-135">カスタム関数のチュートリアル</span><span class="sxs-lookup"><span data-stu-id="f194b-135">Custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
