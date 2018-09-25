---
title: Office アドインにおける同一生成元ポリシーの制限への対処
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 054a01d554c529579917218361bcb8aeebb04c3c
ms.sourcegitcommit: eb74e94d3e1bc1930a9c6582a0a99355d0da34f2
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/25/2018
ms.locfileid: "25004883"
---
# <a name="addressing-same-origin-policy-limitations-in-office-add-ins"></a><span data-ttu-id="96466-102">Office アドインにおける同一生成元ポリシーの制限への対処</span><span class="sxs-lookup"><span data-stu-id="96466-102">Addressing same-origin policy limitations in Office Add-ins</span></span>


<span data-ttu-id="96466-p101">ブラウザーによって適用される同一生成元ポリシーでは、あるドメインから読み込まれたスクリプトで別のドメインの Web ページのプロパティを取得または操作できないようにしています。つまり、既定で、要求された URL のドメインは現在の Web ページのドメインと同じである必要があります。たとえば、このポリシーを適用すると、あるドメインの Web ページから、そのページがホストされているドメインとは別のドメインに対して [XmlHttpRequest](https://www.w3.org/TR/XMLHttpRequest/) Web サービスを呼び出せません。</span><span class="sxs-lookup"><span data-stu-id="96466-p101">The same-origin policy enforced by the browser prevents a script loaded from one domain from getting or manipulating properties of a webpage from another domain. This means that, by default, the domain of a requested URL must be the same as the domain of the current webpage. For example, this policy will prevent a webpage in one domain from making [XmlHttpRequest](https://www.w3.org/TR/XMLHttpRequest/) web-service calls to a domain other than the one where it is hosted.</span></span>

<span data-ttu-id="96466-106">Office アドインはブラウザー コントロールでホストされるので、それらの Web ページで実行されるスクリプトにも同一生成元ポリシーが適用されます。</span><span class="sxs-lookup"><span data-stu-id="96466-106">Because Office Add-ins are hosted in a browser control, the same-origin policy applies to script running in their web pages as well.</span></span>

<span data-ttu-id="96466-107">アドインを開発する際に、同一生成元ポリシーの適用に対処するには、次のようにします。</span><span class="sxs-lookup"><span data-stu-id="96466-107">To overcome same-origin policy enforcement when you develop add-ins, you can:</span></span>

- <span data-ttu-id="96466-108">JSON/P を使用して匿名アクセスする。</span><span class="sxs-lookup"><span data-stu-id="96466-108">Use JSON/P for anonymous access.</span></span> 
    
- <span data-ttu-id="96466-109">トークン ベースの認証スキームを使用してサーバーサイド スクリプトを実装する。</span><span class="sxs-lookup"><span data-stu-id="96466-109">Implement server-side script using a token-based authentication scheme.</span></span>
    
- <span data-ttu-id="96466-110">クロス オリジン リソース共有 (CORS) を使用する。</span><span class="sxs-lookup"><span data-stu-id="96466-110">Using cross-origin resource sharing (CORS).</span></span>
    
- <span data-ttu-id="96466-111">IFRAME および POST MESSAGE を使用して独自のプロキシを作成する。</span><span class="sxs-lookup"><span data-stu-id="96466-111">Build your own proxy using IFRAME and POST MESSAGE.</span></span>
    

## <a name="using-jsonp-for-anonymous-access"></a><span data-ttu-id="96466-112">JSON/P を使用した匿名アクセス</span><span class="sxs-lookup"><span data-stu-id="96466-112">Using JSON/P for anonymous access</span></span>


<span data-ttu-id="96466-p102">この制限に対処する方法の 1 つは、JSON/P を使用して Web サービスのプロキシを提供することです。これを行うためには、任意のドメインでホストされているスクリプトを参照する `src` 属性を持つ `script` タグを使用します。`script` タグをプログラムで作成し、`src` 属性で参照する URL を動的に作成すると、URI クエリ パラメーターを介してパラメーターを URL に渡すことができます。Web サービス プロバイダーは、固有の URL で JavaScript コードを作成およびホストし、URI クエリ パラメーターに応じて異なるスクリプトを返します。それらのスクリプトは挿入された場所で実行され、想定どおりに動作します。</span><span class="sxs-lookup"><span data-stu-id="96466-p102">One way to overcome this limitation is to use JSON/P to provide a proxy for the web service. You do this by including a `script` tag with a `src` attribute that points to some script hosted on any domain. You can programmatically create the `script` tags, dynamically create the URL to point the `src` attribute to, and then pass parameters to the URL via URI query parameters. Web service providers create and host JavaScript code at specific URLs, and return different scripts depending on the URI query parameters. These scripts then execute where they are inserted and work as expected.</span></span>

<span data-ttu-id="96466-118">いずれの Office アドインでも機能する手法を使用する JSON/P の例を以下に示します。</span><span class="sxs-lookup"><span data-stu-id="96466-118">The following is an example of JSON/P that uses a technique that will work in any Office Add-in.</span></span>

```js
// Dynamically create an HTML SCRIPT element that obtains the details for the specified video.
function loadVideoDetails(videoIndex) {
    // Dynamically create a new HTML SCRIPT element in the webpage.
    var script = document.createElement("script");
    // Specify the URL to retrieve the indicated video from a feed of a current list of videos,
    // as the value of the src attribute of the SCRIPT element. 
    script.setAttribute("src", "https://gdata.youtube.com/feeds/api/videos/" + 
        videos[videoIndex].Id + "?alt=json-in-script&amp;callback=videoDetailsLoaded");
    // Insert the SCRIPT element at the end of the HEAD section.
    document.getElementsByTagName('head')[0].appendChild(script);
}

```


## <a name="implementing-server-side-script-using-a-token-based-authentication-scheme"></a><span data-ttu-id="96466-119">トークン ベースの認証スキームを使用するサーバーサイド スクリプトの実装</span><span class="sxs-lookup"><span data-stu-id="96466-119">Implementing server-side script using a token-based authentication scheme</span></span>


<span data-ttu-id="96466-120">同一生成元ポリシーの制限に対処する他の方法として、OAuth を使用する ASP ページ、または Cookie で資格情報をキャッシュする ASP ページとして、アドインの Web ページを実装する方法があります。</span><span class="sxs-lookup"><span data-stu-id="96466-120">Another way to address same-origin policy limitations is to implement the add-in's webpage as an ASP page that uses OAuth or caches credentials in cookies.</span></span>

<span data-ttu-id="96466-121">
  `System.Net` の `Cookie\` オブジェクトを使用して、Cookie の値を取得および設定する方法を示すサーバー側のコード例については、[Value](https://docs.microsoft.com/dotnet/api/system.net.cookie.value?view=netframework-4.7.2) プロパティを参照してください。</span><span class="sxs-lookup"><span data-stu-id="96466-121">For an example of server-side code that shows how to use the  `Cookie` object in `System.Net` to get and set cookie values, see the [Value](https://docs.microsoft.com/dotnet/api/system.net.cookie.value?view=netframework-4.7.2) property.</span></span>


## <a name="using-cross-origin-resource-sharing-cors"></a><span data-ttu-id="96466-122">クロス オリジン リソース共有 (CORS) の使用</span><span class="sxs-lookup"><span data-stu-id="96466-122">Using cross-origin resource sharing (CORS)</span></span>


<span data-ttu-id="96466-123">[XmlHttpRequest2](http://dvcs.w3.org/hg/xhr/raw-file/tip/Overview.html) のクロス オリジン リソース共有機能を使用する例については、「[New Tricks in XMLHttpRequest2 に関する新しいヒント](http://www.html5rocks.com/en/tutorials/file/xhr2/)」の「Cross Origin Resource Sharing (CORS)」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="96466-123">For an example of using the cross-origin resource sharing feature of [XmlHttpRequest2](http://dvcs.w3.org/hg/xhr/raw-file/tip/Overview.html), see the "Cross Origin Resource Sharing (CORS)" section of [New Tricks in XMLHttpRequest2](http://www.html5rocks.com/en/tutorials/file/xhr2/).</span></span>


## <a name="building-your-own-proxy-using-iframe-and-post-message"></a><span data-ttu-id="96466-124">IFRAME および POST MESSAGE を使用する独自のプロキシの作成</span><span class="sxs-lookup"><span data-stu-id="96466-124">Building your own proxy using IFRAME and POST MESSAGE</span></span>


<span data-ttu-id="96466-125">IFRAME および POST MESSAGE を使用して独自のプロキシを作成する例については、「[Cross-Window Messaging](http://ejohn.org/blog/cross-window-messaging/)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="96466-125">For an example of how to build your own proxy using IFRAME and POST MESSAGE, see [Cross-Window Messaging](http://ejohn.org/blog/cross-window-messaging/).</span></span>


## <a name="see-also"></a><span data-ttu-id="96466-126">関連項目</span><span class="sxs-lookup"><span data-stu-id="96466-126">See also</span></span>

- [<span data-ttu-id="96466-127">Office アドインのプライバシーとセキュリティ</span><span class="sxs-lookup"><span data-stu-id="96466-127">Privacy and security for Office Add-ins</span></span>](../concepts/privacy-and-security.md)
    
