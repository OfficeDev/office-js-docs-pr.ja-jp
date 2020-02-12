---
title: Office アドインにおける同一生成元ポリシーの制限への対処
description: ''
ms.date: 10/17/2019
localization_priority: Normal
ms.openlocfilehash: 2a47339bd5cc0b0bf919152b7078d5373382124f
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950447"
---
# <a name="addressing-same-origin-policy-limitations-in-office-add-ins"></a><span data-ttu-id="18545-102">Office アドインにおける同一生成元ポリシーの制限への対処</span><span class="sxs-lookup"><span data-stu-id="18545-102">Addressing same-origin policy limitations in Office Add-ins</span></span>

<span data-ttu-id="18545-p101">ブラウザーによって適用される同一生成元ポリシーでは、あるドメインから読み込まれたスクリプトで別のドメインの Web ページのプロパティを取得または操作できないようにしています。つまり、既定で、要求された URL のドメインは現在の Web ページのドメインと同じである必要があります。たとえば、このポリシーを適用すると、あるドメインの Web ページから、そのページがホストされているドメインとは別のドメインに対して [XmlHttpRequest](https://www.w3.org/TR/XMLHttpRequest/) Web サービスを呼び出せません。</span><span class="sxs-lookup"><span data-stu-id="18545-p101">The same-origin policy enforced by the browser prevents a script loaded from one domain from getting or manipulating properties of a webpage from another domain. This means that, by default, the domain of a requested URL must be the same as the domain of the current webpage. For example, this policy will prevent a webpage in one domain from making [XmlHttpRequest](https://www.w3.org/TR/XMLHttpRequest/) web-service calls to a domain other than the one where it is hosted.</span></span>

<span data-ttu-id="18545-106">Office アドインはブラウザー コントロールでホストされるので、それらの Web ページで実行されるスクリプトにも同一生成元ポリシーが適用されます。</span><span class="sxs-lookup"><span data-stu-id="18545-106">Because Office Add-ins are hosted in a browser control, the same-origin policy applies to script running in their web pages as well.</span></span>

<span data-ttu-id="18545-107">同一生成元ポリシーは、Web アプリケーションが複数のサブドメインに渡るコンテンツと API をホストしているときなど、多くの場合に不要な制約になることがあります。</span><span class="sxs-lookup"><span data-stu-id="18545-107">The same-origin policy can be an unnecessary handicap in many situations, such as when a web application hosts content and APIs across multiple subdomains.</span></span> <span data-ttu-id="18545-108">同一生成元ポリシーの適用に関する制約を安全に解消するための一般的な手法がいくつかあります。</span><span class="sxs-lookup"><span data-stu-id="18545-108">There are a few common techniques for securely overcoming same-origin policy enforcement.</span></span> <span data-ttu-id="18545-109">この記事では、その一部について簡単な紹介のみを示します。</span><span class="sxs-lookup"><span data-stu-id="18545-109">This article can only provide the briefest introduction to some of them.</span></span> <span data-ttu-id="18545-110">ここに示したリンクを使用して、こうした手法の調査を開始してください。</span><span class="sxs-lookup"><span data-stu-id="18545-110">Please use the links provided to get started in your research of these techniques.</span></span>

## <a name="use-jsonp-for-anonymous-access"></a><span data-ttu-id="18545-111">匿名アクセスに JSONP を使用する</span><span class="sxs-lookup"><span data-stu-id="18545-111">Use JSONP for anonymous access</span></span>

<span data-ttu-id="18545-112">同一生成元ポリシーの制限を解消する 1 つの方法として、[JSONP](https://www.w3schools.com/js/js_json_jsonp.asp) を使用して Web サービスのプロキシを提供します。</span><span class="sxs-lookup"><span data-stu-id="18545-112">One way to overcome same-origin policy limitations is to use [JSONP](https://www.w3schools.com/js/js_json_jsonp.asp) to provide a proxy for the web service.</span></span> <span data-ttu-id="18545-113">これを行うためには、任意のドメインでホストされているスクリプトを参照する `src` 属性を持つ `script` タグを使用します。</span><span class="sxs-lookup"><span data-stu-id="18545-113">You do this by including a `script` tag with a `src` attribute that points to some script hosted on any domain.</span></span> <span data-ttu-id="18545-114">`script` タグをプログラムで作成し、`src` 属性で参照する URL を動的に作成すると、URI クエリ パラメーターを介してパラメーターを URL に渡すことができます。</span><span class="sxs-lookup"><span data-stu-id="18545-114">You can programmatically create the `script` tags, dynamically create the URL to point the `src` attribute to, and then pass parameters to the URL via URI query parameters.</span></span> <span data-ttu-id="18545-115">Web サービス プロバイダーは、固有の URL で JavaScript コードを作成およびホストし、URI クエリ パラメーターに応じて異なるスクリプトを返します。</span><span class="sxs-lookup"><span data-stu-id="18545-115">Web service providers create and host JavaScript code at specific URLs, and return different scripts depending on the URI query parameters.</span></span> <span data-ttu-id="18545-116">それらのスクリプトは挿入された場所で実行され、想定どおりに動作します。</span><span class="sxs-lookup"><span data-stu-id="18545-116">These scripts then execute where they are inserted and work as expected.</span></span>

<span data-ttu-id="18545-117">次に、あらゆる Office アドインで機能する手法を使用する JSONP の例を示します。</span><span class="sxs-lookup"><span data-stu-id="18545-117">The following is an example of JSONP that uses a technique that will work in any Office Add-in.</span></span>

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


## <a name="implement-server-side-code-using-a-token-based-authorization-scheme"></a><span data-ttu-id="18545-118">トークン ベースの認証スキームを使用してサーバー側のコードを実装する</span><span class="sxs-lookup"><span data-stu-id="18545-118">Implement server-side code using a token-based authorization scheme</span></span>

<span data-ttu-id="18545-119">同一生成元ポリシーの制限に対処するもう 1 つの方法として、[OAuth 2.0](https://oauth.net/2/) フローを使用するサーバー側のコードを用意します。このコードによって、別のドメインでホストされているリソースへの許可されたアクセスを可能にします。</span><span class="sxs-lookup"><span data-stu-id="18545-119">Another way to address same-origin policy limitations is to provide server-side code that uses [OAuth 2.0](https://oauth.net/2/) flows to enable one domain to get authorized access to resources hosted on another.</span></span> 


## <a name="use-cross-origin-resource-sharing-cors"></a><span data-ttu-id="18545-120">クロス オリジン リソース共有 (CORS) を使用する</span><span class="sxs-lookup"><span data-stu-id="18545-120">Use cross-origin resource sharing (CORS)</span></span>


<span data-ttu-id="18545-121">[XmlHttpRequest2](https://dvcs.w3.org/hg/xhr/raw-file/tip/Overview.html) のクロス オリジン リソース共有機能を使用する例については、「[New Tricks in XMLHttpRequest2 に関する新しいヒント](https://www.html5rocks.com/en/tutorials/file/xhr2/)」の「Cross Origin Resource Sharing (CORS)」セクションを参照してください。</span><span class="sxs-lookup"><span data-stu-id="18545-121">For an example of using the cross-origin resource sharing feature of [XmlHttpRequest2](https://dvcs.w3.org/hg/xhr/raw-file/tip/Overview.html), see the "Cross Origin Resource Sharing (CORS)" section of [New Tricks in XMLHttpRequest2](https://www.html5rocks.com/en/tutorials/file/xhr2/).</span></span>


## <a name="build-your-own-proxy-using-iframe-and-post-message-cross-window-messaging"></a><span data-ttu-id="18545-122">IFRAME と POST MESSAGE を使用して独自のプロキシを作成する (クロス ウィンドウ メッセージング)</span><span class="sxs-lookup"><span data-stu-id="18545-122">Build your own proxy using IFRAME and POST MESSAGE (Cross-Window Messaging)</span></span>


<span data-ttu-id="18545-123">IFRAME および POST MESSAGE を使用して独自のプロキシを作成する例については、「[Cross-Window Messaging](http://ejohn.org/blog/cross-window-messaging/)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="18545-123">For an example of how to build your own proxy using IFRAME and POST MESSAGE, see [Cross-Window Messaging](http://ejohn.org/blog/cross-window-messaging/).</span></span>


## <a name="see-also"></a><span data-ttu-id="18545-124">関連項目</span><span class="sxs-lookup"><span data-stu-id="18545-124">See also</span></span>

- [<span data-ttu-id="18545-125">Office アドインのプライバシーとセキュリティ</span><span class="sxs-lookup"><span data-stu-id="18545-125">Privacy and security for Office Add-ins</span></span>](../concepts/privacy-and-security.md)
    
