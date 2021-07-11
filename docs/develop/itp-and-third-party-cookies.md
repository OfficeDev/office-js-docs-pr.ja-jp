---
title: サードパーティ cookie をOffice ITP で動作するアドインを開発する
description: サードパーティ Cookie を使用する場合Office ITP とアドインを使用する方法
ms.date: 03/12/2021
localization_priority: Normal
ms.openlocfilehash: dbc23e4ead0abc94ffa173ffc22919342c4fca6d
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349862"
---
# <a name="develop-your-office-add-in-to-work-with-itp-when-using-third-party-cookies"></a><span data-ttu-id="fbe70-103">サードパーティ cookie をOffice ITP で動作するアドインを開発する</span><span class="sxs-lookup"><span data-stu-id="fbe70-103">Develop your Office Add-in to work with ITP when using third-party cookies</span></span>

<span data-ttu-id="fbe70-104">カスタム アドインOfficeサード パーティ Cookie が必要な場合、アドインを読み込んだブラウザー ランタイムによってインテリジェント 追跡防止 (ITP) が使用されている場合、これらの Cookie はブロックされます。</span><span class="sxs-lookup"><span data-stu-id="fbe70-104">If your Office Add-in requires third-party cookies, those cookies are blocked if Intelligent Tracking Prevention (ITP) is used by the browser runtime that loaded your add-in.</span></span> <span data-ttu-id="fbe70-105">サードパーティの Cookie を使用してユーザーを認証したり、設定の保存などの他のシナリオで使用している場合があります。</span><span class="sxs-lookup"><span data-stu-id="fbe70-105">You may be using third-party cookies to authenticate users, or for other scenarios, such as storing settings.</span></span>

<span data-ttu-id="fbe70-106">アドインとOfficeがサードパーティの Cookie に依存している必要のある場合は、次の手順を使用して ITP を使用します。</span><span class="sxs-lookup"><span data-stu-id="fbe70-106">If your Office Add-in and website must rely on third-party cookies, use the following steps to work with ITP:</span></span>

1. <span data-ttu-id="fbe70-107">OAuth [2.0 Authorization](https://tools.ietf.org/html/rfc6749)を設定して、認証ドメイン (Cookie を要求するサード パーティ) が承認トークンを Web サイト   に転送します。</span><span class="sxs-lookup"><span data-stu-id="fbe70-107">Set up [OAuth 2.0 Authorization](https://tools.ietf.org/html/rfc6749) so that the authenticating domain (in your case, the third-party that expects cookies) forwards an authorization token to your website.</span></span> <span data-ttu-id="fbe70-108">トークンを使用して、サーバーセットの Secure Cookie と HttpOnly Cookie を使用してファースト パーティのログイン [セッションを確立します](https://developer.mozilla.org/docs/Web/HTTP/Cookies#Secure_and_HttpOnly_cookies)。</span><span class="sxs-lookup"><span data-stu-id="fbe70-108">Use the token to establish a first-party login session with a server-set Secure and [HttpOnly cookie](https://developer.mozilla.org/docs/Web/HTTP/Cookies#Secure_and_HttpOnly_cookies).</span></span>
2. <span data-ttu-id="fbe70-109">サードパーティが[ファーストStorage Cookie](https://webkit.org/blog/8124/introducing-storage-access-api/)へのアクセスを取得するためのアクセス許可を要求するには、Storage Access API   を使用します。</span><span class="sxs-lookup"><span data-stu-id="fbe70-109">Use the [Storage Access API](https://webkit.org/blog/8124/introducing-storage-access-api/) so that the third-party can request permission to get access to its first-party cookies.</span></span> <span data-ttu-id="fbe70-110">Mac 上の現在OfficeバージョンとOffice on the web API がサポートされています。</span><span class="sxs-lookup"><span data-stu-id="fbe70-110">Current versions of Office on Mac and Office on the web both support this API.</span></span>
    > [!NOTE]
    > <span data-ttu-id="fbe70-111">認証以外の目的で Cookie を使用している場合は、シナリオでの使用 `localStorage` を検討してください。</span><span class="sxs-lookup"><span data-stu-id="fbe70-111">If you're using cookies for purposes other than authentication, then consider using `localStorage` for your scenario.</span></span>

<span data-ttu-id="fbe70-112">次のコード サンプルは、Access API のStorage示しています。</span><span class="sxs-lookup"><span data-stu-id="fbe70-112">The following code sample shows how to use the Storage Access API.</span></span>

```javascript
function displayLoginButton() {
  var button = createLoginButton();
  button.addEventListener("click", function(ev) {
    document.requestStorageAccess().then(function() {
      authenticateWithCookies(); 
    }).catch(function() {
      // User must have previously interacted with this domain loaded in a top frame
      // Also you should have previously written a cookie when domain was loaded in the top frame
      console.error("User cancelled or requirements were not met.");
    });
  });
}

if (document.hasStorageAccess) { 
  document.hasStorageAccess().then(function(hasStorageAccess) { 
    if (!hasStorageAccess) { 
      displayLoginButton(); 
    } else { 
      authenticateWithCookies(); 
    } 
  }); 
} else { 
    authenticateWithCookies(); 
} 
```

## <a name="about-itp-and-third-party-cookies"></a><span data-ttu-id="fbe70-113">ITP およびサード パーティの Cookie について</span><span class="sxs-lookup"><span data-stu-id="fbe70-113">About ITP and third-party cookies</span></span>

<span data-ttu-id="fbe70-114">サード パーティ Cookie は、ドメインがトップ レベル のフレームとは異なる iframe に読み込まれる Cookie です。</span><span class="sxs-lookup"><span data-stu-id="fbe70-114">Third-party cookies are cookies that are loaded in an iframe, where the domain is different from the top level frame.</span></span> <span data-ttu-id="fbe70-115">ITP は複雑な認証シナリオに影響を与える可能性があります。ポップアップ ダイアログを使用して資格情報を入力し、認証フローを完了するためにアドイン iframe によって Cookie アクセスが必要になります。</span><span class="sxs-lookup"><span data-stu-id="fbe70-115">ITP could affect complex authentication scenarios, where a popup dialog is used to enter credentials and then the cookie access is needed by an add-in iframe to complete the authentication flow.</span></span> <span data-ttu-id="fbe70-116">ITP は、以前にポップアップ ダイアログを使用して認証を行ったサイレント認証シナリオにも影響を与える可能性がありますが、その後アドインを使用して非表示の iframe を介して認証を試みる場合があります。</span><span class="sxs-lookup"><span data-stu-id="fbe70-116">ITP could also affect silent authentication scenarios, where you have previously used a popup dialog to authenticate, but subsequent use of the add-in tries to authenticate through a hidden iframe.</span></span>

<span data-ttu-id="fbe70-117">Mac でOfficeを開発する場合、サードパーティの Cookie へのアクセスは MacOS Big Sur SDK によってブロックされます。</span><span class="sxs-lookup"><span data-stu-id="fbe70-117">When developing Office Add-ins on Mac, access to third-party cookies is blocked by the MacOS Big Sur SDK.</span></span> <span data-ttu-id="fbe70-118">これは、WKWebView ITP が Safari ブラウザーで既定で有効にされ、WKWebView によってすべてのサードパーティ Cookie がブロックされるためです。</span><span class="sxs-lookup"><span data-stu-id="fbe70-118">This is because WKWebView ITP is enabled by default on the Safari browser, and WKWebView blocks all third-party cookies.</span></span> <span data-ttu-id="fbe70-119">Officeバージョン 16.44 以降のバージョンは、MacOS Big Sur SDK と統合されています。</span><span class="sxs-lookup"><span data-stu-id="fbe70-119">Office on Mac version 16.44 or later is integrated with the MacOS Big Sur SDK.</span></span>

<span data-ttu-id="fbe70-120">Safari ブラウザーで、エンド ユーザーは、[基本設定のプライバシー] の[クロスサイト追跡を防止する] チェック ボックスをオンに切り替え  >  、ITP をオフにできます。</span><span class="sxs-lookup"><span data-stu-id="fbe70-120">In the Safari browser, end users can toggle the **Prevent cross-site tracking** checkbox under **Preference** > **Privacy** to turn off ITP.</span></span> <span data-ttu-id="fbe70-121">ただし、埋め込み WKWebView コントロールの ITP をオフにすることはできません。</span><span class="sxs-lookup"><span data-stu-id="fbe70-121">However, ITP cannot be turned off for the embedded WKWebView control.</span></span>

## <a name="see-also"></a><span data-ttu-id="fbe70-122">関連項目</span><span class="sxs-lookup"><span data-stu-id="fbe70-122">See also</span></span>

- [<span data-ttu-id="fbe70-123">サードパーティの Cookie がブロックされている Safari や他のブラウザーで ITP を処理する</span><span class="sxs-lookup"><span data-stu-id="fbe70-123">Handle ITP in Safari and other browsers where third-party cookies are blocked</span></span>](/azure/active-directory/develop/reference-third-party-cookies-spas)
- [<span data-ttu-id="fbe70-124">WebKit での追跡防止</span><span class="sxs-lookup"><span data-stu-id="fbe70-124">Tracking Prevention in WebKit</span></span>](https://webkit.org/tracking-prevention/)
- [<span data-ttu-id="fbe70-125">Chrome の "Privacy Sandbox"</span><span class="sxs-lookup"><span data-stu-id="fbe70-125">Chrome’s “Privacy Sandbox”</span></span>](https://blog.chromium.org/2020/01/building-more-private-web-path-towards.html)
- [<span data-ttu-id="fbe70-126">Access API Storage紹介</span><span class="sxs-lookup"><span data-stu-id="fbe70-126">Introducing the Storage Access API</span></span>](https://blogs.windows.com/msedgedev/2020/07/08/introducing-storage-access-api/)