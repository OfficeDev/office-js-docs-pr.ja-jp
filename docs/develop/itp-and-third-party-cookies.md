---
title: サード パーティの Cookie を使用する場合に ITP を使用するように Office アドインを開発する
description: サード パーティの Cookie を使用するときに ITP アドインと Office アドインを操作する方法
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 9e2a949045fdb5bff87480d1077e692f5e8b9af6
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889270"
---
# <a name="develop-your-office-add-in-to-work-with-itp-when-using-third-party-cookies"></a>サード パーティの Cookie を使用する場合に ITP を使用するように Office アドインを開発する

Office アドインにサード パーティの Cookie が必要な場合、アドインを読み込んだブラウザー ランタイムによってインテリジェント追跡防止 (ITP) が使用されている場合、これらの Cookie はブロックされます。 サード パーティの Cookie を使用してユーザーを認証したり、設定の保存などのその他のシナリオに使用したりすることがあります。

Office アドインと Web サイトがサード パーティの Cookie に依存する必要がある場合は、次の手順に従って ITP を操作します。

1. [OAuth 2.0 Authorization](https://tools.ietf.org/html/rfc6749) を設定して、認証ドメイン (お客様の場合は Cookie を必要とするサード パーティ) が認証トークンを Web サイトに転送するようにします。 トークンを使用して、サーバー セットの Secure Cookie と [HttpOnly Cookie](https://developer.mozilla.org/docs/Web/HTTP/Cookies#Secure_and_HttpOnly_cookies) を使用してファースト パーティのログイン セッションを確立します。
1. サードパーティがファースト パーティの Cookie へのアクセスを取得するためのアクセス許可を要求できるように、 [Storage Access API](https://webkit.org/blog/8124/introducing-storage-access-api/) を使用します。 現在のバージョンの Office on Mac とOffice on the webの両方がこの API をサポートしています。
    > [!NOTE]
    > 認証以外の目的で Cookie を使用している場合は、シナリオでの使用 `localStorage` を検討してください。

次のコード サンプルは、Storage Access API を使用する方法を示しています。

```javascript
function displayLoginButton() {
  const button = createLoginButton();
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

## <a name="about-itp-and-third-party-cookies"></a>ITP とサード パーティの Cookie について

サード パーティの Cookie は、ドメインが最上位フレームとは異なる iframe に読み込まれる Cookie です。 ITP は、資格情報の入力にポップアップ ダイアログが使用され、認証フローを完了するためにアドイン iframe によって Cookie アクセスが必要になる複雑な認証シナリオに影響を与える可能性があります。 ITP は、以前にポップアップ ダイアログを使用して認証を行ったサイレント認証シナリオにも影響を及ぼす可能性がありますが、それ以降アドインを使用すると、非表示の iframe を使用して認証が試行されます。

Mac で Office アドインを開発する場合、サードパーティの Cookie へのアクセスは MacOS Big Sur SDK によってブロックされます。 これは、Safari ブラウザーで WKWebView ITP が既定で有効になっており、WKWebView によってすべてのサード パーティ Cookie がブロックされるためです。 Office on Mac バージョン 16.44 以降は、MacOS Big Sur SDK と統合されています。

Safari ブラウザーのエンド ユーザーは、[**基本設定** > **のプライバシー**] で [**クロスサイト追跡を禁止** する] チェック ボックスを切り替えて ITP をオフにすることができます。 ただし、埋め込み WKWebView コントロールの ITP をオフにすることはできません。

## <a name="see-also"></a>関連項目

- [サードパーティの Cookie がブロックされている Safari やその他のブラウザーで ITP を処理する](/azure/active-directory/develop/reference-third-party-cookies-spas)
- [WebKit での追跡防止](https://webkit.org/tracking-prevention/)
- [Chrome の "プライバシー サンドボックス"](https://blog.chromium.org/2020/01/building-more-private-web-path-towards.html)
- [Storage Access API の概要](https://blogs.windows.com/msedgedev/2020/07/08/introducing-storage-access-api/)
