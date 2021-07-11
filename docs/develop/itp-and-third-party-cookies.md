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
# <a name="develop-your-office-add-in-to-work-with-itp-when-using-third-party-cookies"></a>サードパーティ cookie をOffice ITP で動作するアドインを開発する

カスタム アドインOfficeサード パーティ Cookie が必要な場合、アドインを読み込んだブラウザー ランタイムによってインテリジェント 追跡防止 (ITP) が使用されている場合、これらの Cookie はブロックされます。 サードパーティの Cookie を使用してユーザーを認証したり、設定の保存などの他のシナリオで使用している場合があります。

アドインとOfficeがサードパーティの Cookie に依存している必要のある場合は、次の手順を使用して ITP を使用します。

1. OAuth [2.0 Authorization](https://tools.ietf.org/html/rfc6749)を設定して、認証ドメイン (Cookie を要求するサード パーティ) が承認トークンを Web サイト   に転送します。 トークンを使用して、サーバーセットの Secure Cookie と HttpOnly Cookie を使用してファースト パーティのログイン [セッションを確立します](https://developer.mozilla.org/docs/Web/HTTP/Cookies#Secure_and_HttpOnly_cookies)。
2. サードパーティが[ファーストStorage Cookie](https://webkit.org/blog/8124/introducing-storage-access-api/)へのアクセスを取得するためのアクセス許可を要求するには、Storage Access API   を使用します。 Mac 上の現在OfficeバージョンとOffice on the web API がサポートされています。
    > [!NOTE]
    > 認証以外の目的で Cookie を使用している場合は、シナリオでの使用 `localStorage` を検討してください。

次のコード サンプルは、Access API のStorage示しています。

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

## <a name="about-itp-and-third-party-cookies"></a>ITP およびサード パーティの Cookie について

サード パーティ Cookie は、ドメインがトップ レベル のフレームとは異なる iframe に読み込まれる Cookie です。 ITP は複雑な認証シナリオに影響を与える可能性があります。ポップアップ ダイアログを使用して資格情報を入力し、認証フローを完了するためにアドイン iframe によって Cookie アクセスが必要になります。 ITP は、以前にポップアップ ダイアログを使用して認証を行ったサイレント認証シナリオにも影響を与える可能性がありますが、その後アドインを使用して非表示の iframe を介して認証を試みる場合があります。

Mac でOfficeを開発する場合、サードパーティの Cookie へのアクセスは MacOS Big Sur SDK によってブロックされます。 これは、WKWebView ITP が Safari ブラウザーで既定で有効にされ、WKWebView によってすべてのサードパーティ Cookie がブロックされるためです。 Officeバージョン 16.44 以降のバージョンは、MacOS Big Sur SDK と統合されています。

Safari ブラウザーで、エンド ユーザーは、[基本設定のプライバシー] の[クロスサイト追跡を防止する] チェック ボックスをオンに切り替え  >  、ITP をオフにできます。 ただし、埋め込み WKWebView コントロールの ITP をオフにすることはできません。

## <a name="see-also"></a>関連項目

- [サードパーティの Cookie がブロックされている Safari や他のブラウザーで ITP を処理する](/azure/active-directory/develop/reference-third-party-cookies-spas)
- [WebKit での追跡防止](https://webkit.org/tracking-prevention/)
- [Chrome の "Privacy Sandbox"](https://blog.chromium.org/2020/01/building-more-private-web-path-towards.html)
- [Access API Storage紹介](https://blogs.windows.com/msedgedev/2020/07/08/introducing-storage-access-api/)