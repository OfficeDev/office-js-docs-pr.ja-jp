---
title: Office アドインにおける同一生成元ポリシーの制限への対処
description: ''
ms.date: 12/04/2017
localization_priority: Priority
ms.openlocfilehash: 75bc42cd7d2a7acc8cb57ee08807a8486e21f467
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/23/2019
ms.locfileid: "29387758"
---
# <a name="addressing-same-origin-policy-limitations-in-office-add-ins"></a>Office アドインにおける同一生成元ポリシーの制限への対処


ブラウザーによって適用される同一生成元ポリシーでは、あるドメインから読み込まれたスクリプトで別のドメインの Web ページのプロパティを取得または操作できないようにしています。つまり、既定で、要求された URL のドメインは現在の Web ページのドメインと同じである必要があります。たとえば、このポリシーを適用すると、あるドメインの Web ページから、そのページがホストされているドメインとは別のドメインに対して [XmlHttpRequest](https://www.w3.org/TR/XMLHttpRequest/) Web サービスを呼び出せません。

Office アドインはブラウザー コントロールでホストされるので、それらの Web ページで実行されるスクリプトにも同一生成元ポリシーが適用されます。

アドインを開発する際に、同一生成元ポリシーの適用に対処するには、次のようにします。

- JSON/P を使用して匿名アクセスする。 
    
- トークン ベースの認証スキームを使用してサーバーサイド スクリプトを実装する。
    
- クロス オリジン リソース共有 (CORS) を使用する。
    
- IFRAME および POST MESSAGE を使用して独自のプロキシを作成する。
    

## <a name="using-jsonp-for-anonymous-access"></a>JSON/P を使用した匿名アクセス


この制限に対処する方法の 1 つは、JSON/P を使用して Web サービスのプロキシを提供することです。これを行うためには、任意のドメインでホストされているスクリプトを参照する `src` 属性を持つ `script` タグを使用します。`script` タグをプログラムで作成し、`src` 属性で参照する URL を動的に作成すると、URI クエリ パラメーターを介してパラメーターを URL に渡すことができます。Web サービス プロバイダーは、固有の URL で JavaScript コードを作成およびホストし、URI クエリ パラメーターに応じて異なるスクリプトを返します。それらのスクリプトは挿入された場所で実行され、想定どおりに動作します。

いずれの Office アドインでも機能する手法を使用する JSON/P の例を以下に示します。

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


## <a name="implementing-server-side-script-using-a-token-based-authentication-scheme"></a>トークン ベースの認証スキームを使用するサーバーサイド スクリプトの実装


同一生成元ポリシーの制限に対処する他の方法として、OAuth を使用する ASP ページ、または Cookie で資格情報をキャッシュする ASP ページとして、アドインの Web ページを実装する方法があります。

`System.Net` の `Cookie` オブジェクトを使用して、Cookie の値を取得および設定する方法を示すサーバー側のコード例については、[Value](https://docs.microsoft.com/dotnet/api/system.net.cookie.value?view=netframework-4.7.2) プロパティを参照してください。


## <a name="using-cross-origin-resource-sharing-cors"></a>クロス オリジン リソース共有 (CORS) の使用


[XmlHttpRequest2](https://dvcs.w3.org/hg/xhr/raw-file/tip/Overview.html) のクロス オリジン リソース共有機能を使用する例については、「[New Tricks in XMLHttpRequest2 に関する新しいヒント](https://www.html5rocks.com/en/tutorials/file/xhr2/)」の「Cross Origin Resource Sharing (CORS)」セクションを参照してください。


## <a name="building-your-own-proxy-using-iframe-and-post-message"></a>IFRAME および POST MESSAGE を使用する独自のプロキシの作成


IFRAME および POST MESSAGE を使用して独自のプロキシを作成する例については、「[Cross-Window Messaging](http://ejohn.org/blog/cross-window-messaging/)」を参照してください。


## <a name="see-also"></a>関連項目

- [Office アドインのプライバシーとセキュリティ](../concepts/privacy-and-security.md)
    
