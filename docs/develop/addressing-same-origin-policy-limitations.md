---
title: Office アドインにおける同一生成元ポリシーの制限への対処
description: JSONP、CORS、IFRAMEs、その他の手法で同じオリジン ポリシーの制限に対応する方法について説明します。
ms.date: 10/17/2019
ms.localizationpriority: medium
ms.openlocfilehash: fa152bf42b1d0f7ad16172324c7a9e75314e4f34
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743657"
---
# <a name="addressing-same-origin-policy-limitations-in-office-add-ins"></a>Office アドインにおける同一生成元ポリシーの制限への対処

ブラウザーによって適用される同一生成元ポリシーでは、あるドメインから読み込まれたスクリプトで別のドメインの Web ページのプロパティを取得または操作できないようにしています。つまり、既定で、要求された URL のドメインは現在の Web ページのドメインと同じである必要があります。たとえば、このポリシーを適用すると、あるドメインの Web ページから、そのページがホストされているドメインとは別のドメインに対して [XmlHttpRequest](https://www.w3.org/TR/XMLHttpRequest/) Web サービスを呼び出せません。

Office アドインはブラウザー コントロールでホストされるので、それらの Web ページで実行されるスクリプトにも同一生成元ポリシーが適用されます。

同一生成元ポリシーは、Web アプリケーションが複数のサブドメインに渡るコンテンツと API をホストしているときなど、多くの場合に不要な制約になることがあります。 同一生成元ポリシーの適用に関する制約を安全に解消するための一般的な手法がいくつかあります。 この記事では、その一部について簡単な紹介のみを示します。 ここに示したリンクを使用して、こうした手法の調査を開始してください。

## <a name="use-jsonp-for-anonymous-access"></a>匿名アクセスに JSONP を使用する

同一生成元ポリシーの制限を解消する 1 つの方法として、[JSONP](https://www.w3schools.com/js/js_json_jsonp.asp) を使用して Web サービスのプロキシを提供します。 これを行うためには、任意のドメインでホストされているスクリプトを参照する `src` 属性を持つ `script` タグを使用します。 `script` タグをプログラムで作成し、`src` 属性で参照する URL を動的に作成すると、URI クエリ パラメーターを介してパラメーターを URL に渡すことができます。 Web サービス プロバイダーは、固有の URL で JavaScript コードを作成およびホストし、URI クエリ パラメーターに応じて異なるスクリプトを返します。 それらのスクリプトは挿入された場所で実行され、想定どおりに動作します。

次に、あらゆる Office アドインで機能する手法を使用する JSONP の例を示します。

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


## <a name="implement-server-side-code-using-a-token-based-authorization-scheme"></a>トークン ベースの認証スキームを使用してサーバー側のコードを実装する

同一生成元ポリシーの制限に対処するもう 1 つの方法として、[OAuth 2.0](https://oauth.net/2/) フローを使用するサーバー側のコードを用意します。このコードによって、別のドメインでホストされているリソースへの許可されたアクセスを可能にします。 


## <a name="use-cross-origin-resource-sharing-cors"></a>クロス オリジン リソース共有 (CORS) を使用する


[XmlHttpRequest2](https://dvcs.w3.org/hg/xhr/raw-file/tip/Overview.html) のクロス オリジン リソース共有機能を使用する例については、「[New Tricks in XMLHttpRequest2 に関する新しいヒント](https://www.html5rocks.com/en/tutorials/file/xhr2/)」の「Cross Origin Resource Sharing (CORS)」セクションを参照してください。


## <a name="build-your-own-proxy-using-iframe-and-post-message-cross-window-messaging"></a>IFRAME と POST MESSAGE を使用して独自のプロキシを作成する (クロス ウィンドウ メッセージング)


IFRAME および POST MESSAGE を使用して独自のプロキシを作成する例については、「[Cross-Window Messaging](http://ejohn.org/blog/cross-window-messaging/)」を参照してください。


## <a name="see-also"></a>関連項目

- [Office アドインのプライバシーとセキュリティ](../concepts/privacy-and-security.md)
    
