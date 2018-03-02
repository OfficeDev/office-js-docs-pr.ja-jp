---
title: Word アドインの概要
description: ''
ms.date: 01/23/2018
---


# <a name="word-add-ins-overview"></a>Word アドインの概要

Word の機能を拡張するソリューション (たとえば、ドキュメントの自動アセンブリや、他のデータ ソースからの Word 文書のデータへのバインドやアクセスを可能にするソリューション) を作成したい場合は、Word JavaScript API と JavaScript API for Office を含む Office アドイン プラットフォームを使用して、Windows デスクトップ、Mac、またはクラウドで実行する Word クライアントを拡張できます。

Word のアドインは、[Office アドイン プラットフォーム](../overview/office-add-ins.md)にある数多くの開発オプションのひとつです。アドイン コマンドを使用して、Word の UI を拡張したり、Word 文書のコンテンツと対話する JavaScript を実行する作業ウィンドウを起動したりすることができます。ブラウザーで実行できるあらゆるコードは、Word アドインで実行できます。Word 文書のコンテンツと対話するアドインは、Word オブジェクトを操作し、オブジェクトの状態を同期する要求を作成します。 

> [!NOTE]
> アドインをビルドするとき、アドインを AppSource に[発行](../publish/publish.md)する予定であれば、[AppSource 検証ポリシー](https://docs.microsoft.com/ja-jp/office/dev/store/validation-policies)に準拠していることを確認してください。たとえば、検証に合格するには、アドインは、定義したメソッドをサポートするすべてのプラットフォーム全体で機能する必要があります (詳細については、[セクション 4.12](https://docs.microsoft.com/ja-jp/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) と「[Office アドインを使用できるホストおよびプラットフォーム](../overview/office-add-in-availability.md)」のページを参照してください)。

次の図は、作業ウィンドウで実行される Word アドインの例を示します。

*図 1. Word の作業ウィンドウで実行されているアドイン*

![Word の作業ウィンドウで実行されているアドイン](../images/word-add-in-show-host-client.png)

Word アドイン (1) は、Word 文書 (2) に要求を送信し、JavaScript を使用して段落オブジェクトにアクセスして段落を更新、削除、または移動することができます。たとえば、次のコードは、その段落に新しい文を追加する方法を示しています。

```js
Word.run(function (context) {
    var paragraphs = context.document.getSelection().paragraphs;
    paragraphs.load();
    return context.sync().then(function () {
        paragraphs.items[0].insertText(' New sentence in the paragraph.',
                                       Word.InsertLocation.end);
    }).then(context.sync);
});

```

ASP.NET、NodeJS、Python などの任意の Web サーバー テクノロジを使用して、Word アドインをホストさせることができます。お気に入りのクライアント側のフレームワーク (Ember、Backbone、Angular、React) を使用するか VanillaJS を引き続き使用してソリューションを開発し、Azure のようなサービスを使用してアプリケーションを[認証](../develop/use-the-oauth-authorization-framework-in-an-office-add-in.md)し、ホストできます。

Word JavaScript API を使用すると、アプリケーションから Word 文書内にあるオブジェクトやメタデータにアクセスできます。これらの API を使用して、以下をターゲットとするアドインを作成できます。

* Word 2013 for Windows
* Word 2016 for Windows
* Word Online
* Word 2016 for Mac
* Word for iOS

アドインを 1 回作成すれば、それをプラットフォームの異なるすべてのバージョンの Word で実行できます。詳細については、「[Office アドインを使用できるホストおよびプラットフォーム](../overview/office-add-in-availability.md)」を参照してください。

## <a name="javascript-apis-for-word"></a>Word 用 JavaScript API

2 セットの JavaScript API を使用して、Word 文書のオブジェクトおよびメタデータと対話できます。1 つ目は、Office 2013 で導入された [JavaScript API for Office](https://dev.office.com/reference/add-ins/javascript-api-for-office?product=word) です。これは共有 API です -- 2 つ以上の Office クライアントでホストされているアドインで、多くのオブジェクトを使用することができます。この API は、広範囲にわたってコールバックを使用します。 

2 つ目は、[Word JavaScript API](https://dev.office.com/reference/add-ins/word/word-add-ins-reference-overview) です。これは、Mac と Windows の Word 2016 を対象とする Word アドインを作成するために使用できる、厳密に型指定されたオブジェクト モデルです。このオブジェクト モデルは promise を使用し、[本文](https://dev.office.com/reference/add-ins/word/body)、[コンテンツ コントロール](https://dev.office.com/reference/add-ins/word/contentcontrol)、[インライン画像](https://dev.office.com/reference/add-ins/word/inlinepicture)、および[段落](https://dev.office.com/reference/add-ins/word/paragraph)などの Word 固有のオブジェクトへのアクセスを提供します。Word JavaScript API には、IDE 内のコード ヒントを取得できるように、TypeScript の定義と vsdoc ファイルが含まれています。

現在、Word のすべてのクライアントは共有の JavaScript API for Office をサポートし、ほとんどのクライアントは Word JavaScript API をサポートします。サポート対象のクライアントの詳細については、「[API リファレンスのドキュメント](https://dev.office.com/reference/add-ins/javascript-api-for-office?product=word)」を参照してください。

Word JavaScript API のオブジェクト モデルはより簡単に使用できるため、Word JavaScript APから始めることをお勧めします。次のような必要がある場合は、Word JavaScript API を使用します。

* Word 文書内のオブジェクトにアクセスする。

次のような必要がある場合は、共有の JavaScript API for Office を使用します。

* Word 2013 を対象とする。
* アプリケーションの初期アクションを実行する。
* サポートされている要件セットを確認する。
* メタデータ、設定、およびドキュメントの環境情報にアクセスする。
* ドキュメント内のセクションにバインドし、イベントをキャプチャする。
* カスタム XML パーツを使用する。
* ダイアログ ボックスを開く。

## <a name="next-steps"></a>次の手順

最初の Word アドインを作成する準備はできましたか?「[最初の Word アドインをビルドする](word-add-ins.md)」を参照してください。また、対話式の「[作業の開始エクスペリエンス](http://dev.office.com/getting-started/addins?product=Word)」も使用できます。[アドインのマニフェスト](../develop/add-in-manifests.md)を使用して、アドインがホストされている場所や表示方法の説明と、アクセス許可およびその他の情報の定義を行います。

ユーザーにとって魅力的なエクスペリエンスを提供する世界クラスの Word アドインを設計する方法の詳細については、「[設計のガイドライン](../design/add-in-design.md)」と「[ベスト プラクティス](../concepts/add-in-development-best-practices.md)」を参照してください。

アドインを作成したら、ネットワーク共有、アプリ カタログ、または AppSource に[発行](../publish/publish.md)できます。

## <a name="whats-coming-up-for-word-add-ins"></a>今後の Word アドイン

新しい Word アドイン用の API の設計と開発にあたり、[API のオープン仕様](https://dev.office.com/reference/add-ins/openspec)ページでこれらに対するフィードバックの提供が可能になります。Word JavaScript API 用のパイプラインの新機能をご確認いただき、設計の仕様に関する情報をお寄せください。

## <a name="see-also"></a>関連項目

* [Office アドイン プラットフォームの概要](../overview/office-add-ins.md)
* [Word JavaScript API リファレンス](https://dev.office.com/reference/add-ins/word/word-add-ins-reference-overview)

