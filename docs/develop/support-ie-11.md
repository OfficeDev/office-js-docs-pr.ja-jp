---
title: Internet Explorer 11 をサポート
description: アドインで 11 Internet Explorer ES5 Javascript をサポートする方法について説明します。
ms.date: 10/22/2021
ms.localizationpriority: medium
ms.openlocfilehash: d2a504a6e030e6cf8d06c766cb500d6c11710ea9
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744231"
---
# <a name="support-internet-explorer-11"></a>Internet Explorer 11 をサポート

> [!IMPORTANT]
> **Internet Explorerアドインで引き続きOffice使用される場合**
>
> Microsoft は、アドインのサポートInternet Explorer終了していますが、これはアドインのOffice大きな影響を及ぼします。Office アドインで使用されるブラウザーで説明したように、プラットフォームと Office バージョンの一部の組み合わせ (Office 2019 までの 1 回限り購入バージョンを含む) は、Internet Explorer 11 に付属する webview コントロールを引き続き使用してアドインをホスト[](../concepts/browsers-used-by-office-web-add-ins.md)します。さらに、これらの組み合わせのサポートは、AppSource にInternet Explorerアドインに対して引き続き[必要です](/office/dev/store/submit-to-appsource-via-partner-center)。 次の *2 つの点が変化* しています。
>
> - Office on the webで開かなくなったInternet Explorer。 そのため、AppSource はブラウザーとしてアプリケーションを使用してOffice on the webアドインInternet Explorerテストしなくなりました。 ただし、AppSource は引き続き、プラットフォームとデスクトップ バージョンの組み合わせOffice *を* テストInternet Explorer。
> - この[Script Labツールは](../overview/explore-with-script-lab.md)、この機能をサポートInternet Explorer。

Officeアドインは、IFrame 内でアプリケーションを実行するときに表示される web アプリケーションOffice on the web。 Officeアドインは、Mac 上または Mac 上の Office または Office Windowsで実行するときに、埋め込みブラウザー コントロールを使用して表示されます。 埋め込みブラウザー コントロールは、オペレーティング システムまたはユーザーのコンピューターにインストールされているブラウザーによって提供されます。

AppSource を使用してアドインを販売する予定がある場合、または以前のバージョンの Windows および Office をサポートする予定の場合、アドインは Internet Explorer 11 (IE11) に基づく埋め込み可能なブラウザー コントロールで動作する必要があります。 IE11 ベースのブラウザー Windowsと Officeを使用するブラウザーの組み合わせについては、「Office アドインで使用されるブラウザー」を[参照](../concepts/browsers-used-by-office-web-add-ins.md)してください。

> [!IMPORTANT]
> Internet Explorer 11 では、メディア、レコーディング、場所などの一部の HTML5 機能はサポートされていません。 アドインで Internet Explorer 11 をサポートする必要がある場合は、これらのサポートされていない機能を回避するためにアドインを設計するか、アドインが Internet Explorer が使用されているときに検出し、サポートされていない機能を使用しない代替エクスペリエンスを提供する必要があります。 詳細については、「 [アドインが実行中](#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer)かどうかを実行時に判断する」を参照Internet Explorer。

## <a name="support-for-recent-versions-of-javascript"></a>JavaScript の最新バージョンのサポート

Internet Explorer 11 では、ES5 以降の JavaScript バージョンはサポートされていません。 ECMAScript 2015 以降または TypeScript の構文と機能を使用する場合は、この記事で説明する 2 つのオプションがあります。 これら 2 つの手法を組み合わせて使用できます。

### <a name="use-a-transpiler"></a>トランスピラーを使用する

コードは TypeScript またはモダン JavaScript で記述し、ビルド時に ES5 JavaScript にトランスピレルできます。 結果の ES5 ファイルは、アドインの Web アプリケーションにアップロードするファイルです。

2 つの一般的なトランスピラーがあります。 どちらも、TypeScript または ES5 後の JavaScript のソース ファイルを使用できます。 これらのファイルは、Reactファイル (.jsx と .tsx) でも動作します。

- [babel](https://babeljs.io/)
- [tsc](https://www.typescriptlang.org/index.html)

アドイン プロジェクトでのトランスピラーのインストールと構成の詳細については、どちらかのドキュメントを参照してください。 Grunt や [WebPack](https://webpack.js.org/) などのタスク ランナーを使用して[](https://gruntjs.com/)、トランスピレーションを自動化することをお勧めします。 tsc を使用するサンプル アドインについては、「microsoft Office アドイン」を[参照Graph React](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React)。 babel を使用するサンプルについては、「[Offline Storage アドイン」を参照してください](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/Excel.OfflineStorageAddin)。

> [!NOTE]
> (Visual Studio) Visual Studio Code tsc を使用するのが最も簡単です。 nuget パッケージを使用してサポートをインストールできます。 詳細については、「[JavaScript と TypeScript in Visual Studio 2019」を参照](/visualstudio/javascript/javascript-in-vs-2019)してください。 Visual Studio でバベルを使用するには、ビルド スクリプトを作成するか、Visual Studio で [WebPack](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.WebPackTaskRunner) タスク ランナーや NPM タスク ランナーのようなツールを使用してタスク ランナー エクスプローラーを[使用します](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.NPMTaskRunner)。

### <a name="use-a-polyfill"></a>ポリフィルを使用する

ポリ [フィルは](https://en.wikipedia.org/wiki/Polyfill_(programming)) 、より新しいバージョンの JavaScript の機能を複製する以前のバージョンの JavaScript です。 polyfill は、以降の JavaScript バージョンをサポートしないブラウザーで動作します。 たとえば、文字列 `startsWith` メソッドは ES5 バージョンの JavaScript の一部ではないので、11 では実行Internet Explorerされません。 メソッドを定義して実装する、ES5 で記述されたポリフィル ライブラリ `startsWith` があります。 [core-js polyfill](https://github.com/zloirock/core-js) ライブラリをお勧めします。

ポリフィル ライブラリを使用するには、他の JavaScript ファイルまたはモジュールと同様に読み込む必要があります。 たとえば、アドイン`<script>`のホーム ページ HTML ファイル (たとえば) でタグを使用したり、JavaScript ファイル (`<script src="/js/core-js.js"></script>``import`たとえば) でステートメントを使用できます`import 'core-js';`。 JavaScript エンジンに次 `startsWith`のようなメソッドが表示される場合は、まずその名前のメソッドが言語に組み込み込みされているのか確認します。 ある場合は、ネイティブ メソッドを呼び出します。 メソッドが組み込みではない場合にのみ、エンジンは読み込まれたすべてのファイルを検索します。 したがって、ポリフィルされたバージョンは、ネイティブ バージョンをサポートするブラウザーでは使用されません。

core-js ライブラリ全体をインポートすると、すべての core-js 機能がインポートされます。 また、アドインで必要なポリフィルOfficeインポートできます。 これを行う方法については、「 [CommonJS API」を参照してください](https://github.com/zloirock/core-js#commonjs-api)。 core-js ライブラリには、必要なポリフィルのほとんどがあります。 core-js のドキュメントの「不足している [Polyfills](https://github.com/zloirock/core-js#missing-polyfills) 」セクションで詳しく説明されているいくつかの例外があります。 たとえば、サポートされていません `fetch`が、フェッチ ポリフィル [を](https://github.com/github/fetch) 使用できます。

アプリケーションを使用するサンプル アドインcore.js [Word アドイン Angular2 StyleChecker を参照してください](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)。

## <a name="determine-at-runtime-if-the-add-in-is-running-in-internet-explorer"></a>アドインが実行中かどうかを実行時に判断Internet Explorer

アドインは、 [window.navigator.userAgent](https://developer.mozilla.org/docs/Web/API/Navigator/userAgent) プロパティを読み取Internet Explorerで実行中のアドインを検出できます。 これにより、アドインは別のエクスペリエンスを提供するか、正常に失敗します。 次に例を示します。 ユーザーはInternet Explorer値として "Trident" で始まる文字列を送信します。

```javascript
if (navigator.userAgent.indexOf("Trident") === -1) {

    // IE is not the browser. Provide a full-featured version of the add-in here.

} else {

    // IE is the browser. So here, do one of the following: 
    //  1. Provide an alternate experience that does not use any of the HTML5
    //     features that are not supported in IE.
    //  2. Enable the add-in to gracefully fail by putting a message in the UI that
    //     says something similar to: 
    //      "This add-in won't run in your version of Office. Please upgrade to 
    //      either one-time purchase Office 2021 or to a Microsoft 365 account."          

}
```

> [!IMPORTANT]
> 通常、プロパティを読み取るのは良い方法 `userAgent` ではありません。 記事、ユーザー エージェントを使用した [ブラウザー](https://developer.mozilla.org/en-US/docs/Web/HTTP/Browser_detection_using_the_user_agent)の検出に関する推奨事項や読み取り方法を理解している必要があります `userAgent`。 特に、上記の句でオプション 1 `else` を使用する場合は、ユーザー エージェントのテストではなく機能検出の使用を検討してください。
>
> 2021 年 9 月 30 日の現在、セクションのテキスト ユーザー エージェントのどの部分に探している情報が含まれていますか [?](https://developer.mozilla.org/en-US/docs/Web/HTTP/Browser_detection_using_the_user_agent#which_part_of_the_user_agent_contains_the_information_you_are_looking_for) Internet Explorer 11 がリリースされる前の日付。 一般に正確であり、英語版の記事のセクションの表は最新の情報です。 同様に、記事の英語以外のバージョンのテキストとほとんどの場合、テーブルは古いものです。

## <a name="test-an-add-in-on-internet-explorer"></a>アプリでアドインをテストInternet Explorer

「 [11 Internet Explorerテスト」を参照してください](../testing/ie-11-testing.md)。

## <a name="additional-resources"></a>その他のリソース

- [ECMAScript 6 互換テーブル](https://kangax.github.io/compat-table/es6/)
- [使用できます。..HTML5、CSS3 などのサポート テーブル](https://caniuse.com/)
