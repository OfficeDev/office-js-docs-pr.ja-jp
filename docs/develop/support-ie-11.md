---
title: Internet Explorer 11 をサポート
description: アドインで Internet Explorer 11 および ES5 Javascript をサポートする方法について説明します。
ms.date: 05/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: aff6004af4ce28aea865cb34cd34e13e23fb549f
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810275"
---
# <a name="support-internet-explorer-11"></a>Internet Explorer 11 をサポート

> [!IMPORTANT]
> **Office アドインで引き続き使用される Internet Explorer**
>
> Office 2019 の永続的なバージョンを含む、プラットフォームと Office バージョンの組み合わせによっては、Office アドイン [で使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md)で説明されているように、Internet Explorer 11 に付属する Webview コントロールを使用してアドインをホストします。Internet Explorer Webview でアドインを起動したときに、アドインのユーザーに正常なエラー メッセージを提供することで、少なくとも最小限の方法で、これらの組み合わせを引き続きサポートすることをお勧めします (ただし、必要ありません)。 次の点に注意してください。
>
> - Internet Explorer でOffice on the webが開かなくなりました。 そのため、[AppSource](/office/dev/store/submit-to-appsource-via-partner-center) は、ブラウザーとして Internet Explorer を使用してOffice on the webアドインをテストしなくなりました。
> - AppSource は引き続き Internet Explorer を使用するプラットフォームと Office *デスクトップ* バージョンの組み合わせをテストしますが、アドインが Internet Explorer をサポートしていない場合にのみ警告を発行します。アドインは AppSource によって拒否されません。
> - [Script Lab ツール](../overview/explore-with-script-lab.md)は Internet Explorer をサポートしなくなりました。

Office アドインは、Office on the webで実行するときに IFrame 内に表示される Web アプリケーションです。 Office アドインは、Windows 上の Office または Mac 上の Office で実行するときに、埋め込みブラウザー コントロールを使用して表示されます。 埋め込みブラウザー コントロールは、オペレーティング システムまたはユーザーのコンピューターにインストールされているブラウザーによって提供されます。

古いバージョンの Windows と Office をサポートする予定の場合、アドインは Internet Explorer 11 (IE11) に基づく埋め込み可能なブラウザー コントロールで動作する必要があります。 IE11 ベースのブラウザー コントロールを使用する Windows と Office の組み合わせについては、「 [Office アドインで使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md)」を参照してください。

> [!IMPORTANT]
> Internet Explorer 11 では、メディア、記録、場所など、一部の HTML5 機能はサポートされていません。 アドインで Internet Explorer 11 をサポートする必要がある場合は、これらのサポートされていない機能を回避するようにアドインを設計するか、アドインで Internet Explorer がいつ使用されているかを検出し、サポートされていない機能を使用しない代替エクスペリエンスを提供する必要があります。 詳細については、「アドイン [が Internet Explorer で実行されているかどうかを実行時に判断](#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer)する」を参照してください。

## <a name="support-for-recent-versions-of-javascript"></a>JavaScript の最新バージョンのサポート

Internet Explorer 11 では、ES5 より後の JavaScript バージョンはサポートされていません。 ECMAScript 2015 以降または TypeScript の構文と機能を使用する場合は、この記事の説明に従って 2 つのオプションがあります。 これら 2 つの手法を組み合わせることもできます。

### <a name="use-a-transpiler"></a>トランスパイラーを使用する

TypeScript またはモダン JavaScript でコードを記述し、ビルド時に ES5 JavaScript にコードをトランスパイルできます。 結果として得られる ES5 ファイルは、アドインの Web アプリケーションにアップロードするファイルです。

人気のあるトランスパイラーは 2 つあります。 どちらも、TypeScript または ES5 後の JavaScript のソース ファイルを操作できます。 また、React ファイル (.jsx および .tsx) でも動作します。

- [バベル](https://babeljs.io/)
- [Tsc](https://www.typescriptlang.org/index.html)

アドイン プロジェクトでの transpiler のインストールと構成については、いずれかのドキュメントを参照してください。 トランスパイルを自動化するには、 [Grunt](https://gruntjs.com/) や [WebPack](https://webpack.js.org/) などのタスク ランナーを使用することをお勧めします。 tsc を使用するサンプル アドインについては、「[Office アドイン Microsoft Graph React](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React)」を参照してください。 babel を使用するサンプルについては、「 [オフライン ストレージ アドイン](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/Excel.OfflineStorageAddin)」を参照してください。

> [!NOTE]
> (Visual Studio Code ではなく) Visual Studio を使用している場合は、tsc を使用するのが最も簡単です。 nuget パッケージを使用してサポートをインストールできます。 詳細については、「 [Visual Studio 2019 の JavaScript と TypeScript](/visualstudio/javascript/javascript-in-vs-2019)」を参照してください。 Visual Studio で babel を使用するには、ビルド スクリプトを作成するか、 [WebPack タスク](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.WebPackTaskRunner) ランナーや [NPM タスク](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.NPMTaskRunner) ランナーなどのツールで Visual Studio のタスク ランナー エクスプローラーを使用します。

### <a name="use-a-polyfill"></a>ポリフィルを使用する

[ポリフィル](https://en.wikipedia.org/wiki/Polyfill_(programming))は、より新しいバージョンの JavaScript の機能を複製する以前のバージョンの JavaScript です。 ポリフィルは、以降の JavaScript バージョンをサポートしていないブラウザーで機能します。 たとえば、文字列メソッド `startsWith` は ES5 バージョンの JavaScript に含まれていないため、Internet Explorer 11 では実行されません。 メソッドを定義して実装する ES5 で記述されたポリフィル ライブラリがあります `startsWith` 。 [core-js](https://github.com/zloirock/core-js) ポリフィル ライブラリをお勧めします。

ポリフィル ライブラリを使用するには、他の JavaScript ファイルやモジュールと同様に読み込みます。 たとえば、アドインのホーム ページ HTML ファイル (たとえば`<script src="/js/core-js.js"></script>`) でタグを使用`<script>`したり、`import 'core-js';`JavaScript ファイルで ステートメントを使用`import`したりできます (例: )。 JavaScript エンジンに のような `startsWith`メソッドが表示されると、最初にその名前のメソッドが言語に組み込まれているかどうかを確認します。 ある場合は、ネイティブ メソッドを呼び出します。 メソッドが組み込まれていない場合にのみ、エンジンは読み込まれたすべてのファイルを検索します。 そのため、ポリフィルされたバージョンは、ネイティブ バージョンをサポートするブラウザーでは使用されません。

core-js ライブラリ全体をインポートすると、すべての core-js 機能がインポートされます。 Office アドインで必要なポリフィルのみをインポートすることもできます。 これを行う方法については、 [CommonJS API に関するページを](https://github.com/zloirock/core-js#commonjs-api)参照してください。 core-js ライブラリには、必要なほとんどのポリフィルがあります。 core-js ドキュメントの [「不足している Polyfills](https://github.com/zloirock/core-js#missing-polyfills) 」セクションで詳しく説明されている例外がいくつかあります。 たとえば、サポートされていません `fetch`が、 [フェッチ](https://github.com/github/fetch) ポリフィルを使用できます。

core.jsを使用するサンプル アドインについては、「 [Word アドイン Angular2 StyleChecker](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)」を参照してください。

## <a name="determine-at-runtime-if-the-add-in-is-running-in-internet-explorer"></a>アドインが Internet Explorer で実行されているかどうかを実行時に確認する

アドインが Internet Explorer で実行されているかどうかを検出するには、 [window.navigator.userAgent](https://developer.mozilla.org/docs/Web/API/Navigator/userAgent) プロパティを読み取ります。 これにより、アドインは代替エクスペリエンスを提供するか、正常に失敗します。 次に例を示します。 Internet Explorer では、userAgent の値として "Trident" で始まる文字列が送信されることに注意してください。

```javascript
if (navigator.userAgent.indexOf("Trident") === -1) {

    // IE is not the browser. Provide a full-featured version of the add-in here.

} else {

    // IE is the browser. So here, do one of the following: 
    //  1. Provide an alternate experience that does not use any of the HTML5
    //     features that are not supported in IE.
    //  2. Enable the add-in to gracefully fail by putting a message in the UI that
    //     says something similar to: 
    //      "This add-in won't run in your version of Office. Please upgrade 
    //      either to perpetual Office 2021 or to a Microsoft 365 account."          

}
```

> [!IMPORTANT]
> 通常、プロパティを読み取 `userAgent` る方法は適していません。 [「ユーザー エージェントを使用したブラウザーの検出](https://developer.mozilla.org/docs/Web/HTTP/Browser_detection_using_the_user_agent)」の記事に精通していることを確認してください。これには、推奨事項や読み取`userAgent`りの代替手段が含まれます。 特に、上記の句でオプション 1 を `else` 使用する場合は、ユーザー エージェントをテストする代わりに機能検出を使用することを検討してください。
>
> 2021 年 9 月 30 日の時点で、「 [探している情報がユーザー エージェントのどの部分に含まれているか](https://developer.mozilla.org/docs/Web/HTTP/Browser_detection_using_the_user_agent#which_part_of_the_user_agent_contains_the_information_you_are_looking_for) 」セクションのテキストは、Internet Explorer 11 がリリースされる前の日付です。 これは一般的に正確であり、英語版の記事のセクションの *テーブル* は最新です。 同様に、英語以外のバージョンの記事のテキストとほとんどの場合、テーブルは古いものです。

## <a name="test-an-add-in-on-internet-explorer"></a>Internet Explorer でアドインをテストする

[Internet Explorer 11 のテストに](../testing/ie-11-testing.md)関するページを参照してください。

## <a name="additional-resources"></a>その他のリソース

- [ECMAScript 6 互換性テーブル](https://kangax.github.io/compat-table/es6/)
- [使用できますか...HTML5、CSS3 などのサポート テーブル](https://caniuse.com/)
