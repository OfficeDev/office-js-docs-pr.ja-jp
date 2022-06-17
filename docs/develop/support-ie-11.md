---
title: Internet Explorer 11 をサポート
description: アドインで Internet Explorer 11 および ES5 Javascript をサポートする方法について説明します。
ms.date: 05/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1cb641f1ed1a75fcff23291d1fa566bbf6dc008b
ms.sourcegitcommit: fb3b1c6055e664d015703623661d624251ceb6b7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/17/2022
ms.locfileid: "66136426"
---
# <a name="support-internet-explorer-11"></a>Internet Explorer 11 をサポート

> [!IMPORTANT]
> **Office アドインで引き続き使用される Internet Explorer**
>
> Office 2019 までの 1 回限りの購入バージョンなど、一部のプラットフォームとOffice バージョンの組み合わせでは、Office アドイン[で使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md)で説明されているように、Internet Explorer 11 に付属する Web ビュー コントロールを引き続き使用してアドインをホストします。Internet Explorer Webview でアドインを起動したときにアドインのユーザーに正常なエラー メッセージを提供することで、少なくとも最小限の方法でこれらの組み合わせを引き続きサポートすることをお勧めします (ただし、必要ありません)。 次の点に注意してください。
>
> - Internet Explorer でOffice on the webが開かなくなりました。 その結果、[AppSource は](/office/dev/store/submit-to-appsource-via-partner-center)、ブラウザーとして Internet Explorer を使用してOffice on the webでアドインをテストしなくなりました。
> - AppSource は引き続き Internet Explorer を使用するプラットフォームとOffice *デスクトップ* バージョンの組み合わせをテストしますが、アドインが Internet Explorer をサポートしていない場合にのみ警告が発行されます。アドインは AppSource によって拒否されません。
> - [Script Lab ツール](../overview/explore-with-script-lab.md)は Internet Explorer をサポートしなくなりました。

Office アドインは、Office on the webで実行するときに IFrame 内に表示される Web アプリケーションです。 Office アドインは、Mac 上のWindowsまたはOfficeでOfficeで実行するときに、埋め込みブラウザー コントロールを使用して表示されます。 埋め込みブラウザー コントロールは、オペレーティング システムまたはユーザーのコンピューターにインストールされているブラウザーによって提供されます。

以前のバージョンのWindowsとOfficeをサポートする予定の場合、アドインは Internet Explorer 11 (IE11) に基づく埋め込み可能なブラウザー コントロールで動作する必要があります。 IE11 ベースのブラウザー コントロールを使用するWindowsとOfficeの組み合わせについては、「[Office アドインで使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md)」を参照してください。

> [!IMPORTANT]
> Internet Explorer 11 では、メディア、レコーディング、場所などの一部の HTML5 機能はサポートされていません。 アドインが Internet Explorer 11 をサポートする必要がある場合は、これらのサポートされていない機能を回避するためにアドインを設計するか、アドインで Internet Explorer がいつ使用されているかを検出し、サポートされていない機能を使用しない代替エクスペリエンスを提供する必要があります。 詳細については、「 [Internet Explorer でアドインが実行されているかどうかを実行時に判断](#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer)する」を参照してください。

## <a name="support-for-recent-versions-of-javascript"></a>JavaScript の最新バージョンのサポート

Internet Explorer 11 では、ES5 より後の JavaScript バージョンはサポートされていません。 ECMAScript 2015 以降 (TypeScript) の構文と機能を使用する場合は、この記事で説明する 2 つのオプションがあります。 これら 2 つの手法を組み合わせることもできます。

### <a name="use-a-transpiler"></a>transpiler を使用する

TypeScript または最新の JavaScript でコードを記述し、ビルド時に ES5 JavaScript にトランスパイルできます。 結果として得られる ES5 ファイルは、アドインの Web アプリケーションにアップロードするファイルです。

一般的なトランスパイラーは 2 つあります。 どちらも、TypeScript または Post-ES5 JavaScript であるソース ファイルを操作できます。 また、React ファイル (.jsx と .tsx) でも動作します。

- [バベル](https://babeljs.io/)
- [Tsc](https://www.typescriptlang.org/index.html)

アドイン プロジェクトでの transpiler のインストールと構成については、どちらかのドキュメントを参照してください。 [Grunt](https://gruntjs.com/) や [WebPack](https://webpack.js.org/) などのタスク ランナーを使用して、トランスパイルを自動化することをお勧めします。 tsc を使用するサンプル アドインについては、「[Microsoft Graph ReactアドインOffice](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React)参照してください。 バベルを使用するサンプルについては、「[オフライン Storage アドイン](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/Excel.OfflineStorageAddin)」を参照してください。

> [!NOTE]
> Visual Studio (Visual Studio コードではない) を使用している場合は、おそらく tsc を使用するのが最も簡単です。 nuget パッケージを使用して、サポートをインストールできます。 詳細については、[Visual Studio 2019 の JavaScript と TypeScript を参照してください](/visualstudio/javascript/javascript-in-vs-2019)。 Visual Studioでバベルを使用するには、ビルド スクリプトを作成するか、[WebPack タスク](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.WebPackTaskRunner) ランナーや [NPM](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.NPMTaskRunner) タスク ランナーなどのツールを使用してVisual Studioでタスク ランナー エクスプローラーを使用します。

### <a name="use-a-polyfill"></a>ポリフィルを使用する

[ポリフィル](https://en.wikipedia.org/wiki/Polyfill_(programming))は、より新しいバージョンの JavaScript の機能を複製する、以前のバージョンの JavaScript です。 ポリフィルは、以降の JavaScript バージョンをサポートしていないブラウザーで機能します。 たとえば、文字列メソッド `startsWith` は ES5 バージョンの JavaScript の一部ではないため、Internet Explorer 11 では実行されません。 メソッドを定義して実装 `startsWith` する ES5 で記述されたポリフィル ライブラリがあります。 [core-js](https://github.com/zloirock/core-js) のポリフィル ライブラリをお勧めします。

ポリフィル ライブラリを使用するには、他の JavaScript ファイルまたはモジュールと同様に読み込みます。 たとえば、アドインのホーム ページ HTML ファイル (たとえば`<script src="/js/core-js.js"></script>`) でタグを使用`<script>``import`したり、JavaScript ファイル内のステートメント (たとえば、 `import 'core-js';` JavaScript エンジンにこのような `startsWith`メソッドが表示されると、最初にその名前のメソッドが言語に組み込まれているかどうかを確認します。 存在する場合は、ネイティブ メソッドを呼び出します。 メソッドが組み込まれていない場合にのみ、エンジンは読み込まれたすべてのファイルを検索します。 そのため、ネイティブ バージョンをサポートするブラウザーでは、ポリフィルされたバージョンは使用されません。

core-js ライブラリ全体をインポートすると、すべての core-js 機能がインポートされます。 また、Office アドインで必要なポリフィルのみをインポートすることもできます。 これを行う方法については、 [CommonJS API を](https://github.com/zloirock/core-js#commonjs-api)参照してください。 core-js ライブラリには、必要なほとんどのポリフィルがあります。 core-js ドキュメントの [「Missing Polyfills](https://github.com/zloirock/core-js#missing-polyfills) 」セクションで詳しく説明されている例外がいくつかあります。 たとえば、サポートされていません `fetch`が、 [フェッチ](https://github.com/github/fetch) ポリフィルを使用できます。

core.jsを使用するサンプル アドインについては、「 [Word アドイン Angular2 StyleChecker」を](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)参照してください。

## <a name="determine-at-runtime-if-the-add-in-is-running-in-internet-explorer"></a>アドインが Internet Explorer で実行されているかどうかを実行時に確認する

アドインは、 [window.navigator.userAgent](https://developer.mozilla.org/docs/Web/API/Navigator/userAgent) プロパティを読み取ることで、Internet Explorer で実行されているかどうかを検出できます。 これにより、アドインは代替エクスペリエンスを提供するか、正常に失敗します。 次に例を示します。 Internet Explorer では、userAgent の値として "Trident" で始まる文字列が送信されることに注意してください。

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
> 通常、プロパティを読み取 `userAgent` るのをお勧めしません。 ユーザー [エージェントを使用したブラウザーの検出](https://developer.mozilla.org/docs/Web/HTTP/Browser_detection_using_the_user_agent)に関する記事(推奨事項や読み取りの `userAgent`代替手段など)について理解していることを確認してください。 特に、上記の句でオプション 1 を `else` 使用している場合は、ユーザー エージェントのテストではなく機能検出を使用することを検討してください。
>
> 2021 年 9 月 30 日の時点で、「 [ユーザー エージェントのどの部分に探している情報が含まれているか」セクションのテキストは](https://developer.mozilla.org/docs/Web/HTTP/Browser_detection_using_the_user_agent#which_part_of_the_user_agent_contains_the_information_you_are_looking_for) 、Internet Explorer 11 がリリースされる前の日付です。 一般的にはまだ正確であり、英語版の記事のセクションの *テーブル* は最新です。 同様に、記事の英語以外のバージョンのテキストとほとんどの場合、テーブルは最新ではありません。

## <a name="test-an-add-in-on-internet-explorer"></a>Internet Explorer でアドインをテストする

[Internet Explorer 11 のテストを](../testing/ie-11-testing.md)参照してください。

## <a name="additional-resources"></a>その他のリソース

- [ECMAScript 6 互換性テーブル](https://kangax.github.io/compat-table/es6/)
- [使用できますか...HTML5、CSS3 などのサポート テーブル](https://caniuse.com/)
