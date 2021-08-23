---
title: Internet Explorer 11 をサポート
description: アドインで 11 Internet Explorer ES5 Javascript をサポートする方法について説明します。
ms.date: 08/13/2021
localization_priority: Normal
ms.openlocfilehash: dea458cbabb71e23432db8cb6eb3dfcddc6e1bac
ms.sourcegitcommit: bc6203dd8f21d1c375039c5ee8f1388ede9be93b
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/18/2021
ms.locfileid: "58382943"
---
# <a name="support-internet-explorer-11"></a>Internet Explorer 11 をサポート

> [!IMPORTANT]
> **Internet ExplorerアドインOffice引き続き使用する**
>
> Microsoft は、アドインのサポートInternet Explorer終了していますが、これはアドインのOffice大きな影響を及ぼします。Office アドインで使用されるブラウザーで説明したように、プラットフォームと Office バージョンの一部の組み合わせ (Office 2019 までのすべての一時購入バージョンを含む) は、Internet Explorer 11 に付属する webview[](../concepts/browsers-used-by-office-web-add-ins.md)コントロールを引き続き使用してアドインをホストします。さらに、これらの組み合わせのサポートは、AppSource にInternet Explorerアドインに対して引き続き[必要です](/office/dev/store/submit-to-appsource-via-partner-center)。 次の *2 つの点が変化* しています。
>
> - AppSource は、ブラウザーとしてアプリケーションを使用してOffice on the webアドインInternet Explorerテストしなくなりました。 ただし、AppSource は引き続き、プラットフォームとデスクトップ バージョンの組み合Office *使用* するデスクトップ バージョンの組み合わせをテストInternet Explorer。
> - この[Script Labは](../overview/explore-with-script-lab.md)サポートされなくなりましたInternet Explorer。

Officeアドインは、Web アプリケーションで実行するときに IFrames 内に表示Office on the web。 Officeアドインは、Mac 上または Mac 上の Office または Windowsで実行Officeブラウザー コントロールを使用して表示されます。 埋め込みブラウザー コントロールは、オペレーティング システムまたはユーザーのコンピューターにインストールされているブラウザーによって提供されます。

AppSource を使用してアドインを販売する予定がある場合、または以前のバージョンの Windows および Office をサポートする予定の場合、アドインは Internet Explorer 11 (IE11) に基づく埋め込み可能なブラウザー コントロールで動作する必要があります。 IE11 ベースのブラウザー コントロールをWindowsとOfficeするブラウザーの組み合わせについては、「Office アドインで使用されるブラウザー」を[参照してください](../concepts/browsers-used-by-office-web-add-ins.md)。

> [!IMPORTANT]
> Internet Explorer 11 では、メディア、レコーディング、場所などの一部の HTML5 機能はサポートされていません。 アドインが 11 をサポートInternet Explorer場合は、これらの機能を使用することはできません。

Internet Explorer 11 では、ES5 以降の JavaScript バージョンはサポートされていません。 ECMAScript 2015 以降または TypeScript の構文と機能を使用する場合は、この記事で説明する 2 つのオプションがあります。 これら 2 つの手法を組み合わせて使用できます。

## <a name="use-a-transpiler"></a>トランスピラーを使用する

コードは TypeScript またはモダン JavaScript で記述し、ビルド時に ES5 JavaScript にトランスピレルできます。 結果の ES5 ファイルは、アドインの Web アプリケーションにアップロードするファイルです。

2 つの一般的なトランスピラーがあります。 どちらも、TypeScript または ES5 後の JavaScript のソース ファイルを使用できます。 これらのファイルは、Reactファイル (.jsx と .tsx) でも動作します。

- [babel](https://babeljs.io/)
- [tsc](https://www.typescriptlang.org/index.html)

アドイン プロジェクトでのトランスピラーのインストールと構成の詳細については、どちらかのドキュメントを参照してください。 Grunt や[WebPack](https://webpack.js.org/)などのタスク[](https://gruntjs.com/)ランナーを使用して、トランスピレーションを自動化することをお勧めします。 tsc を使用するサンプル アドインについては、「microsoft Office アドイン」を[参照Graph React。](https://github.com/OfficeDev/PnP-OfficeAddins/tree/3ce0e1b74152dbbe8306a091696bc4455c04c0a1/Samples/auth/Office-Add-in-Microsoft-Graph-React) babel を使用するサンプルについては[、「Offline Storage アドイン」を参照してください](https://github.com/OfficeDev/PnP-OfficeAddins/tree/3ce0e1b74152dbbe8306a091696bc4455c04c0a1/Samples/Excel.OfflineStorageAddin)。

> [!NOTE]
> ユーザーが (Visual Studio Visual Studio Code) tsc を使用する場合は、おそらく最も簡単です。 nuget パッケージを使用してサポートをインストールできます。 詳細については[、「JavaScript と TypeScript in Visual Studio 2019」を参照してください](/visualstudio/javascript/javascript-in-vs-2019)。 Visual Studio で babel を使用するには、ビルド スクリプトを作成するか、Visual Studio でタスク ランナー エクスプローラーを使用して[、WebPack](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.WebPackTaskRunner)タスク ランナーや NPM タスク ランナーのようなツールを[使用します](https://marketplace.visualstudio.com/items?itemName=MadsKristensen.NPMTaskRunner)。

## <a name="use-a-polyfill"></a>ポリフィルを使用する

ポリ [フィルは](https://en.wikipedia.org/wiki/Polyfill_(programming)) 、より新しいバージョンの JavaScript の機能を複製する以前のバージョンの JavaScript です。 polyfill は、以降の JavaScript バージョンをサポートしないブラウザーで動作します。 たとえば、文字列メソッドは ES5 バージョンの JavaScript の一部ではないので、このメソッドは `startsWith` 11 のInternet Explorerされません。 メソッドを定義して実装する、ES5 で記述されたポリフィル ライブラリ `startsWith` があります。 [core-js polyfill](https://github.com/zloirock/core-js)ライブラリをお勧めします。

ポリフィル ライブラリを使用するには、他の JavaScript ファイルまたはモジュールと同様に読み込む必要があります。 たとえば、アドインのホーム ページ HTML ファイル (たとえば) でタグを使用したり、JavaScript ファイル (たとえば) でステートメント `<script>` `<script src="/js/core-js.js"></script>` `import` を使用できます `import 'core-js';` 。 JavaScript エンジンに次のようなメソッドが表示される場合は、まずその名前のメソッドが言語に組み込 `startsWith` み込みされているのか確認します。 ある場合は、ネイティブ メソッドを呼び出します。 メソッドが組み込みではない場合にのみ、エンジンは読み込まれたすべてのファイルを検索します。 したがって、ポリフィルされたバージョンは、ネイティブ バージョンをサポートするブラウザーでは使用されません。

core-js ライブラリ全体をインポートすると、すべての core-js 機能がインポートされます。 また、アドインで必要なポリフィルOfficeインポートできます。 これを行う方法については [、「CommonJS API」を参照してください](https://github.com/zloirock/core-js#commonjs-api)。 core-js ライブラリには、必要なポリフィルのほとんどがあります。 core-js のドキュメントの「不足している [Polyfills」](https://github.com/zloirock/core-js#missing-polyfills) セクションで詳しく説明されているいくつかの例外があります。 たとえば、サポートされていませんが、フェッチ `fetch` ポリフィル [を](https://github.com/github/fetch) 使用できます。

このアプリケーションを使用するサンプル アドインcore.js Word アドイン [Angular2 StyleChecker を参照してください](https://github.com/OfficeDev/Word-Add-in-Angular2-StyleChecker)。

## <a name="testing-an-add-in-on-internet-explorer"></a>アプリでアドインをテストInternet Explorer

「Internet Explorer [11 テスト」を参照してください](../testing/ie-11-testing.md)。

## <a name="additional-resources"></a>その他のリソース

- [ECMAScript 6 互換テーブル](https://kangax.github.io/compat-table/es6/)
- [使用できます。..HTML5、CSS3 などのサポート テーブル](https://caniuse.com/)
