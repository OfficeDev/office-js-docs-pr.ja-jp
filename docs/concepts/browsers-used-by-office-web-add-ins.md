---
title: Office アドインによって使用されるブラウザー
description: Office アドインによって使用されるブラウザーをオペレーティング システムおよび Office バージョンが決定する方法を指定します。
ms.date: 01/20/2021
localization_priority: Normal
ms.openlocfilehash: c540eece3b74bb043cc8f4921c7c774511b5a60a
ms.sourcegitcommit: 54d141cefb7bdc5f16330747d0ec8e8e2bd03e93
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/21/2021
ms.locfileid: "49916462"
---
# <a name="browsers-used-by-office-add-ins"></a>Office アドインによって使用されるブラウザー

Officeは、web 上の Office で実行し、デスクトップおよびモバイル クライアント用の Office で埋め込みブラウザー コントロールを使用するときに iFrame を使用して表示される Web アプリケーションです。 アドインには JavaScript を実行するための JavaScript エンジンも必要です。 埋め込みブラウザーとエンジンの両方が、ユーザーのコンピューターにインストールされているブラウザーによって提供されます。

どのブラウザが使用されているかは、以下によります。

- コンピューターのオペレーティング システム。
- アドインが web、Microsoft 365、Officeサブスクリプション以外のアドインで実行Office 2013かどうか。

次の表は、さまざまなプラットフォームとオペレーティングシステムに使用されているブラウザを示しています。

|OS|Office のバージョン
|Edge WebView2 (Chromium ベース) がインストールされている場合|ブラウザー|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|any|Office on the web|該当なし|Office が開かれているブラウザー。|
|Mac|any|該当なし|Safari|
|iOS|any|該当なし|Safari|
|Android|any|該当なし|Chrome|
|Windows 7、8.1、10 | サブスクリプション以外のOffice 2013以降|かまいません|Internet Explorer 11|
|Windows 7 | Microsoft 365| かまいません | Internet Explorer 11|
|Windows 8.1、<br>Windows 10 ver. &nbsp; < &nbsp;1903| Microsoft 365 | いいえ| Internet Explorer 11|
|Windows 10 ver. &nbsp; >= &nbsp;1903 | Microsoft 365 ver. &nbsp; < &nbsp;16.0.11629<sup>1</sup>| かまいません|Internet Explorer 11|
|Windows 10 ver. &nbsp; >= &nbsp;1903 | Microsoft 365 ver. &nbsp; >= &nbsp;16.0.11629 &nbsp; _および_ &nbsp; < &nbsp; 16.0.13530.20424 <sup>1</sup>| かまいません|Microsoft Edge<sup>2、3</sup> と元の WebView (EdgeHTML)|
|Windows 10 ver. &nbsp; >= &nbsp;1903 | Microsoft 365 ver. &nbsp; >= &nbsp;16.0.13530.20424<sup>1</sup>| いいえ |Microsoft Edge<sup>2、3</sup> と元の WebView (EdgeHTML)|
|Windows 8.1<br>Windows 10| Microsoft 365 ver. &nbsp; >= &nbsp;16.0.13530.20424<sup>1</sup>| は<sup>い 4</sup>|  Microsoft Edge<sup>2、3</sup> と WebView2 (Chromium ベース) |

<sup>1 詳細</sup> については、 [更新履歴ページ](/officeupdates/update-history-office365-proplus-by-date) と、クライアント バージョンとOfficeチャネルを検索 [する方法を](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19) 参照してください。

<sup>2</sup> Microsoft Edge を使用している場合、Windows 10 ナレーター ("スクリーン リーダー" とも呼ばれる) は、作業ウィンドウで開くページ内のタグを読み取 `<title>` ります。 Internet Explorer 11 が使用されている場合、ナレーターはアドイン マニフェストの `<DisplayName>` の値から提供される作業ウィンドウのタイトル バーを読み取ります。

<sup>3</sup> アドインがマニフェストに要素を含む場合、Windows または Microsoft 365 バージョンに関係なく Internet Explorer `Runtimes` 11 を使用します。 詳細については、「[ランタイム](../reference/manifest/runtimes.md)」を参照してください。

<sup>4</sup> 埋め込み可能な WebView2 コントロールは、Microsoft Edge のインストールに加えて、埋め込み可能な WebView2 Officeインストールする必要があります。 インストールするには [、Microsoft Edge WebView2 / Web コンテンツの埋め込みを参照してください。Microsoft Edge WebView2 を使用します](https://developer.microsoft.com/microsoft-edge/webview2/)。


> [!IMPORTANT]
> Internet Explorer 11はES5以降のJavaScriptバージョンをサポートしていません。 アドインのユーザーに Internet Explorer 11 を使用するプラットフォームがある場合、ECMAScript 2015 以降の構文と機能を使用するには、次の 2 つのオプションがあります。
>
> - ECMAScript 2015 (ES6 とも呼ばれる) 以降の JavaScript、または TypeScript でコードを記述し、 [次に、コンパイラ](https://babeljs.io/) (コンパイル時や [tsc](https://www.typescriptlang.org/index.html)など) を使用して ES5 JavaScript にコードをコンパイルします。
> - ECMAScript 2015 以降の JavaScript で記述します[](https://wikipedia.org/wiki/Polyfill_(programming))が、IE でコードを実行できる[core-js](https://github.com/zloirock/core-js)などのポリフィル ライブラリも読み込む。
>
> また、Internet Explorer 11 は、メディア、録音、および位置情報などの HTML 5 機能の一部をサポートしていません。

## <a name="troubleshooting-microsoft-edge-issues"></a>Microsoft Edge の問題のトラブルシューティング

### <a name="service-workers-are-not-working"></a>サービス ワーカーが動作していない

Office Microsoft [Edge WebView](/microsoft-edge/hosting/webview) を使用する場合、アドインはサービス ワーカーをサポートできません。 Chromium ベースの [Edge WebView2 でサポートされています](/microsoft-edge/hosting/webview2)。

### <a name="scroll-bar-does-not-appear-in-task-pane"></a>作業ウィンドウにスクロール バーが表示されない

既定では、Microsoft Edge のスクロール バーはホバーするまで非表示になっています。 スクロールバーが常に表示されるようにするには、作業ウィンドウのページの`<body>`要素に適用される CSS スタイルに [-ms-overflow-style](https://developer.mozilla.org/docs/Archive/Web/CSS/-ms-overflow-style) プロパティを含め、`scrollbar`に設定する必要があります。

### <a name="when-debugging-with-the-microsoft-edge-devtools-the-add-in-crashes-or-reloads"></a>Microsoft Edge DevTools を使用してデバッグすると、アドインがクラッシュまたは再読み込みされる

[Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab) にブレークポイントを設定すると、アドインがハングしていると Office に判断される可能性があります。 これが発生すると、アドインが自動的に再読み込みされます。 これを防ぐには、開発用コンピューターに以下のレジストリ キーと値を追加します: `[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`。

### <a name="when-the-add-in-tries-to-open-get-add-in-error-we-cant-open-this-add-in-from-the-localhost-error"></a>アドインを開こうとすると、"アドイン エラー localhost からこのアドインを開けません" というエラーが表示される

既知の原因の1つとして、Microsoft Edge では開発用コンピューター上では localhost にループバックの除外を与える必要があることが挙げられます。 [Cannot open add-in from localhost (localhostからアドインを開くことができません)](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost)の指示に従ってください。

### <a name="get-errors-trying-to-download-a-pdf-file"></a>PDF ファイルをダウンロードしようとしてエラーを取得する

Edge がブラウザーの場合、アドインで BLOB を PDF ファイルとして直接ダウンロードすることはできません。 回避策として、BLOB を PDF ファイルとしてダウンロードする簡単な Web アプリケーションを作成します。 アドインで、メソッドを呼び `Office.context.ui.openBrowserWindow(url)` 出し、Web アプリケーションの URL を渡します。 これにより、Web アプリケーションがブラウザー ウィンドウの外部で開Office。

## <a name="see-also"></a>関連項目

- [Officeアドインを実行するための要件](requirements-for-running-office-add-ins.md)
