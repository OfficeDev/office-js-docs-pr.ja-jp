---
title: Office アドインによって使用されるブラウザー
description: Office アドインによって使用されるブラウザーをオペレーティング システムおよび Office バージョンが決定する方法を指定します。
ms.date: 09/24/2020
localization_priority: Normal
ms.openlocfilehash: 6fc1661a49bd5ba60a42ab891eee5a640b579feb
ms.sourcegitcommit: 09e1d8ff14b3c09a3eb11c91432c224a539181a4
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/25/2020
ms.locfileid: "48268559"
---
# <a name="browsers-used-by-office-add-ins"></a>Office アドインによって使用されるブラウザー

Office アドインは、デスクトップおよびモバイル クライアント用の Office に埋め込まれたブラウザー コントロールを使用してOffice on the web での実行時に、 iFrame を使用して表示されるWeb アプリケーションです。 アドインには JavaScript を実行するための JavaScript エンジンも必要です。 埋め込まれたブラウザーおよびエンジン、どちらもユーザーのコンピュータにインストールされているブラウザーによって提供されます。

どのブラウザが使用されているかは、以下によります。

- コンピュータのオペレーティングシステム。
- アドインが Office on the web、Office 365、または非サブスクリプションのOffice 2013 以降で実行されているかどうか。

次の表は、さまざまなプラットフォームとオペレーティングシステムに使用されているブラウザを示しています。

|OS|Office バージョン|Edge WebView2 (Chromium ベース) がインストールされていますか?|ブラウザー|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|任意|Office on the web|該当なし|Office が開かれているブラウザー。|
|Mac|任意|該当なし|Safari|
|iOS|任意|該当なし|Safari|
|Android|任意|該当なし|Chrome|
|Windows 7, 8.1, 10 | 非サブスクリプションの Office 2013以降|問題ありません。|Internet Explorer 11|
|Windows 7 | Microsoft 365| 問題ありません。 | Internet Explorer 11|
|Windows 8.1、<br>Windows 10 バージョン&nbsp;<&nbsp;1903| Microsoft 365 | いいえ| Internet Explorer 11|
|Windows 10 バージョン&nbsp;>=&nbsp;1903 | Microsoft 365 バージョン&nbsp;<&nbsp;16.0.11629<sup>1</sup>| 問題ありません。|Internet Explorer 11|
|Windows 10 バージョン&nbsp;>=&nbsp;1903 | Microsoft 365 バージョン&nbsp;>=&nbsp;16.0.11629&nbsp;_AND_&nbsp;<&nbsp;16.0.13127.20082<sup>1</sup>| 問題ありません。|Microsoft Edge<sup>2、3</sup> に元の WebView (EdgeHTML)|
|Windows 10 バージョン&nbsp;>=&nbsp;1903 | Microsoft 365 バージョン&nbsp;>=&nbsp;16.0.13127.20082<sup>1</sup>| いいえ |Microsoft Edge<sup>2、3</sup> に元の WebView (EdgeHTML)|
|Windows 8.1<br>Windows 10| Microsoft 365 バージョン&nbsp;>=&nbsp;16.0.13127.20082<sup>1</sup>| はい<sup>5</sup>|  後述の注4を参照してください。 |

<sup>1</sup>確認いただきたいのは、[更新プログラムの履歴のページ](/officeupdates/update-history-office365-proplus-by-date)で、どのように[Officeクライアントのバージョンおよびチャネルを更新するのかの](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)詳細をご覧ください。

<sup>2</sup> Microsoft Edge が使用されている場合、Windows 10 ナレーター (「スクリーン リーダー」と呼ばれることもあります) は、作業ウィンドウで開いているページの `<title>` タグを読み取ります。 Internet Explorer 11 が使用されている場合、ナレーターはアドイン マニフェストの `<DisplayName>` の値から提供される作業ウィンドウのタイトル バーを読み取ります。

<sup>3</sup>アドインにマニフェストの `Runtimes`要素が含まれている場合、Windows または Microsoft 365 のバージョンに関係なく、Internet Explorer 11 が使用されます。 詳細については、「[ランタイム](../reference/manifest/runtimes.md)」を参照してください。

<sup>4</sup> このバージョンの組み合わせに使用されているブラウザーは、Microsoft 365 サブスクリプションの更新プログラムチャネルに依存します。 ユーザーが [ベータチャネル](https://insider.office.com/join/windows) (以前の Insider ファーストチャネル) を使用している場合、Office では、WebView2 (Chromium ベース) を使用して Microsoft Edge を使用します。 その他のチャネルでは、Office は Microsoft Edge を元の WebView (EdgeHTML) と共に使用します。 その他のチャネルでのWebView2 のサポートは2021年の早い時期に実施が期待されています。 *ノート5*も参照ください。

<sup>5</sup> Office がそれを組み込めるように、組み込みの WebView2 コントロールを Microsoft Edge のインストールに加えてインストールしておく必要があります。 インストールするには、[Microsoft Edge WebView2 (プレビュー) / Microsoft Edge WebView2を使用して、web コンテンツ ...の埋め込み](https://developer.microsoft.com/microsoft-edge/webview2/)を参照してください。


> [!IMPORTANT]
> Internet Explorer 11はES5以降のJavaScriptバージョンをサポートしていません。 アドインのユーザーに Internet Explorer 11 を使用するプラットフォームがある場合、ECMAScript 2015 以降の構文と機能を使用するには、2つの方法があります。
>
> - ECMAScript 2015 (ES6 とも呼ばれます) または後期の JavaScript、あるいは TypeScript でコードを作成するか、あるいは[babel](https://babeljs.io/) や [tsc](https://www.typescriptlang.org/index.html)などのコンパイラーを使用して、コードを ES5 JavaScriptへコンパイルします。
> - ECMAScript 2015 以降の JavaScript で記述します。または、IEがコードを実行できるように[core-js](https://github.com/zloirock/core-js)などの[polyfill](https://wikipedia.org/wiki/Polyfill_(programming))ライブラリを読み込みます。
>
> また、Internet Explorer 11 は、メディア、録音、および位置情報などの HTML 5 機能の一部をサポートしていません。

## <a name="troubleshooting-microsoft-edge-issues"></a>Microsoft Edge の問題のトラブルシューティング

### <a name="service-workers-are-not-working"></a>サービスワーカーが機能しない

元の [Microsoft Edge WebView](/microsoft-edge/hosting/webview) を使用している場合、Office アドインはサービスワーカーをサポートしません。 [Chromium ベースのEdge WebView2](/microsoft-edge/hosting/webview2)を使用すると、サポートされます。

### <a name="scroll-bar-does-not-appear-in-task-pane"></a>作業ウィンドウにスクロール バーが表示されない

既定では、Microsoft Edge のスクロール バーはホバーするまで非表示になっています。 スクロールバーが常に表示されるようにするには、作業ウィンドウのページの`<body>`要素に適用される CSS スタイルに [-ms-overflow-style](https://developer.mozilla.org/docs/Web/CSS/-ms-overflow-style) プロパティを含め、`scrollbar`に設定する必要があります。 

### <a name="when-debugging-with-the-microsoft-edge-devtools-the-add-in-crashes-or-reloads"></a>Microsoft Edge DevTools を使用してデバッグすると、アドインがクラッシュまたは再読み込みされる

[Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab) にブレークポイントを設定すると、アドインがハングしていると Office に判断される可能性があります。 これが発生すると、アドインが自動的に再読み込みされます。 これを防ぐには、開発用コンピューターに以下のレジストリ キーと値を追加します: `[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`。

### <a name="when-the-add-in-tries-to-open-get-add-in-error-we-cant-open-this-add-in-from-the-localhost-error"></a>アドインを開こうとすると、"アドイン エラー localhost からこのアドインを開けません" というエラーが表示される

既知の原因の1つとして、Microsoft Edge では開発用コンピューター上では localhost にループバックの除外を与える必要があることが挙げられます。 [Cannot open add-in from localhost (localhostからアドインを開くことができません)](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost)の指示に従ってください。

### <a name="get-errors-trying-to-download-a-pdf-file"></a>PDF ファイルのダウンロード中にエラーが発生する

Edgeがブラウザーの場合は、blob をアドインのPDF ファイルとして直接ダウンロードすることはできません。 回避策として、blob を PDF ファイルとしてダウンロードする簡単な web アプリケーションを作成します。 アドインで、 `Office.context.ui.openBrowserWindow(url)` メソッドを呼び出し、web アプリケーションの URL を渡します。 この操作を行うと、Office の外のブラウザーウィンドウでweb アプリケーションが開きます。

## <a name="see-also"></a>関連項目

- [Officeアドインを実行するための要件](requirements-for-running-office-add-ins.md)
