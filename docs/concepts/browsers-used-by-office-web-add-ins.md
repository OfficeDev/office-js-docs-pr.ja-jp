---
title: Office アドインによって使用されるブラウザー
description: Office アドインによって使用されるブラウザーをオペレーティング システムおよび Office バージョンが決定する方法を指定します。
ms.date: 05/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1fedeb7f7e1e972a2a7fe4befa5a990ff8cc698d
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/06/2022
ms.locfileid: "66659655"
---
# <a name="browsers-used-by-office-add-ins"></a>Office アドインによって使用されるブラウザー

Office アドインは、Office on the webで実行するときに iFrame を使用して表示される Web アプリケーションです。 Office for desktop クライアントとモバイル クライアントでは、Office アドインは埋め込みブラウザー コントロール (Webview とも呼ばれます) を使用します。 アドインには JavaScript を実行するための JavaScript エンジンも必要です。 埋め込みブラウザーとエンジンの両方が、ユーザーのコンピューターにインストールされているブラウザーによって提供されます。

どのブラウザが使用されているかは、以下によります。

- コンピューターのオペレーティング システム。
- アドインがOffice on the web、Microsoft 365、またはサブスクリプション外の Office 2013 以降で実行されているかどうか。

> [!IMPORTANT]
> **Office アドインで引き続き使用される Internet Explorer**
>
> Office 2019 を通じて 1 回限りの購入バージョンを含む一部のプラットフォームと Office バージョンの組み合わせでは、この記事で説明されているように、Internet Explorer 11 に付属する Web ビュー コントロールを引き続き使用してアドインをホストします。 Internet Explorer Webview でアドインを起動したときにアドインのユーザーに正常なエラー メッセージを提供することで、少なくとも最小限の方法でこれらの組み合わせを引き続きサポートすることをお勧めします (ただし、必要ありません)。 次の点に注意してください。
>
> - Internet Explorer でOffice on the webが開かなくなりました。 その結果、[AppSource は](/office/dev/store/submit-to-appsource-via-partner-center)、ブラウザーとして Internet Explorer を使用してOffice on the webでアドインをテストしなくなりました。
> - AppSource は引き続き Internet Explorer を使用するプラットフォームと Office *デスクトップ* バージョンの組み合わせをテストしますが、アドインが Internet Explorer をサポートしていない場合にのみ警告が発行されます。アドインは AppSource によって拒否されません。
> - [Script Lab ツール](../overview/explore-with-script-lab.md)は Internet Explorer をサポートしなくなりました。
>
> Internet Explorer のサポートとアドインでの正常なエラー メッセージの構成の詳細については、 [Internet Explorer 11 のサポート](../develop/support-ie-11.md)に関するページを参照してください。

次の表は、さまざまなプラットフォームとオペレーティングシステムに使用されているブラウザを示しています。

|OS|Office のバージョン
|Edge WebView2 (Chromium ベース) がインストールされていますか?|ブラウザー|
|:-----|:-----|:-----|:-----|
|any|Office on the web|×|Office が開かれているブラウザー。<br>(ただし、Office on the webは Internet Explorer では開かないことに注意してください。<br>これを行おうとすると、Edge でOffice on the web開きます)。 |
|Mac|any|×|WKWebView を使用した Safari|
|iOS|any|×|WKWebView を使用した Safari|
|Android|any|×|Chrome|
|Windows 7、8.1、10、11 | サブスクリプション外の Office 2013 から Office 2019|かまいません|Internet Explorer 11|
|Windows 10, 11 | サブスクリプション以外のOffice 2021以降|はい|<sup>WebView2</sup> を使用した Microsoft Edge 1 (Chromium ベース)|
|Windows 7 | Microsoft 365| かまいません | Internet Explorer 11|
|Windows 8.1、<br>Windows 10 ver.&nbsp;<&nbsp;1903| Microsoft 365 | 不要| Internet Explorer 11|
|Windows 10 ver.&nbsp;>=&nbsp;1903,<br>Windows 11 | Microsoft 365 ver.&nbsp;<&nbsp;16.0.11629<sup>2</sup>| かまいません|Internet Explorer 11|
|Windows 10 ver.&nbsp;>=&nbsp;1903,<br>Windows 11 | Microsoft 365 ver.&nbsp;>=&nbsp;16.0.11629&nbsp;_AND_&nbsp;<&nbsp;16.0.13530.20424 <sup>2</sup>| かまいません|Microsoft Edge<sup>1、3 と</sup> 元の WebView (EdgeHTML)|
|Windows 10 ver.&nbsp;>=&nbsp;1903,<br>ウィンドウ 11 | Microsoft 365 ver.&nbsp;>=&nbsp;16.0.13530.20424<sup>2</sup>| 不要 |Microsoft Edge<sup>1、3 と</sup> 元の WebView (EdgeHTML)|
|Windows 8.1<br>Windows 10、<br>Windows 11| Microsoft 365 ver.&nbsp;>=&nbsp;16.0.13530.20424<sup>2</sup>| はい<sup>4</sup>|  <sup>WebView2</sup> を使用した Microsoft Edge 1 (Chromium ベース) |

<sup>1</sup> Microsoft Edge を使用している場合、Windows ナレーター ("スクリーン リーダー" とも呼ばれます) は、作業ウィンドウで開いたページでタグを読み取ります `<title>` 。 Internet Explorer 11 が使用されている場合、ナレーターは作業ウィンドウのタイトル バーを読み取ります。これは、アドインのマニフェストの値から **\<DisplayName\>** 取得されます。

<sup>2</sup> 詳細については、 [更新履歴ページ](/officeupdates/update-history-office365-proplus-by-date) と [Office クライアントのバージョンと更新チャネルを見つける](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19) 方法を参照してください。

<sup>3</sup> アドインにマニフェスト内の要素が **\<Runtimes\>** 含まれている場合、元の WebView (EdgeHTML) で Microsoft Edge は使用されません。 WebView2 で Microsoft Edge を使用するための条件 (Chromium ベース) が満たされている場合、アドインはそのブラウザーを使用します。 それ以外の場合は、Windows または Microsoft 365 バージョンに関係なく Internet Explorer 11 が使用されます。 詳細については、「[ランタイム](/javascript/api/manifest/runtimes)」を参照してください。

<sup>4</sup> Windows 11より前の Windows バージョンでは、Office が埋め込むことができるように WebView2 コントロールをインストールする必要があります。 Microsoft 365 バージョン 2101 以降、および 1 回限りの購入Office 2021以降でインストールされますが、Microsoft Edge では自動的にインストールされません。 以前のバージョンの Microsoft 365 または 1 回限りの購入 Office をお持ちの場合は、 [Microsoft Edge WebView2 /Embed Web コンテンツ ... でコントロールをインストールする手順を使用します。Microsoft Edge WebView2 を使用](https://developer.microsoft.com/microsoft-edge/webview2/)します。 16.0.14326.xxxxx より前の Microsoft 365 ビルドでは、レジストリ キー **HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Win32WebView2** を作成し、その値 `dword:00000001`を .

> [!IMPORTANT]
> Internet Explorer 11はES5以降のJavaScriptバージョンをサポートしていません。 アドインのユーザーのいずれかが Internet Explorer 11 を使用するプラットフォームを持っている場合は、ECMAScript 2015 以降の構文と機能を使用するには、2 つのオプションがあります。
>
> - ECMAScript 2015 (ES6 とも呼ばれます) または TypeScript でコードを記述し、 [バベル](https://babeljs.io/) や [tsc](https://www.typescriptlang.org/index.html) などのコンパイラを使用してコードを ES5 JavaScript にコンパイルします。
> - ECMAScript 2015 以降の JavaScript で記述しますが、IE でコードを実行できるようにする [core-js](https://github.com/zloirock/core-js) などの[ポリフィル](https://en.wikipedia.org/wiki/Polyfill_(programming)) ライブラリも読み込みます。
>
> これらのオプションの詳細については、 [Internet Explorer 11 のサポートに関するページを](../develop/support-ie-11.md)参照してください。
>
> また、Internet Explorer 11 は、メディア、録音、および位置情報などの HTML 5 機能の一部をサポートしていません。 詳細については、「 [Internet Explorer でアドインが実行されているかどうかを実行時に確認](../develop/support-ie-11.md#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer)する」を参照してください。

## <a name="troubleshooting-microsoft-edge-issues"></a>Microsoft Edge の問題のトラブルシューティング

### <a name="service-workers-are-not-working"></a>Service Worker が機能していない

Office アドインでは、元の Microsoft Edge WebView [EdgeHTML](https://en.wikipedia.org/wiki/EdgeHTML) が使用されている場合、Service Worker はサポートされません。 [これらは、Chromium ベースの Edge WebView2](/microsoft-edge/hosting/webview2) でサポートされています。

### <a name="scroll-bar-does-not-appear-in-task-pane"></a>作業ウィンドウにスクロール バーが表示されない

既定では、Microsoft Edge のスクロール バーはホバーするまで非表示になっています。 スクロールバーが常に表示されるようにするには、作業ウィンドウのページの`<body>`要素に適用される CSS スタイルに [-ms-overflow-style](https://developer.mozilla.org/docs/Web/CSS/Microsoft_Extensions) プロパティを含め、`scrollbar`に設定する必要があります。

### <a name="when-debugging-with-the-microsoft-edge-devtools-the-add-in-crashes-or-reloads"></a>Microsoft Edge DevTools を使用してデバッグすると、アドインがクラッシュまたは再読み込みされる

[Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab) にブレークポイントを設定すると、アドインがハングしていると Office に判断される可能性があります。 これが発生すると、アドインが自動的に再読み込みされます。 これを防ぐには、開発用コンピューターに以下のレジストリ キーと値を追加します: `[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`。

### <a name="when-the-add-in-tries-to-open-get-add-in-error-we-cant-open-this-add-in-from-the-localhost-error"></a>アドインを開こうとすると、"アドイン エラー localhost からこのアドインを開けません" というエラーが表示される

既知の原因の1つとして、Microsoft Edge では開発用コンピューター上では localhost にループバックの除外を与える必要があることが挙げられます。 [Cannot open add-in from localhost (localhostからアドインを開くことができません)](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost)の指示に従ってください。

### <a name="get-errors-trying-to-download-a-pdf-file"></a>PDF ファイルのダウンロード中にエラーが発生する

Edge がブラウザーの場合、アドイン内の PDF ファイルとして BLOB を直接ダウンロードすることはできません。 回避策は、PDF ファイルとして BLOB をダウンロードする単純な Web アプリケーションを作成することです。 アドインでメソッドを呼び出 `Office.context.ui.openBrowserWindow(url)` し、Web アプリケーションの URL を渡します。 これにより、Office の外部のブラウザー ウィンドウで Web アプリケーションが開きます。

## <a name="see-also"></a>関連項目

- [Officeアドインを実行するための要件](requirements-for-running-office-add-ins.md)
