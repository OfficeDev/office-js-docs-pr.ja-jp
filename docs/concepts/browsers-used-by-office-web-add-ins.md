---
title: Office アドインによって使用されるブラウザー
description: Office アドインによって使用されるブラウザーをオペレーティング システムおよび Office バージョンが決定する方法を指定します。
ms.date: 09/29/2022
ms.localizationpriority: medium
ms.openlocfilehash: a75cab613605760e774f8b2a163172e4ec6cb5bd
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810156"
---
# <a name="browsers-used-by-office-add-ins"></a>Office アドインによって使用されるブラウザー

Office アドインは、Office on the webで実行するときに iFrames を使用して表示される Web アプリケーションです。 Office for デスクトップおよびモバイル クライアントでは、Office アドインは埋め込みブラウザー コントロール (Webview とも呼ばれます) を使用します。 アドインには JavaScript を実行するための JavaScript エンジンも必要です。 埋め込みブラウザーとエンジンの両方が、ユーザーのコンピューターにインストールされているブラウザーによって提供されます。

どのブラウザが使用されているかは、以下によります。

- コンピューターのオペレーティング システム。
- アドインがOffice on the webで実行されているか、Office で Microsoft 365 サブスクリプションからダウンロードされているか、永続的な Office 2013 以降で実行されているか。
- Windows 上の Office の永続的なバージョン内で、アドインが "リテール" または "ボリューム ライセンス" のバリエーションで実行されているかどうか。

> [!NOTE]
> この記事では、[Windows Information Protection (WIP)](/windows/uwp/enterprise/wip-hub) で保護 *されていない* ドキュメントでアドインが実行されていることを前提としています。 WIP で保護されたドキュメントの場合、この記事の情報にはいくつかの例外があります。 詳細については、「 [WIP で保護されたドキュメント](#wip-protected-documents)」を参照してください。

> [!IMPORTANT]
> **Office アドインで引き続き使用される Internet Explorer**
>
> この記事で説明するように、Office 2019 を通じてボリューム ライセンスの永続的なバージョンを含む、プラットフォームと Office バージョンの組み合わせによっては、Internet Explorer 11 に付属する Webview コントロールを引き続き使用してアドインをホストします。 Internet Explorer Webview でアドインを起動したときに、アドインのユーザーに正常なエラー メッセージを提供することで、少なくとも最小限の方法で、これらの組み合わせを引き続きサポートすることをお勧めします (ただし、必要ありません)。 次の点に注意してください。
>
> - Internet Explorer でOffice on the webが開かなくなりました。 そのため、[AppSource](/office/dev/store/submit-to-appsource-via-partner-center) は、ブラウザーとして Internet Explorer を使用してOffice on the webアドインをテストしなくなりました。
> - AppSource は引き続き、Internet Explorer を使用するプラットフォームと Office *デスクトップ* バージョンの組み合わせをテストします。 ただし、アドインが Internet Explorer をサポートしていない場合にのみ警告が発行されます。アドインは AppSource によって拒否されません。
> - [Script Lab ツール](../overview/explore-with-script-lab.md)は Internet Explorer をサポートしなくなりました。
>
> Internet Explorer のサポートとアドインでの正常なエラー メッセージの構成の詳細については、「 [Internet Explorer 11 のサポート](../develop/support-ie-11.md)」を参照してください。

次のセクションでは、さまざまなプラットフォームとオペレーティング システムに使用されるブラウザーを指定します。

## <a name="non-windows-platforms"></a>Windows 以外のプラットフォーム

これらのプラットフォームでは、使用されるブラウザーがプラットフォームによって決定されます。

|OS|Office のバージョン
|ブラウザー|
|:-----|:-----|:-----|
|any|Office on the web|Office が開かれているブラウザー。<br>(ただし、Internet Explorer ではOffice on the webが開かないことに注意してください。<br>これを行おうとすると、Edge でOffice on the webが開きます)。 |
|Mac|any|WKWebView を使用した Safari|
|iOS|any|WKWebView を使用した Safari|
|Android|any|Chrome|

## <a name="perpetual-versions-of-office-on-windows"></a>Windows 上の Office の永続的なバージョン

Windows 上の Office の永続的なバージョンの場合、使用されるブラウザーは、Office のバージョン、ライセンスが製品版かボリューム ライセンスか、Edge WebView2 (Chromium ベース) がインストールされているかどうかによって決まります。 Windows のバージョンは関係ありませんが、Office Web アドインは Windows 7 より前のバージョンではサポートされておらず、Windows 10より前のバージョンではOffice 2021はサポートされていないことに注意してください。

Office 2016 または Office 2019 がリテール版かボリューム ライセンス版かを判断するには、Office のバージョンとビルド番号の形式を使用します。 (Office 2013 と Office 2021の場合、ボリューム ライセンスと小売りの違いは関係ありません)。

- **リテール**: Office 2016 と 2019 の両方の形式は `YYMM (xxxxx.xxxxxx)`で、5 桁の 2 ブロックで終わるです(例: `2206 (Build 15330.20264`)。
- **ボリューム ライセンス**:
  - Office 2016 の場合、形式は です。末尾は `16.0.xxxx.xxxxx`*4* 桁の 2 ブロックです(例: `16.0.5197.1000`)。
  - Office 2019 の場合、形式は です。末尾は `1808 (xxxxx.xxxxxx)`*5* 桁の 2 ブロックです(例: `1808 (Build 10388.20027)`)。 年と月は常に であることに注意してください `1808`。

| Office のバージョン
 | リテールライセンスとボリュームライセンス | Edge WebView2 (Chromium ベース) がインストールされていますか? | ブラウザー |
|:-----|:-----|:-----|:-----|
| Office 2013 | かまいません | かまいません | Internet Explorer 11 |
| Office 2016 | ボリューム ライセンス | かまいません | Internet Explorer 11 |
| Office 2019 | ボリューム ライセンス | かまいません | Internet Explorer 11 |
| Office 2016 から Office 2019 | 小売 | いいえ | 元の WebView (EdgeHTML) を使用した Microsoft Edge<sup>1、2</sup></br>Edge がインストールされていない場合は、Internet Explorer 11 が使用されます。 |
| Office 2016 から Office 2019 | 小売 | はい<sup>3</sup> | WebView2 を使用した Microsoft Edge<sup>1</sup> (Chromium ベース) |
| Office 2021 | かまいません | はい<sup>3</sup> | WebView2 を使用した Microsoft Edge<sup>1</sup> (Chromium ベース) |

<sup>1</sup> Microsoft Edge を使用すると、Windows ナレーター ("スクリーン リーダー" とも呼ばれます) は、作業ウィンドウで開くページのタグを読み取 `<title>` ります。 Internet Explorer 11 では、ナレーターは作業ウィンドウのタイトル バーを読み取ります。これはアドインのマニフェストの値から **\<DisplayName\>** 取得されます。

<sup>2</sup> アドインにマニフェストに 要素が **\<Runtimes\>** 含まれている場合、元の WebView (EdgeHTML) で Microsoft Edge は使用されません。 WebView2 (Chromium ベース) で Microsoft Edge を使用するための条件が満たされている場合、アドインはそのブラウザーを使用します。 それ以外の場合は、Internet Explorer 11 を使用します。 詳細については、「[ランタイム](/javascript/api/manifest/runtimes)」を参照してください。

<sup>3</sup> Windows 11より前の Windows バージョンでは、Office が埋め込むことができるように WebView2 コントロールをインストールする必要があります。 永続的なOffice 2021以降でインストールされますが、Microsoft Edge では自動的にインストールされません。 以前のバージョンの永続的な Office がある場合は、 [Microsoft Edge WebView2/ Embed Web コンテンツ ... にコントロールをインストールする手順を使用します。Microsoft Edge WebView2](https://developer.microsoft.com/microsoft-edge/webview2/) を使用します。

## <a name="microsoft-365-subscription-versions-of-office-on-windows"></a>Windows 上の Office の Microsoft 365 サブスクリプション バージョン

サブスクリプション Office on Windows の場合、使用されるブラウザーは、オペレーティング システム、Office のバージョン、および Edge WebView2 (Chromium ベース) がインストールされているかどうかによって決まります。

|OS|Office のバージョン
|Edge WebView2 (Chromium ベース) がインストールされていますか?|ブラウザー|
|:-----|:-----|:-----|:-----|
|Windows 7 | Microsoft 365| かまいません | Internet Explorer 11|
|Windows 8.1,<br>Windows 10 ver.&nbsp;<&nbsp;1903| Microsoft 365 | いいえ| Internet Explorer 11|
|Windows 10 ver.&nbsp;>=&nbsp;1903,<br>Windows 11 | Microsoft 365 ver.&nbsp;<&nbsp;16.0.11629<sup>2</sup>| かまいません|Internet Explorer 11|
|Windows 10 ver.&nbsp;>=&nbsp;1903,<br>Windows 11 | Microsoft 365 ver.&nbsp;>=&nbsp;16.0.11629&nbsp;_と_&nbsp;<&nbsp;16.0.13530.20424 <sup>2</sup>| かまいません|元の WebView を使用した Microsoft Edge<sup>1、3</sup> (EdgeHTML)|
|Windows 10 ver.&nbsp;>=&nbsp;1903,<br>ウィンドウ 11 | Microsoft 365 ver.&nbsp;>=&nbsp;16.0.13530.20424<sup>2</sup>| いいえ |元の WebView を使用した Microsoft Edge<sup>1、3</sup> (EdgeHTML)|
|Windows 8.1<br>Windows 10、<br>Windows 11| Microsoft 365 ver.&nbsp;>=&nbsp;16.0.13530.20424<sup>2</sup>| はい<sup>4</sup>|  WebView2 を使用した Microsoft Edge<sup>1</sup> (Chromium ベース) |

<sup>1</sup> Microsoft Edge を使用すると、Windows ナレーター ("スクリーン リーダー" とも呼ばれます) は、作業ウィンドウで開くページのタグを読み取 `<title>` ります。 Internet Explorer 11 では、ナレーターは作業ウィンドウのタイトル バーを読み取ります。これはアドインのマニフェストの値から **\<DisplayName\>** 取得されます。

<sup>2</sup> 詳細については、 [更新履歴ページ](/officeupdates/update-history-office365-proplus-by-date) と [、Office クライアントのバージョンと更新チャネルを見つける](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19) 方法を参照してください。

<sup>3</sup> アドインにマニフェストに 要素が **\<Runtimes\>** 含まれている場合、元の WebView (EdgeHTML) で Microsoft Edge は使用されません。 WebView2 (Chromium ベース) で Microsoft Edge を使用するための条件が満たされている場合、アドインはそのブラウザーを使用します。 それ以外の場合は、Windows または Microsoft 365 のバージョンに関係なく、Internet Explorer 11 が使用されます。 詳細については、「[ランタイム](/javascript/api/manifest/runtimes)」を参照してください。

<sup>4</sup> Windows 11より前の Windows バージョンでは、Office が埋め込むことができるように WebView2 コントロールをインストールする必要があります。 Microsoft 365 バージョン 2101 以降でインストールされますが、Microsoft Edge では自動的にインストールされません。 以前のバージョンの Microsoft 365 をお持ちの場合は、 [Microsoft Edge WebView2 /埋め込み Web コンテンツ ... にコントロールをインストールする手順を使用します。Microsoft Edge WebView2](https://developer.microsoft.com/microsoft-edge/webview2/) を使用します。 16.0.14326.xxxxx より前の Microsoft 365 ビルドでは、レジストリ キー **HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Win32WebView2** を作成し、その値を に設定する `dword:00000001`必要もあります。

## <a name="working-with-internet-explorer"></a>Internet Explorer の操作

Internet Explorer 11はES5以降のJavaScriptバージョンをサポートしていません。 アドインのユーザーに Internet Explorer 11 を使用するプラットフォームがある場合、ECMAScript 2015 以降の構文と機能を使用するには、2 つのオプションがあります。

- ECMAScript 2015 (ES6 とも呼ばれます) 以降の JavaScript または TypeScript でコードを記述し、 [babel](https://babeljs.io/) や [tsc](https://www.typescriptlang.org/index.html) などのコンパイラを使用してコードを ES5 JavaScript にコンパイルします。
- ECMAScript 2015 以降の JavaScript で記述しますが、[CORE-js](https://github.com/zloirock/core-js) などの[ポリフィル](https://en.wikipedia.org/wiki/Polyfill_(programming)) ライブラリも読み込み、IE でコードを実行できるようにします。

これらのオプションの詳細については、「 [Internet Explorer 11 のサポート](../develop/support-ie-11.md)」を参照してください。

また、Internet Explorer 11 は、メディア、録音、および位置情報などの HTML 5 機能の一部をサポートしていません。 詳細については、「アドイン [が Internet Explorer で実行されているかどうかを実行時に判断](../develop/support-ie-11.md#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer)する」を参照してください。

## <a name="troubleshoot-microsoft-edge-issues"></a>Microsoft Edge の問題のトラブルシューティング

### <a name="service-workers-are-not-working"></a>Service Worker が機能していない

Office アドインでは、元の Microsoft Edge WebView [EdgeHTML](https://en.wikipedia.org/wiki/EdgeHTML) が使用されている場合、Service Worker はサポートされません。 [これらは、Chromium ベースの Edge WebView2](/microsoft-edge/hosting/webview2) でサポートされています。

### <a name="scroll-bar-does-not-appear-in-task-pane"></a>作業ウィンドウにスクロール バーが表示されない

既定では、Microsoft Edge のスクロール バーはホバーするまで非表示になっています。 スクロールバーが常に表示されるようにするには、作業ウィンドウのページの`<body>`要素に適用される CSS スタイルに [-ms-overflow-style](https://developer.mozilla.org/docs/Web/CSS/Microsoft_Extensions) プロパティを含め、`scrollbar`に設定する必要があります。

### <a name="when-debugging-with-the-microsoft-edge-devtools-the-add-in-crashes-or-reloads"></a>Microsoft Edge DevTools を使用してデバッグすると、アドインがクラッシュまたは再読み込みされる

[Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab) にブレークポイントを設定すると、アドインがハングしていると Office に判断される可能性があります。 これが発生すると、アドインが自動的に再読み込みされます。 これを防ぐには、開発用コンピューターに以下のレジストリ キーと値を追加します: `[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`。

### <a name="when-the-add-in-tries-to-open-get-add-in-error-we-cant-open-this-add-in-from-the-localhost-error"></a>アドインを開こうとすると、"アドイン エラー localhost からこのアドインを開けません" というエラーが表示される

既知の原因の1つとして、Microsoft Edge では開発用コンピューター上では localhost にループバックの除外を与える必要があることが挙げられます。 [Cannot open add-in from localhost (localhostからアドインを開くことができません)](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost)の指示に従ってください。

### <a name="get-errors-trying-to-download-a-pdf-file"></a>PDF ファイルのダウンロード中にエラーが発生する

Edge がブラウザーの場合、アドインで BLOB を PDF ファイルとして直接ダウンロードすることはできません。 回避策は、BLOB を PDF ファイルとしてダウンロードする単純な Web アプリケーションを作成することです。 アドインで、 メソッドを `Office.context.ui.openBrowserWindow(url)` 呼び出し、Web アプリケーションの URL を渡します。 これにより、Office の外部のブラウザー ウィンドウで Web アプリケーションが開きます。

## <a name="wip-protected-documents"></a>WIP で保護されたドキュメント

[WIP で保護された](/windows/uwp/enterprise/wip-hub)ドキュメントで実行されているアドインでは、**WebView2 (Chromium ベース) で Microsoft Edge** が使用されることはありません。 この記事の「[Windows 上の Office の永続的なバージョン](#perpetual-versions-of-office-on-windows)」および「Windows [上の Office の Microsoft 365 サブスクリプション バージョン](#microsoft-365-subscription-versions-of-office-on-windows)」のセクションでは、**Microsoft Edge を WebView2 (Chromium ベース) の元の WebView (EdgeHTML)** に置き換えます。

ドキュメントが WIP で保護されているかどうかを判断するには、次の手順に従います。

1. ファイルを開きます。
1. リボンの [ **ファイル** ] タブを選択します。
1. [ **情報] を選択します**。
1. [ **情報** ] ページの左上のファイル名のすぐ下に、WIP 対応ドキュメントのブリーフケース アイコンが続き、 **Managed by Work (...) が** 表示されます。

## <a name="see-also"></a>関連項目

- [Officeアドインを実行するための要件](requirements-for-running-office-add-ins.md)
- [Office アドインのランタイム](../testing/runtimes.md)
