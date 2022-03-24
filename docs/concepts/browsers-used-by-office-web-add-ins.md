---
title: Office アドインによって使用されるブラウザー
description: Office アドインによって使用されるブラウザーをオペレーティング システムおよび Office バージョンが決定する方法を指定します。
ms.date: 10/22/2021
ms.localizationpriority: medium
ms.openlocfilehash: 2ff3bd07ff10e46705ac9a23139bf3cafaf7ef8b
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63742836"
---
# <a name="browsers-used-by-office-add-ins"></a>Office アドインによって使用されるブラウザー

Officeアドインは、iFrames を使用して表示される Web アプリケーションです。Office on the web。 デスクトップOfficeモバイル クライアントの場合、Officeアドインは埋め込みブラウザー コントロール (Webview とも呼ばれる) を使用します。 アドインには JavaScript を実行するための JavaScript エンジンも必要です。 埋め込みブラウザーとエンジンの両方が、ユーザーのコンピューターにインストールされているブラウザーによって提供されます。

どのブラウザが使用されているかは、以下によります。

- コンピューターのオペレーティング システム。
- アドインが 2013 以降の Office on the web、Microsoft 365、またはサブスクリプション以外で実行Officeかどうか。

> [!IMPORTANT]
> **Internet Explorerアドインで引き続きOffice使用される場合**
>
> Microsoft は、アドインのサポートInternet Explorer終了していますが、これはアドインのOffice大きな影響を及ぼします。この記事で説明したように、プラットフォームと Office バージョンの組み合わせ (Office 2019 年 1 回限り購入バージョンを含む) は、Internet Explorer 11 に付属する webview コントロールを引き続き使用してアドインをホストします。 さらに、これらの組み合わせに対するサポートは、Internet Explorerに提出されたアドインに対して引き続き [必要です](/office/dev/store/submit-to-appsource-via-partner-center)。 次の *2 つの点が変化* しています。
>
> - Office on the webで開かなくなったInternet Explorer。 そのため、AppSource はブラウザーとしてアプリケーションを使用してOffice on the webアドインInternet Explorerテストしなくなりました。 ただし、AppSource は引き続き、プラットフォームとデスクトップ バージョンの組み合わせOffice *を* テストInternet Explorer。
> - この[Script Labツールは](../overview/explore-with-script-lab.md)、この機能をサポートInternet Explorer。

次の表は、さまざまなプラットフォームとオペレーティングシステムに使用されているブラウザを示しています。

|OS|Office のバージョン
|Edge WebView2 (Chromiumベース) がインストールされていますか?|ブラウザー|
|:-----|:-----|:-----|:-----|
|any|Office on the web|該当なし|Office が開かれているブラウザー。<br>(ただし、Office on the webは開かInternet Explorer。<br>そうしようとすると、Edge でOffice on the web開きます)。 |
|Mac|any|該当なし|Safari|
|iOS|any|該当なし|Safari|
|Android|any|該当なし|Chrome|
|Windows 7、8.1、10、11 | サブスクリプション以外のOffice 2013 ~ Office 2019|かまいません|Internet Explorer 11|
|Windows 10, 11 | サブスクリプション以外のOffice 2021 以降|はい|Microsoft Edge <sup>1</sup> with WebView2 (Chromiumベース)|
|Windows 7 | Microsoft 365| かまいません | Internet Explorer 11|
|Windows 8.1、<br>Windows 10 ver.&nbsp;<&nbsp;1903| Microsoft 365 | 不要| Internet Explorer 11|
|Windows 10 ver.&nbsp;>=&nbsp;1903,<br>Windows 11 | Microsoft 365 ver.&nbsp;<&nbsp;16.0.116292<sup></sup>| かまいません|Internet Explorer 11|
|Windows 10 ver.&nbsp;>=&nbsp;1903,<br>Windows 11 | Microsoft 365 ver.&nbsp;>=&nbsp;16.0.11629AND16.0.13530.204242&nbsp;&nbsp;<sup></sup><&nbsp;| かまいません|Microsoft Edge <sup>1、3</sup>、元の WebView (EdgeHTML)|
|Windows 10 ver.&nbsp;>=&nbsp;1903,<br>ウィンドウ 11 | Microsoft 365 ver.&nbsp;>=&nbsp;16.0.13530.204242<sup></sup>| 不要 |Microsoft Edge <sup>1、3</sup>、元の WebView (EdgeHTML)|
|Windows 8.1<br>Windows 10、<br>Windows 11| Microsoft 365 ver.&nbsp;>=&nbsp;16.0.13530.204242<sup></sup>| はい<sup>4</sup>|  Microsoft Edge <sup>1</sup> with WebView2 (Chromiumベース) |

<sup>1</sup> Microsoft Edgeされている場合、Windows ナレーター ("スクリーン リーダー" とも呼ばれる) `<title>` は、作業ウィンドウで開くページ内のタグを読み取ります。 Internet Explorer 11 が使用されている場合、ナレーターはアドイン マニフェストの `<DisplayName>` の値から提供される作業ウィンドウのタイトル バーを読み取ります。

<sup>2 詳細</sup>については、「[更新履歴](/officeupdates/update-history-office365-proplus-by-date)」ページと、「クライアント バージョンと更新Officeを見つける[方法」を](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19)参照してください。

<sup>3</sup> アドインにマニフェスト`<Runtimes>`に要素が含まれる場合、元の WebView (EdgeHTML) Microsoft Edgeを使用しない。 WebView2 で使用するMicrosoft Edge (Chromiumベース) が満たされている場合、アドインはそのブラウザーを使用します。 それ以外の場合は、Internet Explorerバージョンに関係なく、Windows 11 Microsoft 365します。 詳細については、「[ランタイム](../reference/manifest/runtimes.md)」を参照してください。

<sup>4</sup> Windows 11 よりWindows前のバージョンでは、WebView2 コントロールをインストールして、Office埋め込む必要があります。 Microsoft 365 バージョン 2101 以降、および 2021 以降の 1 回購入 Officeでインストールされますが、Microsoft Edge では自動的にはインストールされません。 以前のバージョンの Microsoft 365 または 1 回の購入 Office がある場合は、Microsoft Edge [WebView2 / Embed Web content ..でコントロールをインストールする手順を使用します。WebView2 Microsoft Edgeを使用します](https://developer.microsoft.com/microsoft-edge/webview2/)。 Microsoft 365 16.0.14326.xxxxx より前のビルドでは、レジストリ キーを作成し、HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Win32WebView2に設定する必要があります。**** `dword:00000001`

> [!IMPORTANT]
> Internet Explorer 11はES5以降のJavaScriptバージョンをサポートしていません。 アドインのユーザーの中に Internet Explorer 11 を使用するプラットフォームがある場合、ECMAScript 2015 以降の構文と機能を使用するには、2 つのオプションがあります。
>
> - ECMAScript 2015 (ES6 とも呼ばれる) 以降の JavaScript または TypeScript でコードを記述し、バベルや [tsc](https://www.typescriptlang.org/index.html) などのコンパイラを使用してコードを ES5 JavaScript にコンパイル[](https://babeljs.io/)します。
> - ECMAScript 2015 以降の JavaScript で記述しますが、IE でコード[](https://en.wikipedia.org/wiki/Polyfill_(programming))を実行できる [core-js](https://github.com/zloirock/core-js) などのポリフィル ライブラリも読み込む必要があります。
>
> これらのオプションの詳細については、「Support [Internet Explorer 11」を参照してください](../develop/support-ie-11.md)。
>
> また、Internet Explorer 11 は、メディア、録音、および位置情報などの HTML 5 機能の一部をサポートしていません。 詳細については、「アドインが実行中かどうかを実行時に確認する」を参照[Internet Explorer。](../develop/support-ie-11.md#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer)

## <a name="troubleshooting-microsoft-edge-issues"></a>問題Microsoft Edgeトラブルシューティング

### <a name="service-workers-are-not-working"></a>サービス ワーカーが動作していない

Office WebView [EdgeHTML](https://en.wikipedia.org/wiki/EdgeHTML) を使用する場合、アドインはサービス ワーカー Microsoft Edgeサポートされません。 これらは、Chromium [Edge WebView2 でサポートされています](/microsoft-edge/hosting/webview2)。

### <a name="scroll-bar-does-not-appear-in-task-pane"></a>作業ウィンドウにスクロール バーが表示されない

既定では、Microsoft Edge のスクロール バーはホバーするまで非表示になっています。 スクロールバーが常に表示されるようにするには、作業ウィンドウのページの`<body>`要素に適用される CSS スタイルに [-ms-overflow-style](https://developer.mozilla.org/docs/Web/CSS/Microsoft_Extensions) プロパティを含め、`scrollbar`に設定する必要があります。

### <a name="when-debugging-with-the-microsoft-edge-devtools-the-add-in-crashes-or-reloads"></a>Microsoft Edge DevTools を使用してデバッグすると、アドインがクラッシュまたは再読み込みされる

[Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab) にブレークポイントを設定すると、アドインがハングしていると Office に判断される可能性があります。 これが発生すると、アドインが自動的に再読み込みされます。 これを防ぐには、開発用コンピューターに以下のレジストリ キーと値を追加します: `[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`。

### <a name="when-the-add-in-tries-to-open-get-add-in-error-we-cant-open-this-add-in-from-the-localhost-error"></a>アドインを開こうとすると、"アドイン エラー localhost からこのアドインを開けません" というエラーが表示される

既知の原因の1つとして、Microsoft Edge では開発用コンピューター上では localhost にループバックの除外を与える必要があることが挙げられます。 [Cannot open add-in from localhost (localhostからアドインを開くことができません)](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost)の指示に従ってください。

### <a name="get-errors-trying-to-download-a-pdf-file"></a>PDF ファイルをダウンロードしようとしてエラーを取得する

エッジがブラウザーの場合、アドインで BLOB を PDF ファイルとして直接ダウンロードすることはできません。 回避策は、BLOB を PDF ファイルとしてダウンロードする簡単な Web アプリケーションを作成することです。 アドインで、メソッドを呼び出 `Office.context.ui.openBrowserWindow(url)` し、Web アプリケーションの URL を渡します。 これにより、Web アプリケーションがブラウザー ウィンドウの外部で開Office。

## <a name="see-also"></a>関連項目

- [Officeアドインを実行するための要件](requirements-for-running-office-add-ins.md)
