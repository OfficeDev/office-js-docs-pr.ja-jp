---
title: Office アドインによって使用されるブラウザー
description: Office アドインによって使用されるブラウザーをオペレーティング システムおよび Office バージョンが決定する方法を指定します。
ms.date: 04/21/2020
localization_priority: Normal
ms.openlocfilehash: 9ef4b6d4c09140fc6d6bb04eca51d845b79b6dc7
ms.sourcegitcommit: 3355c6bd64ecb45cea4c0d319053397f11bc9834
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/22/2020
ms.locfileid: "43744853"
---
# <a name="browsers-used-by-office-add-ins"></a>Office アドインによって使用されるブラウザー

Office アドインは、Office on the web での実行時に iFrame を使用して表示され、デスクトップおよびモバイル クライアント用に Office に埋め込まれたブラウザー コントロールを使用して表示される Web アプリケーションです。 アドインには JavaScript を実行するための JavaScript エンジンも必要です。 埋め込みブラウザーとエンジンの両方が、ユーザーのコンピューターにインストールされているブラウザーから提供されます。

どのブラウザが使用されているかは、以下によります。

- コンピューターのオペレーティングシステム。
- アドインが Office on the web、Office 365、または登録のない Office 2013 以降で実行されているかどうか。

次の表は、さまざまなプラットフォームとオペレーティングシステムに使用されているブラウザを示しています。

|**OS / Platform**|**Browser**|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|Office on the web|Office が開かれているブラウザー。|
|Mac|Safari|
|iOS|Safari|
|Android|Chrome|
|Windows / 非登録 Office 2013以降|Internet Explorer 11|
|Windows 10 バージョン < 1903 / Office 365|Internet Explorer 11|
|Windows 10 バージョン >= 1903/Office 365 ver < 16.0.11629<sup>1</sup>|Internet Explorer 11|
|Windows 10 バージョン >= 1903/Office 365 ver >= 16.0.11629<sup>1</sup>|Microsoft Edge<sup>2</sup>|

<sup>1</sup> [[更新履歴] ページ](/officeupdates/update-history-office365-proplus-by-date)を参照してください。詳細については、「 [Office クライアントのバージョンと更新プログラムのチャネルを見つける](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)方法」を参照してください。

<sup>2</sup> Microsoft Edge を使用している場合、Windows 10 ナレーター ("スクリーンリーダー" と呼ばれることも`<title>`あります) は、作業ウィンドウに表示されるページのタグを読み取ります。 Internet Explorer 11 が使用されている場合、ナレーターはアドイン マニフェストの `<DisplayName>` の値から提供される作業ウィンドウのタイトル バーを読み取ります。

> [!IMPORTANT]
> Internet Explorer 11はES5以降のJavaScriptバージョンをサポートしていません。 アドインのユーザーが Internet Explorer 11 を使用するプラットフォームを使用している場合、ECMAScript 2015 以降の構文と機能を使用するには、JavaScript を ES 5 にトランスパイルするか、ポリフィルを使用する必要があります。 また、Internet Explorer 11 は、メディア、録音、および位置情報などの HTML 5 機能の一部をサポートしていません。

## <a name="troubleshooting-microsoft-edge-issues"></a>Microsoft Edge の問題のトラブルシューティング

### <a name="service-workers-are-not-working"></a>サービスワーカーが動作しない

Office アドインでは、 [Microsoft Edge WebView](/microsoft-edge/hosting/webview)でサービスワーカーをサポートしていません。 エッジ WebView コントロールでサポートされている最新の機能については、「 [Office アドインの概要](../overview/office-add-ins.md)」を参照してください。 Microsoft は、サービスワーカーをサポートすると予想される Office アドインプラットフォームに新しい[Chromium ベースのエッジ WebView2](/microsoft-edge/hosting/webview2)をもたらすのは困難です。

### <a name="chromium-based-edge-is-installed-on-my-development-computer-but-my-add-in-does-not-use-it"></a>Chromium ベースのエッジは開発用のコンピューターにインストールされていますが、アドインでは使用されません

[Microsoft Edge](https://support.microsoft.com/help/4501095/download-the-new-microsoft-edge-based-on-chromium)のベースブラウザーが Chromium に変更されました。 Chromium ベースのエッジがインストールされている場合、EdgeHTML と呼ばれる以前のベースは削除されません。 Office は、Chromium をサポートする Office 365 のビルドがコンピューターにインストールされるまで、アドインの EdgeHTML ベースを引き続き使用します。 これらのビルドは2020で出荷されることが予想されます。 そのようなメッセージは、年の前半に、内部の Insider チャネルに表示される可能性があります。

### <a name="scroll-bar-does-not-appear-in-task-pane"></a>作業ウィンドウにスクロール バーが表示されない

既定では、Microsoft Edge のスクロール バーはホバーするまで非表示になっています。 スクロールバーが常に表示されるようにするには、作業ウィンドウのページの`<body>`要素に適用される CSS スタイルに [-ms-overflow-style](https://developer.mozilla.org/docs/Web/CSS/-ms-overflow-style) プロパティを含め、`scrollbar`に設定する必要があります。 

### <a name="when-debugging-with-the-microsoft-edge-devtools-the-add-in-crashes-or-reloads"></a>Microsoft Edge DevTools を使用してデバッグすると、アドインがクラッシュまたは再読み込みされる

[Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab) にブレークポイントを設定すると、アドインがハングしていると Office に判断される可能性があります。 これが発生すると、アドインが自動的に再読み込みされます。 これを防ぐには、開発用コンピューターに以下のレジストリ キーと値を追加します: `[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`。

### <a name="when-the-add-in-tries-to-open-get-add-in-error-we-cant-open-this-add-in-from-the-localhost-error"></a>アドインを開こうとすると、"アドイン エラー localhost からこのアドインを開けません" というエラーが表示される

既知の原因の1つとして、Microsoft Edge では開発用コンピューター上では localhost にループバックの除外を与える必要があることが挙げられます。 [Cannot open add-in from localhost (localhostからアドインを開くことができません)](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost)の指示に従ってください。

### <a name="get-errors-trying-to-download-a-pdf-file"></a>PDF ファイルをダウンロードしようとしてエラーを取得する

アドインで blob を PDF ファイルとして直接ダウンロードすることは、エッジがブラウザーの場合はサポートされていません。 回避策は、blob を PDF ファイルとしてダウンロードする簡単な web アプリケーションを作成することです。 アドインで`Office.context.ui.openBrowserWindow(url)`メソッドを呼び出し、web アプリケーションの URL を渡します。 これにより、Office の外部にあるブラウザーウィンドウで web アプリケーションが開きます。

## <a name="see-also"></a>関連項目

- [Officeアドインを実行するための要件](requirements-for-running-office-add-ins.md)
