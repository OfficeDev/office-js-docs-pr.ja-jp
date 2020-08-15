---
title: Office アドインによって使用されるブラウザー
description: Office アドインによって使用されるブラウザーをオペレーティング システムおよび Office バージョンが決定する方法を指定します。
ms.date: 08/13/2020
localization_priority: Normal
ms.openlocfilehash: 53e3061f1729ac792e91a10e31bc9d0d908ab07b
ms.sourcegitcommit: 3efa932b70035dde922929d207896e1a6007f620
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/15/2020
ms.locfileid: "46757360"
---
# <a name="browsers-used-by-office-add-ins"></a>Office アドインによって使用されるブラウザー

Office アドインは、Office on the web での実行時に iFrame を使用して表示され、デスクトップおよびモバイル クライアント用に Office に埋め込まれたブラウザー コントロールを使用して表示される Web アプリケーションです。 アドインには JavaScript を実行するための JavaScript エンジンも必要です。 埋め込みブラウザーとエンジンの両方が、ユーザーのコンピューターにインストールされているブラウザーから提供されます。

どのブラウザが使用されているかは、以下によります。

- コンピューターのオペレーティングシステム。
- アドインが Office on the web、Microsoft 365、またはサブスクリプション外の Office 2013 以降で実行されているかどうか。

次の表は、さまざまなプラットフォームとオペレーティングシステムに使用されているブラウザを示しています。

|HP-UX|Office のバージョン
|エッジ WebView2 (Chromium ベース) がインストールされているかどうか|ブラウザー|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|any|Office on the web|該当なし|Office が開かれているブラウザー。|
|Mac|any|該当なし|Safari|
|iOS|any|該当なし|Safari|
|Android|any|該当なし|Chrome|
|Windows 7、8.1、10 | サブスクリプション以外の Office 2013 以降|かまいません|Internet Explorer 11|
|Windows 7 | Microsoft 365| かまいません | Internet Explorer 11|
|Windows 8.1、<br>Windows 10 のバージョン。 &nbsp; < &nbsp;1903| Microsoft 365 | いいえ| Internet Explorer 11|
|Windows 10 のバージョン。 &nbsp; >= &nbsp;1903 | Microsoft 365 ver。 &nbsp; < &nbsp;16.0.11629<sup>1</sup>| かまいません|Internet Explorer 11|
|Windows 10 のバージョン。 &nbsp; >= &nbsp;1903 | Microsoft 365 ver。 &nbsp; >= &nbsp;16.0.11629 &nbsp; _および_ &nbsp; < &nbsp; 16.0.13127.20082<sup>1</sup>| かまいません|Microsoft Edge<sup>2、3</sup> (元の WebView を使用) (EdgeHTML)|
|Windows 10 のバージョン。 &nbsp; >= &nbsp;1903 | Microsoft 365 ver。 &nbsp; >= &nbsp;16.0.13127.20082<sup>1</sup>| いいえ |Microsoft Edge<sup>2、3</sup> (元の WebView を使用) (EdgeHTML)|
|Windows 8.1<br>Windows 10| Microsoft 365 ver。 &nbsp; >= &nbsp;16.0.13127.20082<sup>1</sup>| はい|  下のメモ4を参照してください。 |

<sup>1</sup> [ [更新履歴] ページ](/officeupdates/update-history-office365-proplus-by-date) を参照してください。詳細については、「 [Office クライアントのバージョンと更新プログラムのチャネルを見つける](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19) 方法」を参照してください。

<sup>2</sup> Microsoft Edge を使用している場合、Windows 10 ナレーター ("スクリーンリーダー" と呼ばれることもあります) は、 `<title>` 作業ウィンドウに表示されるページのタグを読み取ります。 Internet Explorer 11 が使用されている場合、ナレーターはアドイン マニフェストの `<DisplayName>` の値から提供される作業ウィンドウのタイトル バーを読み取ります。

<sup>3</sup> アドインにマニフェスト内の要素が含まれている場合は `Runtimes` 、Windows または Microsoft 365 のバージョンに関係なく Internet Explorer 11 を使用します。 詳細については、「 [ランタイム](../reference/manifest/runtimes.md)」を参照してください。

<sup>4</sup> このバージョンの組み合わせに使用されるブラウザーは、Microsoft 365 サブスクリプションの更新プログラムチャネルによって異なります。 ユーザーが [ベータチャネル](https://insider.office.com/join/windows) (旧称 insider Fast channel) 上にある場合、Office は WebView2 (Chromium ベース) で Microsoft Edge を使用します。 その他のチャネルについては、Office は Microsoft Edge と元の WebView (EdgeHTML) を使用します。 2021の初期段階では、他のチャネルでの WebView2 のサポートが期待されています。
> [!IMPORTANT]
> Internet Explorer 11はES5以降のJavaScriptバージョンをサポートしていません。 アドインのユーザーが Internet Explorer 11 を使用するプラットフォームを使用している場合、ECMAScript 2015 以降の構文と機能を使用するには、JavaScript を ES 5 にトランスパイルするか、ポリフィルを使用する必要があります。 また、Internet Explorer 11 は、メディア、録音、および位置情報などの HTML 5 機能の一部をサポートしていません。

## <a name="troubleshooting-microsoft-edge-issues"></a>Microsoft Edge の問題のトラブルシューティング

### <a name="service-workers-are-not-working"></a>サービスワーカーが動作しない

Office アドインでは、元の [Microsoft Edge WebView](/microsoft-edge/hosting/webview) が使用されている場合、サービスワーカーはサポートされません。 [Chromium ベースのエッジ WebView2](/microsoft-edge/hosting/webview2)でサポートされています。

### <a name="scroll-bar-does-not-appear-in-task-pane"></a>作業ウィンドウにスクロール バーが表示されない

既定では、Microsoft Edge のスクロール バーはホバーするまで非表示になっています。 スクロールバーが常に表示されるようにするには、作業ウィンドウのページの`<body>`要素に適用される CSS スタイルに [-ms-overflow-style](https://developer.mozilla.org/docs/Web/CSS/-ms-overflow-style) プロパティを含め、`scrollbar`に設定する必要があります。 

### <a name="when-debugging-with-the-microsoft-edge-devtools-the-add-in-crashes-or-reloads"></a>Microsoft Edge DevTools を使用してデバッグすると、アドインがクラッシュまたは再読み込みされる

[Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab) にブレークポイントを設定すると、アドインがハングしていると Office に判断される可能性があります。 これが発生すると、アドインが自動的に再読み込みされます。 これを防ぐには、開発用コンピューターに以下のレジストリ キーと値を追加します: `[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`。

### <a name="when-the-add-in-tries-to-open-get-add-in-error-we-cant-open-this-add-in-from-the-localhost-error"></a>アドインを開こうとすると、"アドイン エラー localhost からこのアドインを開けません" というエラーが表示される

既知の原因の1つとして、Microsoft Edge では開発用コンピューター上では localhost にループバックの除外を与える必要があることが挙げられます。 [Cannot open add-in from localhost (localhostからアドインを開くことができません)](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost)の指示に従ってください。

### <a name="get-errors-trying-to-download-a-pdf-file"></a>PDF ファイルをダウンロードしようとしてエラーを取得する

アドインで blob を PDF ファイルとして直接ダウンロードすることは、エッジがブラウザーの場合はサポートされていません。 回避策は、blob を PDF ファイルとしてダウンロードする簡単な web アプリケーションを作成することです。 アドインで、メソッドを呼び出し、 `Office.context.ui.openBrowserWindow(url)` web アプリケーションの URL を渡します。 これにより、Office の外部にあるブラウザーウィンドウで web アプリケーションが開きます。

## <a name="see-also"></a>関連項目

- [Officeアドインを実行するための要件](requirements-for-running-office-add-ins.md)
