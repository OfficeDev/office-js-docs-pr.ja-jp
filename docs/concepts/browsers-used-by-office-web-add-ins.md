---
title: Office アドインによって使用されるブラウザー
description: Office アドインによって使用されるブラウザーをオペレーティング システムおよび Office バージョンが決定する方法を指定します。
ms.date: 09/25/2019
localization_priority: Priority
ms.openlocfilehash: b5d7198e556f020bccdf7ba1e0a0fcffa3a9171b
ms.sourcegitcommit: c8914ce0f48a0c19bbfc3276a80d090bb7ce68e1
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/26/2019
ms.locfileid: "37235296"
---
# <a name="browsers-used-by-office-add-ins"></a>Office アドインによって使用されるブラウザー

Office アドインは、Office on the web での実行時に iFrame を使用して表示され、デスクトップおよびモバイル クライアント用に Office に埋め込まれたブラウザー コントロールを使用して表示される Web アプリケーションです。 アドインには JavaScript を実行するための JavaScript エンジンも必要です。 埋め込まれたブラウザーおよびエンジン、どちらもユーザーのコンピュータにインストールされているブラウザによって提供されます。

どのブラウザが使用されているかは、以下によります。

- コンピュータのオペレーティングシステム。
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
|Windows 10 バージョン >= 1903 / Office 365 ver < 16.0.11629|Internet Explorer 11|
|Windows 10 バージョン >= 1903 / Office 365 ver >= 16.0.11629|Microsoft Edge\*|

\*Microsoft Edge が使用されている場合、Windows 10 ナレーター (「スクリーン リーダー」と呼ばれることもあります) は、作業ウィンドウで開いているページの `<title>` タグを読み取ります。 Internet Explorer 11 が使用されている場合、ナレーターはアドイン マニフェストの `<DisplayName>` の値から提供される作業ウィンドウのタイトル バーを読み取ります。

> [!IMPORTANT]
> Internet Explorer 11はES5以降のJavaScriptバージョンをサポートしていません。 アドインのユーザーが Internet Explorer 11 を使用するプラットフォームを使用している場合、ECMAScript 2015 以降の構文と機能を使用するには、JavaScript を ES 5 にトランスパイルするか、ポリフィルを使用する必要があります。 また、Internet Explorer 11 は、メディア、録音、および位置情報などの HTML 5 機能の一部をサポートしていません。

> [!NOTE]
> これらが一般に利用可能になるまで、Windows バージョン 1903 以降を入手するには Windows Insider である必要があり、また、Office バージョン 16.0.11629 以降を入手するには Office Insider である必要があります。
>
> Windows インサイダーに参加するには
> 
> 1. [Windows インサイダー](https://insider.windows.com)に移動し、リンクをクリックしてWindows インサイダーに参加してください。
> 2. Windowsのプレビュービルドを有効にするためのWindowsの設定の使用方法についての説明が記載されたページに移動します。 指示に従います。 更新頻度を選択する際は、一番速いオプションを選択してください。
>
> Office インサイダーに参加するには
> 
> 1. [Office Insiderになりましょう](https://insider.office.com/join)に移動してください。
> 2. そのページの指示に従って参加してください。 チャンネルを指定するように求められたら、[インサイダー]を選択します。

## <a name="troubleshooting-microsoft-edge-issues"></a>Microsoft Edge の問題のトラブルシューティング

### <a name="scroll-bar-does-not-appear-in-task-pane"></a>作業ウィンドウにスクロール バーが表示されない

既定では、Microsoft Edge のスクロール バーはホバーするまで非表示になっています。 スクロールバーが常に表示されるようにするには、作業ウィンドウのページの`<body>`要素に適用される CSS スタイルに [-ms-overflow-style](https://developer.mozilla.org/docs/Web/CSS/-ms-overflow-style) プロパティを含め、`scrollbar`に設定する必要があります。 

### <a name="when-debugging-with-the-microsoft-edge-devtools-the-add-in-crashes-or-reloads"></a>Microsoft Edge DevTools を使用してデバッグすると、アドインがクラッシュまたは再読み込みされる

[Microsoft Edge DevTools](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?rtc=1&activetab=pivot%3Aoverviewtab) にブレークポイントを設定すると、アドインがハングしていると Office に判断される可能性があります。 これが発生すると、アドインが自動的に再読み込みされます。 これを防ぐには、開発用コンピューターに以下のレジストリ キーと値を追加します: `[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Wef]"AlertInterval"=dword:00000000`。

### <a name="when-the-add-in-tries-to-open-get-add-in-error-we-cant-open-this-add-in-from-the-localhost-error"></a>アドインを開こうとすると、"アドイン エラー localhost からこのアドインを開けません" というエラーが表示される

既知の原因の1つとして、Microsoft Edge では開発用コンピューター上では localhost にループバックの除外を与える必要があることが挙げられます。 [Cannot open add-in from localhost (localhostからアドインを開くことができません)](/office/troubleshoot/error-messages/cannot-open-add-in-from-localhost)の指示に従ってください。


## <a name="see-also"></a>関連項目

- [Officeアドインを実行するための要件](requirements-for-running-office-add-ins.md)
