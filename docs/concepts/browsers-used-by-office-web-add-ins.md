---
title: Office アドインによって使用されるブラウザー
description: Office アドインによって使用されるブラウザーをオペレーティング システムおよび Office バージョンが決定する方法を指定します。
ms.date: 05/28/2019
localization_priority: Priority
ms.openlocfilehash: 92218bb012ae9031ebfc429606885a0ec0ea85b3
ms.sourcegitcommit: b299b8a5dfffb6102cb14b431bdde4861abfb47f
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/30/2019
ms.locfileid: "34592130"
---
# <a name="browsers-used-by-office-add-ins"></a>Office アドインによって使用されるブラウザー

Office アドインは、Office Online で実行しているときに iFrame を使用して表示され、デスクトップおよびモバイル クライアント用に Office に埋め込まれたブラウザー コントロールを使用して表示される Web アプリケーションです。 アドインには JavaScript を実行するための JavaScript エンジンも必要です。 埋め込まれたブラウザーおよびエンジン、どちらもユーザーのコンピュータにインストールされているブラウザによって提供されます。

どのブラウザが使用されているかは、以下によります。

- コンピュータのオペレーティングシステム。
- アドインがOffice Online、Office 365、または登録のないOffice 2013以降で実行されているかどうか。

次の表は、さまざまなプラットフォームとオペレーティングシステムに使用されているブラウザを示しています。

|**OS / Platform**|**Browser**|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|Office Online|Office Onlineが開かれているブラウザ。|
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

## <a name="see-also"></a>関連項目

- [Officeアドインを実行するための要件](requirements-for-running-office-add-ins.md)
