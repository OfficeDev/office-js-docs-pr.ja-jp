---
title: Officeアドインで使用されるWebビューア
description: ''
ms.date: 05/03/2019
localization_priority: Priority
ms.openlocfilehash: 632f62cbc02917d9e28ab260f3710498156194db
ms.sourcegitcommit: 47b792755e655043d3db2f1fdb9a1eeb7453c636
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/06/2019
ms.locfileid: "33630406"
---
# <a name="web-viewers-used-by-office-add-ins"></a>Officeアドインで使用されるWebビューア

OfficeアドインはWebアプリケーションなので、WebアプリケーションのHTMLページを表示するためのWebページビューアと、JavaScriptを実行するためのJavaScriptエンジンが必要です。 どちらもユーザーのコンピュータにインストールされているブラウザによって提供されます。

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
|Windows 10 バージョン >= 1903 / Office 365 ver >= 16.0.11629|Edge\*|

\*Edgeが使用されている場合、Windows 10ナレータ（ "スクリーンリーダー"と呼ばれることもあります）は、作業ペインに表示されるページの`<title>`タグを読み取ります。 Internet Explorer 11が使用されている場合、ナレータは作業ペインのタイトルバーを読み取ります。これはアドインのマニフェストの`<DisplayName>`の値から取得されます。

> [!IMPORTANT]
> Internet Explorer 11はES5以降のJavaScriptバージョンをサポートしていません。 アドインのユーザーがInternet Explorer 11を使用するプラットフォームを使用している場合に、ECMAScript 2015以降の構文と機能を使用するには、JavaScriptをES 5に変換するか、ポリフィルを使用する必要があります。 また、Internet Explorer 11は、メディア、録音、および位置情報などのHTML 5機能の一部をサポートしていません。

> [!NOTE]
> これらが一般に利用可能になるまで、Windowsバージョン1903以降を入手するためWindowsインサイダーである必要があり、また、Officeバージョン16.0.11629以降を入手するためOfficeインサイダーである必要があります。
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
