---
title: Internet Explorer 11 テスト
description: 11 でOfficeアドインをテストInternet Explorerします。
ms.date: 11/02/2021
ms.localizationpriority: medium
ms.openlocfilehash: 8932545aa692073babeddb6ab22a213466a7c2ba
ms.sourcegitcommit: a3debae780126e03a1b566efdec4d8be83e405b8
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/03/2021
ms.locfileid: "60809040"
---
# <a name="test-your-office-add-in-on-internet-explorer-11"></a>11 でOfficeアドインをテストInternet Explorerする

> [!IMPORTANT]
> **Internet ExplorerアドインOffice引き続き使用する**
>
> Microsoft は、アドインのサポートInternet Explorer終了していますが、これはアドインのOffice大きな影響を及ぼします。Office アドインで使用されるブラウザーで説明したように、プラットフォームと Office バージョンの一部の組み合わせ (Office 2019 までの 1 回限り購入バージョンを含む) は、Internet Explorer 11 に付属する[](../concepts/browsers-used-by-office-web-add-ins.md)webview コントロールを引き続き使用してアドインをホストします。さらに、これらの組み合わせのサポートは、AppSource にInternet Explorerアドインに対して引き続き[必要です](/office/dev/store/submit-to-appsource-via-partner-center)。 次の *2 つの点が変化* しています。
>
> - Office on the webで開かなくなったInternet Explorer。 そのため、AppSource はブラウザーとしてアプリケーション を使用してOffice on the webアドインInternet Explorerテストしなくなりました。 ただし、AppSource は引き続き、プラットフォームとデスクトップ バージョンの組み合Office *使用* するデスクトップ バージョンの組み合わせをテストInternet Explorer。
> - この[Script Labは](../overview/explore-with-script-lab.md)サポートされなくなりましたInternet Explorer。

AppSource を使用してアドインを販売する予定がある場合、または以前のバージョンの Windows および Office をサポートする予定の場合、アドインは Internet Explorer 11 (IE11) に基づく埋め込み可能なブラウザー コントロールで動作する必要があります。 コマンド ラインを使用して、アドインで使用される最新のランタイムから、このテスト用の Internet Explorer 11 ランタイムに切り替えます。 Windows および Office のバージョンで Internet Explorer 11 Web ビュー コントロールを使用する方法については、「Office アドインで使用されるブラウザー」を[参照](../concepts/browsers-used-by-office-web-add-ins.md)してください。

> [!IMPORTANT]
> Internet Explorer 11はES5以降のJavaScriptバージョンをサポートしていません。 ECMAScript 2015 以降の構文と機能を使用する場合は、次の 2 つのオプションがあります。
>
> - ECMAScript 2015 (ES6 とも呼ばれる) 以降の JavaScript または TypeScript でコードを記述し、バベルや[tsc](https://www.typescriptlang.org/index.html)などの[](https://babeljs.io/)コンパイラを使用してコードを ES5 JavaScript にコンパイルします。
> - ECMAScript 2015 以降の JavaScript で記述します[](https://en.wikipedia.org/wiki/Polyfill_(programming))が、IE でコードを実行できる[core-js](https://github.com/zloirock/core-js)などのポリフィル ライブラリも読み込む必要があります。
>
> これらのオプションの詳細については [、「Support Internet Explorer 11」を参照してください](../develop/support-ie-11.md)。
>
> また、Internet Explorer 11 は、メディア、録音、および位置情報などの HTML 5 機能の一部をサポートしていません。 詳細については、「アドインが実行中かどうかを実行時に確認する」を参照[Internet Explorer。](../develop/support-ie-11.md#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer)

> [!NOTE]
> Office on the web 11 で開くInternet Explorerできないので、Office on the web でアドインをテストInternet Explorer。

## <a name="switch-to-the-internet-explorer-11-webview"></a>11 webview Internet Explorerに切り替える

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

Webview に切り替える方法は 2 Internet Explorerがあります。 コマンド プロンプトで簡単なコマンドを実行するか、既定でコマンド を使用OfficeバージョンInternet Explorerインストールできます。 最初の方法をお勧めします。 ただし、次のシナリオでは 2 つ目を使用する必要があります。

- プロジェクトは、プロジェクトと IIS Visual Studio開発されました。 この機能は、node.jsに基づいて行う必要があります。
- テストで絶対に堅牢になる必要があります。
- 何らかの理由でコマンド ライン ツールが機能しない場合。

### <a name="switch-via-the-command-line"></a>コマンド ライン経由で切り替える

[!INCLUDE [Steps to switch browsers with the command line tool](../includes/use-legacy-edge-or-ie.md)]

### <a name="install-a-version-of-office-that-uses-internet-explorer"></a>アプリケーションを使用するOfficeバージョンをインストールInternet Explorer

[!INCLUDE [Steps to install Office that uses Edge Legacy or Internet Explorer](../includes/install-office-that-uses-legacy-edge-or-ie.md)]

## <a name="see-also"></a>関連項目

* [Office アドインのテストとデバッグ](test-debug-office-add-ins.md)
* [テスト用に Office アドインをサイドロードする](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
* [Internet Explorer の開発者ツールを使用してアドインをデバッグする](debug-add-ins-using-f12-tools-ie.md)
* [作業ウィンドウからデバッガーをアタッチする](attach-debugger-from-task-pane.md)
