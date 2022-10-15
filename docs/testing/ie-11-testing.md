---
title: Internet Explorer 11 のテスト
description: Internet Explorer 11 で Office アドインをテストします。
ms.date: 10/12/2022
ms.localizationpriority: medium
ms.openlocfilehash: 887d71e8e054a39c474ef300aeeb588933499ed8
ms.sourcegitcommit: 5544cf174d145e356e33866e2480bde999514ada
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/14/2022
ms.locfileid: "68574164"
---
# <a name="test-your-office-add-in-on-internet-explorer-11"></a>Internet Explorer 11 で Office アドインをテストする

> [!IMPORTANT]
> **Office アドインで引き続き使用される Internet Explorer**
>
> Office 2019 を通じて 1 回限りの購入バージョンを含む一部のプラットフォームと Office バージョンの組み合わせでは、Office [アドインで使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md)で説明されているように、Internet Explorer 11 に付属する Web ビュー コントロールを引き続き使用してアドインをホストします。Internet Explorer Webview でアドインを起動したときにアドインのユーザーに正常なエラー メッセージを提供することで、少なくとも最小限の方法でこれらの組み合わせを引き続きサポートすることをお勧めします (ただし、必要ありません)。 次の点に注意してください。
>
> - Internet Explorer でOffice on the webが開かなくなりました。 その結果、[AppSource は](/office/dev/store/submit-to-appsource-via-partner-center)、ブラウザーとして Internet Explorer を使用してOffice on the webでアドインをテストしなくなりました。
> - AppSource は引き続き Internet Explorer を使用するプラットフォームと Office *デスクトップ* バージョンの組み合わせをテストしますが、アドインが Internet Explorer をサポートしていない場合にのみ警告が発行されます。アドインは AppSource によって拒否されません。
> - [Script Lab ツール](../overview/explore-with-script-lab.md)は Internet Explorer をサポートしなくなりました。

以前のバージョンの Windows と Office をサポートする予定の場合、アドインは Internet Explorer 11 (IE11) に基づく埋め込み可能なブラウザー コントロールで動作する必要があります。 コマンド ラインを使用すると、アドインで使用されるよりモダンなランタイムから、このテスト用の Internet Explorer 11 ランタイムに切り替えることができます。 Internet Explorer 11 Web ビュー コントロールを使用する Windows と Office のバージョンについては、「 [Office アドインで使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md)」を参照してください。

> [!IMPORTANT]
> Internet Explorer 11はES5以降のJavaScriptバージョンをサポートしていません。 ECMAScript 2015 以降の構文と機能を使用する場合は、次の 2 つのオプションがあります。
>
> - ECMAScript 2015 (ES6 とも呼ばれます) または TypeScript でコードを記述し、 [バベル](https://babeljs.io/) や [tsc](https://www.typescriptlang.org/index.html) などのコンパイラを使用してコードを ES5 JavaScript にコンパイルします。
> - ECMAScript 2015 以降の JavaScript で記述しますが、IE でコードを実行できるようにする [core-js](https://github.com/zloirock/core-js) などの[ポリフィル](https://en.wikipedia.org/wiki/Polyfill_(programming)) ライブラリも読み込みます。
>
> これらのオプションの詳細については、 [Internet Explorer 11 のサポートに関するページを](../develop/support-ie-11.md)参照してください。
>
> また、Internet Explorer 11 は、メディア、録音、および位置情報などの HTML 5 機能の一部をサポートしていません。 詳細については、「 [Internet Explorer でアドインが実行されているかどうかを実行時に確認](../develop/support-ie-11.md#determine-at-runtime-if-the-add-in-is-running-in-internet-explorer)する」を参照してください。

> [!NOTE]
> - Office on the web Internet Explorer 11 で開くことができないので、Internet Explorer を使用してOffice on the webでアドインをテストすることはできません (およびテストする必要はありません)。
>
> - Office Web アドインが機能するためには、Internet Explorer のセキュリティ強化の構成 (ESC) がオフになっている必要があります。 アドインを開発する際に Windows Server コンピューターをクライアントとして使用する場合は、Windows Server では既定で ESC がオンになっていることに注意してください。

## <a name="switch-to-the-internet-explorer-11-webview"></a>Internet Explorer 11 Webview に切り替える

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

Internet Explorer Webview を切り替えるには、2 つの方法があります。 コマンド プロンプトで簡単なコマンドを実行することも、Internet Explorer を既定で使用するバージョンの Office をインストールすることもできます。 最初の方法をお勧めします。 ただし、次のシナリオでは 2 つ目を使用する必要があります。

- プロジェクトは Visual Studio と IIS で開発されました。 node.jsベースではありません。
- テストで絶対に堅牢になりたいと考えています。
- 開発用コンピューターで Microsoft 365 のベータ チャネルを使用することはできません。
- Mac で開発中です。 
- 何らかの理由でコマンド ライン ツールが機能しない場合。

### <a name="switch-via-the-command-line"></a>コマンド ラインを使用して切り替える

[!INCLUDE [Steps to switch browsers with the command line tool](../includes/use-legacy-edge-or-ie.md)]

### <a name="install-a-version-of-office-that-uses-internet-explorer"></a>Internet Explorer を使用する Office のバージョンをインストールする

[!INCLUDE [Steps to install Office that uses Edge Legacy or Internet Explorer](../includes/install-office-that-uses-legacy-edge-or-ie.md)]

## <a name="see-also"></a>関連項目

- [Office アドインのテストとデバッグ](test-debug-office-add-ins.md)
- [テスト用に Office アドインをサイドロードする](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
- [Internet Explorer の開発者ツールを使用してアドインをデバッグする](debug-add-ins-using-f12-tools-ie.md)
- [作業ウィンドウからデバッガーをアタッチする](attach-debugger-from-task-pane.md)
- [Office アドインのランタイム](runtimes.md)