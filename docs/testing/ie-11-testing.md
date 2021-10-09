---
title: Internet Explorer 11 テスト
description: 11 でOfficeアドインをテストInternet Explorerします。
ms.date: 10/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: cfa6a35565fdca28eab9734ccde9fc8fbb2e8270
ms.sourcegitcommit: a37be80cf47a37c85b7f5cab216c160f4e905474
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/09/2021
ms.locfileid: "60250519"
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
> また、Internet Explorer 11 は、メディア、録音、および位置情報などの HTML 5 機能の一部をサポートしていません。

> [!NOTE]
> Office on the web 11 で開くInternet Explorerできないので、Office on the web でアドインをテストInternet Explorer。

## <a name="prerequisites"></a>前提条件

- [Node.js](https://nodejs.org/) (最新 [LTS](https://nodejs.org/about/releases) バージョン)

これらの手順では、以前に Yo ジェネレーター プロジェクトをOffice前提とします。 前にこれを行ったことがない場合は、クイック スタート (アドイン用など) を読[Excel検討してください](../quickstarts/excel-quickstart-jquery.md)。

## <a name="switching-to-the-internet-explorer-11-webview"></a>11 webview Internet Explorer切り替える

1. Yo ジェネレーター プロジェクトOffice作成します。 選択するプロジェクトの種類は関係ありませんが、このツールは、すべてのプロジェクトの種類で動作します。

    > [!NOTE]
    > 既存のプロジェクトを持ち、新しいプロジェクトを作成せずにこのツールを追加する場合は、この手順をスキップして次の手順に進みます。 

1. プロジェクトのルート フォルダーで、コマンド ラインで次のコマンドを実行します。 この例では、プロジェクトのマニフェスト ファイルがルートにあると仮定します。 指定されていない場合は、マニフェスト ファイルへの相対パスを指定します。 コマンド ラインに、Web ビューの種類が IE に設定されているというメッセージが表示されます。

    ```command&nbsp;line
    npx office-addin-dev-settings webview manifest.xml ie
    ```

> [!TIP]
> このコマンドを使用する必要はありません。ただし、11 ランタイムに関連する問題の大部分をデバッグInternet Explorer必要があります。 完全な堅牢性を得る場合は、Windows 7、8.1、10、および 11 のさまざまなバージョンの Office のさまざまな組み合わせのコンピューターを使用してテストする必要があります。 詳細については、「Office アドインで使用されるブラウザー」および「How to revert [to](../concepts/browsers-used-by-office-web-add-ins.md) earlier version of Office」 を[参照してください](https://support.microsoft.com/topic/2bd5c457-a917-d57e-35a1-f709e3dda841)。

### <a name="command-options"></a>コマンド オプション

この `office-addin-dev-settings webview` コマンドは、引数として多数のランタイムを受け取る場合があります。

- すなわち
- エッジ
- default

## <a name="see-also"></a>関連項目

* [Office アドインのテストとデバッグ](test-debug-office-add-ins.md)
* [テスト用に Office アドインをサイドロードする](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
* [アプリの開発者ツールを使用してアドインをデバッグWindows](debug-add-ins-using-f12-developer-tools-on-windows.md)
* [作業ウィンドウからデバッガーをアタッチする](attach-debugger-from-task-pane.md)
