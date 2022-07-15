---
title: Outlook アドインで関数コマンドをデバッグする
description: Outlook アドインで関数コマンドをデバッグする方法について説明します。
ms.topic: article
ms.date: 07/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 6189824fd526d48321b355c9b306fa5ef732f411
ms.sourcegitcommit: 9bb790f6264f7206396b32a677a9133ab4854d4e
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/15/2022
ms.locfileid: "66797590"
---
# <a name="debug-function-commands-in-outlook-add-ins"></a>Outlook アドインで関数コマンドをデバッグする

> [!NOTE]
> この記事の手法は、Windows 開発コンピューターでのみ使用できます。 Mac で開発している場合は、「 [関数のデバッグ コマンド](../testing/debug-function-command.md)」を参照してください。

この記事では、Visual Studio Code で Office アドイン デバッガー拡張機能を使用して [関数コマンド](add-in-commands-for-outlook.md#run-a-function-command)をデバッグする方法について説明します。 関数コマンドは、リボンのアドイン コマンド ボタンを使用して開始されます。 アドイン コマンドの詳細については、「 [Outlook のアドイン コマンド」を](add-in-commands-for-outlook.md)参照してください。

この記事では、デバッグするアドイン プロジェクトが既にあることを前提としています。 デバッグを実行する関数コマンドを使用してアドインを作成するには、「 [チュートリアル: メッセージ作成 Outlook アドインをビルド](../tutorials/outlook-tutorial.md)する」の手順に従います。

## <a name="mark-your-add-in-for-debugging"></a>デバッグ用にアドインをマークする

[Office アドイン用 Yeoman ジェネレーター](../develop/yeoman-generator-overview.md)を使用してアドイン プロジェクトを作成した場合は、この記事の後半の[「デバッガーの構成と実行](#configure-and-run-the-debugger)」セクションに進んでください。 アドインをビルドしてローカル サーバーを起動するために実行`npm start`すると、このコマンドは、デバッグ用にアドインを`UseDirectDebugger``HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]`マークするレジストリ キーの値も設定します。

それ以外の場合は、別のツールを使用してアドインを作成した場合は、次の手順を実行します。

1. レジストリ キーに `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]` 移動します。 アドインの **\<Id\>** マニフェストに置き換えます`[Add-in ID]`。

    [!include[Developer registry key](../includes/developer-registry-key.md)]

1. キーの値`1`を `UseDirectDebugger` .

## <a name="configure-and-run-the-debugger"></a>デバッガーを構成して実行する

アドインでデバッグを有効にしたので、デバッガーを構成して実行する準備ができました。 これを行う方法については、Webview コントロールに適用される次のいずれかのオプションを選択します。 開発用コンピューターで使用されている Webview コントロールを特定する方法については、「 [Office アドインで使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md)」を参照してください。

- アドインが Edge Legacy (EdgeHTML) の埋め込み Web ビュー コントロールで実行される場合は、「 [Visual Studio Code 用 Microsoft Office アドイン デバッガー拡張機能」](../testing/debug-with-vs-extension.md)を参照してください。

- アドインが Microsoft Edge Chromium (WebView2) の埋め込み Web ビュー コントロールで実行される場合は、「[Visual Studio Code と Microsoft Edge WebView2 を使用して Windows でアドインをデバッグする (Chromium ベース)」](../testing/debug-desktop-using-edge-chromium.md)を参照してください。

## <a name="see-also"></a>関連項目

- [Outlook のアドイン コマンド](add-in-commands-for-outlook.md)
- [Office アドインのデバッグの概要](../testing/debug-add-ins-overview.md)
- [イベント ベースの Outlook アドインをデバッグする](debug-autolaunch.md)
