---
title: UI レス Outlook アドインをデバッグする
description: UI レス Outlook アドインをデバッグする方法について説明します。
ms.topic: article
ms.date: 05/19/2022
ms.localizationpriority: medium
ms.openlocfilehash: 33aa36f86b7a163e650a23296b4c35aca7cb5492
ms.sourcegitcommit: fcb8d5985ca42537808c6e4ebb3bc2427eabe4d4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/24/2022
ms.locfileid: "65650712"
---
# <a name="debug-your-ui-less-outlook-add-in"></a>UI レス Outlook アドインをデバッグする

この記事では、Visual Studio Code で Office アドイン デバッガー拡張機能を使用して [UI レス Outlook アドイン](add-in-commands-for-outlook.md#executing-a-javascript-function)をデバッグする方法について説明します。UI レスアドイン アクションは、リボンのアドイン コマンド ボタンを使用して開始されます。 アドイン コマンドの詳細については、「[Outlookのアドイン コマンド」を](add-in-commands-for-outlook.md)参照してください。

この記事では、デバッグするアドイン プロジェクトが既にあることを前提としています。 UI レス アドインを作成してデバッグを実行するには、「[チュートリアル: メッセージ作成Outlookアドインをビルド](../tutorials/outlook-tutorial.md)する」の手順に従います。

## <a name="mark-your-add-in-for-debugging"></a>デバッグ用にアドインをマークする

アドインプロジェクトを作成するために[アドインをOfficeするために Yeoman ジェネレーター](../develop/yeoman-generator-overview.md)を使用した場合は、この記事の後半の[「デバッガーの構成と実行](#configure-and-run-the-debugger)」セクションに進んでください。 アドインをビルドしてローカル サーバーを起動するために実行`npm start`すると、このコマンドは、デバッグ用にアドインを`UseDirectDebugger``HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]`マークするレジストリ キーの値も設定します。

それ以外の場合は、別のツールを使用してアドインを作成した場合は、次の手順を実行します。

1. レジストリ キーに `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]` 移動します。 アドインのマニフェストの **ID に** 置き換えます`[Add-in ID]`。

    [!include[Developer registry key](../includes/developer-registry-key.md)]

1. キーの値`1`を `UseDirectDebugger` .

## <a name="configure-and-run-the-debugger"></a>デバッガーを構成して実行する

アドインでデバッグを有効にしたので、デバッガーを構成して実行する準備ができました。 これを行う方法については、ランタイムに適用される次のいずれかのオプションを選択します。

- アドインが WebView ランタイムで実行される場合は、「[Visual Studio Code のアドイン デバッガー拡張機能Microsoft Office](../testing/debug-with-vs-extension.md)参照してください。

- アドインが Microsoft Edge Chromium WebView2 ランタイムで実行される場合は、「Visual Studio [Code と Microsoft Edge WebView2 (Chromium ベース)を使用してWindowsでアドインをデバッグする」を](../testing/debug-desktop-using-edge-chromium.md)参照してください。

## <a name="see-also"></a>関連項目

- [Outlook のアドイン コマンド](add-in-commands-for-outlook.md)
- [Office アドインのデバッグの概要](../testing/debug-add-ins-overview.md)
- [イベント ベースのOutlook アドインをデバッグする](debug-autolaunch.md)
