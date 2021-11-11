---
title: Visual Studio Code と Microsoft Edge WebView2 を使用して Windows でアドインをデバッグする (Chromium ベース)
description: VS Code で拡張機能 Debugger for Microsoft Edge を使用し、Microsoft Edge WebView2 (Chromium ベース) を使用した Office アドインをデバッグする方法について説明します。
ms.date: 11/09/2021
ms.localizationpriority: high
ms.openlocfilehash: 2ffc9226cb5e4fb38c88a98a79f3676ca3b6071e
ms.sourcegitcommit: 3d37c42f5e465dac52d231d31717bdbb3bfa0e30
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/10/2021
ms.locfileid: "60889987"
---
# <a name="debug-add-ins-on-windows-using-visual-studio-code-and-microsoft-edge-webview2-chromium-based"></a>Visual Studio Code と Microsoft Edge WebView2 を使用して Windows でアドインをデバッグする (Chromium ベース)

Windows 上で動作する Office アドインは、Visual Studio Code の Debugger for Microsoft Edge の拡張機能を使用することで、Edge Chromium WebView2 ランタイムに対してデバッグを行うことができます。 

> [!TIP]
> Visual Studio Code に組み込まれているツールを使用してデバッグできない場合、またはデバッグしたくない場合、またはアドインが Visual Studio Code の外部で実行されている場合にのみ問題が発生した場合は、「[Microsoft Edge WebView2 用の開発者ツールを使用してアドインをデバッグする](debug-add-ins-using-devtools-edge-chromium.md)」の説明に従って、Edge (Chromium ベース) 開発者ツールを使用して Edge Chromium WebView2 ランタイムをデバッグできます。

## <a name="prerequisites"></a>前提条件

- [Visual Studio Code](https://code.visualstudio.com/) (管理者として実行する必要があります)
- [Node.js (バージョン 10 以上)](https://nodejs.org/)
- Windows 10, 11
- [Office アドインで使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md)で説明されているように、プラットフォームとOfficeアプリケーションの組み合わせが、WebView2 (Chromium ベース) で Microsoft Edge をサポートしています。Microsoft 365 のバージョンが 2101 より前の場合は、WebView2 をインストールする必要があります。 [Microsoft Edge WebView2 / Embed web content ... with Microsoft Edge WebView2](https://developer.microsoft.com/microsoft-edge/webview2/) でのインストール方法はこちらをご覧ください。

## <a name="install-and-use-the-debugger"></a>デバッガーをインストールして使用する

1. [Office アドイン用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)を使用してプロジェクトを作成してください。これを行うには、「[Outlook アドインのクイック スタート](../quickstarts/outlook-quickstart.md)」などのクイック スタート ガイドのいずれかをご利用ください。

   > [!TIP]
   > Yeoman ジェネレーター ベースのアドインを使用していない場合は、レジストリ キーの調整を求められる場合があります。 プロジェクトのルート フォルダーで、コマンド ラインを使用して以下を実行します。
   >
   > ``` command&nbsp;line
   > npx office-addin-debugging start <your manifest path>
   > ```

1. VS Code でプロジェクトを開きます。 VS Code で **Ctrl キー + Shift キー + X キー** を選択して、拡張機能バーを開きます。 「Debugger for Microsoft Edge」で拡張機能を検索し、これをインストールします。

1. 次に、**[表示]、[実行]** の順に選択するか、**Ctrl キー + Shift キー + D キー** を入力してデバッグ ビューに切り替えます。

1. **RUN AND DEBUG** オプションから、**Excel Desktop (Edge Chromium)** などのホスト アプリケーション用に Edge Chromium オプションを選択します。 **F5** キーを選択するか、メニューから **[実行]、[デバッグの開始]** の順に選択してデバッグを開始します。 この操作により、アドインをホストするローカル サーバーがノード ウィンドウで自動的に起動され、Excel や Word などのホスト アプリケーションが自動的に開きます。 これには数秒かかる場合があります。

1. ホスト アプリケーションで、アドインを使用する準備ができました。 **[作業ウィンドウの表示]** を選択するか、その他のアドイン コマンドを実行します。 ダイアログ ボックスが表示され、以下が表示されます。

   > WebView は読み込み時に停止します。
   > WebView をデバッグするには、拡張機能 Microsoft Debugger for Edge を使用して VS Code を WebView のインスタンスにアタッチし、[OK] をクリックして続行します。 今後このダイアログが表示されないようにするには、[キャンセル] をクリックします。

   **[OK]** を選択します。

   > [!NOTE]
   > **[キャンセル]** を選択すると、このアドインのインスタンスの実行中はダイアログが表示されなくなります。 ただし、アドインを再起動すると、ダイアログはもう一度表示されます。

1. これで、プロジェクトのコードにブレークポイントを設定し、デバッグを実行できるようになりました。

   > [!NOTE]
   > `Office.initialize` または `Office.onReady` の呼び出しのブレークポイントは無視されます。 これらのメソッドの詳細については、「 [Office アドインを初期化する](../develop/initialize-add-in.md)」 を参照してください。

> [!IMPORTANT]
> デバッグ セッションを停止する最善の方法は、**Shift キーを押しながら F5 キー** を押すか、メニューから **[実行] > [デバッグの停止]** を選択することです。 この操作では、ノード サーバー ウィンドウを閉じてホスト アプリケーションを閉じようとしますが、ドキュメントを保存するかどうかを確認するプロンプトがホスト アプリケーションに表示されます。 適切な選択を行い、ホスト アプリケーションを閉じます。 ノード ウィンドウまたはホスト アプリケーションを手動で閉じないようにします。 これを行うと、特にデバッグ セッションの停止と開始を繰り返している時に、バグが発生する可能性があります。
>
> デバッグが動作を停止する場合、たとえば、ブレークポイントが無視される場合などは、デバッグを停止します。 その後、必要に応じて、すべてのホスト アプリケーション ウィンドウとノード ウィンドウを閉じます。 最後に、Visual Studio Code を閉じて、もう一度開きます。

## <a name="see-also"></a>関連項目

- [Office アドインのテストとデバッグ](test-debug-office-add-ins.md)
- [Visual Studio Code 用 Microsoft Office アドイン デバッガー拡張機能](debug-with-vs-extension.md)
- [作業ウィンドウからデバッガーをアタッチする](attach-debugger-from-task-pane.md)
