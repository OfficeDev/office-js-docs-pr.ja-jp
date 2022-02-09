---
title: 従来の WebView (EdgeHTML) WindowsおよびVisual Studio CodeをMicrosoft Edgeアドインをデバッグする
description: Office VS Code の Office アドイン デバッガー拡張機能を使用して、Microsoft Edge 従来版 WebView (EdgeHTML) を使用するアドインをデバッグする方法について説明します。
ms.date: 02/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: 11b728f9b3f467017711c9d75cfd07767957deae
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467695"
---
# <a name="microsoft-office-add-in-debugger-extension-for-visual-studio-code"></a>Visual Studio Code 用 Microsoft Office アドイン デバッガー拡張機能

Office Windows で実行されているアドインでは、Visual Studio Code の Office アドイン デバッガー拡張機能を使用して、Microsoft Edge 従来版 に対して元の WebView (EdgeHTML) ランタイムでデバッグできます。 

> [!IMPORTANT]
> この記事は、「Office アドインで使用されるブラウザー」で説明したように、Office が元の WebView (EdgeHTML Office) ランタイムでアドインを実行[する場合にのみ適用](../concepts/browsers-used-by-office-web-add-ins.md)されます。Visual Studio Microsoft Edge WebView2 (Chromium ベース) に対する Visual Studio コードでのデバッグの手順については、「Microsoft Office アドイン デバッガー拡張機能 for Visual Studio Code」[を参照してください](debug-desktop-using-edge-chromium.md)。

> [!TIP]
> Visual Studio Code に組み込みのツールを使用してデバッグできない、またはデバッグしない場合、またはアドインが Visual Studio Code の外部で実行されている場合にのみ発生する問題が発生する場合は、「開発ツールを使用してアドインをデバッグする」の説明に従ってエッジ レガシ (EdgeHTML) ランタイムをデバッグできます。[Microsoft Edge 従来版](debug-add-ins-using-devtools-edge-legacy.md)。

このデバッグ モードは動的で、コードの実行中にブレークポイントを設定できます。 デバッガーが接続されている間、コード内の変更をすぐに確認できます。すべてデバッグ セッションを失う必要はありません。 コードの変更も保持されます。そのため、コードに対する複数の変更の結果を確認できます。 次の図は、この拡張機能の動作を示しています。

![Officeアドインのセクションをデバッグするアドイン デバッガー拡張機能Excelデバッグします。](../images/vs-debugger-extension-for-office-addins.jpg)

## <a name="prerequisites"></a>前提条件

- [Visual Studio Code](https://code.visualstudio.com/)
- [Node.js (バージョン 10 以上)](https://nodejs.org/)
- Windows 10, 11
- [Microsoft Edge](https://www.microsoft.com/edge) Office アドインで使用されるブラウザーで説明したように、Microsoft Edge 従来版 と元の Web ビュー (EdgeHTML [Office](../concepts/browsers-used-by-office-web-add-ins.md)) をサポートするプラットフォームと Office アプリケーションの組み合わせ。

## <a name="install-and-use-the-debugger"></a>デバッガーをインストールして使用する

これらの手順では、コマンド ラインの使用経験、基本的な JavaScript の理解、および Yo Office ジェネレーターを使用する前に Office アドイン プロジェクトを作成したと仮定します。 前にこれを行ったことがない場合は、次のようなチュートリアルの 1 つを参照Excel Office[検討してください](../tutorials/excel-tutorial.md)。

1. 最初の手順は、プロジェクトと作成方法によって異なります。

   - プロジェクトを作成してデバッグを試す場合は、Visual Studio Codeアドインに [Yeoman ジェネレーター Office使用します](https://github.com/OfficeDev/generator-office)。これを行うには、クイック スタート ガイド (Outlook アドイン のクイック スタート[](../quickstarts/outlook-quickstart.md)など) のいずれかを使用します。 
   - Yo を使用して作成された既存のプロジェクトをデバッグする場合は、Officeに進みます。
   - Yo Office で作成されていない既存のプロジェクトをデバッグする場合は、「[付録](#appendix)」の手順を実行し、次の手順に戻ってください。


1. プロジェクトVS Code開き、プロジェクトを開きます。 

1. VS Code で **Ctrl キー + Shift キー + X キー** を選択して、拡張機能バーを開きます。 "Microsoft Office アドイン デバッガー" 拡張機能を検索してインストールします。

1. [  **ファイルの表示>実行] を** 選択するか、 **Ctrl + Shift + D** と入力してデバッグ ビューに切り替えます。

1. [**RUN AND DEBUG]** オプションから、ホスト アプリケーションの [エッジ レガシ] オプション (デスクトップ (エッジ レガシ) などOutlook **を選択します**。 **F5** キーを選択するか、メニューから **[実行]、[デバッグの開始]** の順に選択してデバッグを開始します。 この操作により、アドインをホストするローカル サーバーがノード ウィンドウで自動的に起動され、Excel や Word などのホスト アプリケーションが自動的に開きます。 これには数秒かかる場合があります。

1. ホスト アプリケーションで、アドインを使用する準備ができました。 **[作業ウィンドウの表示]** を選択するか、その他のアドイン コマンドを実行します。 次のようなダイアログ ボックスが表示されます。

   > WebView は読み込み時に停止します。
   > WebView をデバッグするには、Microsoft Debugger for Edge 拡張機能VS Code WebView インスタンスに接続し、[**OK**] をクリックして続行します。 このダイアログが今後表示されるのを防ぐには、[キャンセル] を **クリックします**。

   **[OK]** を選択します。

   > [!NOTE]
   > **[キャンセル]** を選択すると、このアドインのインスタンスの実行中はダイアログが表示されなくなります。 ただし、アドインを再起動すると、ダイアログはもう一度表示されます。

1. プロジェクトの作業ウィンドウ ファイルにブレークポイントを設定します。 ブレークポイントを設定するには、Visual Studio Code行の横にマウス ポインターを置き、表示される赤い円を選択します。

    ![赤い円は、次のコード行にVisual Studio Code。](../images/set-breakpoint.jpg)

1. ブレークポイントを使用して行を呼び出す機能をアドインで実行します。 ブレークポイントがヒットし、ローカル変数を検査できます。

   > [!NOTE]
   > `Office.initialize` または `Office.onReady` の呼び出しのブレークポイントは無視されます。 これらのメソッドの詳細については、「 [Office アドインを初期化する](../develop/initialize-add-in.md)」 を参照してください。

> [!IMPORTANT]
> デバッグ セッションを停止する最善の方法は、**Shift キーを押しながら F5 キー** を押すか、メニューから **[実行] > [デバッグの停止]** を選択することです。 この操作では、ノード サーバー ウィンドウを閉じてホスト アプリケーションを閉じようとしますが、ドキュメントを保存するかどうかを確認するプロンプトがホスト アプリケーションに表示されます。 適切な選択を行い、ホスト アプリケーションを閉じます。 ノード ウィンドウまたはホスト アプリケーションを手動で閉じないようにします。 これを行うと、特にデバッグ セッションの停止と開始を繰り返している時に、バグが発生する可能性があります。
>
> デバッグが動作を停止する場合、たとえば、ブレークポイントが無視される場合などは、デバッグを停止します。 その後、必要に応じて、すべてのホスト アプリケーション ウィンドウとノード ウィンドウを閉じます。 最後に、Visual Studio Code を閉じて、もう一度開きます。

### <a name="appendix"></a>付録

プロジェクトが Yo Officeで作成されていない場合は、プロジェクトのデバッグ構成を作成Visual Studio Code。 

1. プロジェクトのフォルダーに `launch.json` ファイルが `\.vscode` 存在しない場合は、そのフォルダーに名前の付いたファイルを作成します。 
1. ファイルに配列が含けられているか確認 `configurations` します。 次に示すのは、 の簡単な例です `launch.json`。

    ```json
    {
      // other properities may be here.

      "configurations": [

        // configuration objects may be here.

      ]

      //other properies may be here.
    }
    ```

1. 配列に次のオブジェクトを追加 `configurations` します。

    ```json
    {
      "name": "$HOST$ Desktop (Edge Legacy)",
      "type": "office-addin",
      "request": "attach",
      "url": "https://localhost:3000/taskpane.html?_host_Info=Excel$Win32$16.01$en-US$$$$0",
      "port": 9222,
      "timeout": 600000,
      "webRoot": "${workspaceRoot}",
      "preLaunchTask": "Debug: Excel Desktop",
      "postDebugTask": "Stop Debug"
    }
    ```

1. プレースホルダーを、`$HOST$`アドインが実行Officeアプリケーションの`Outlook`名前に置き換える`Word`。たとえば、.
1. ファイルを保存して閉じます。

## <a name="see-also"></a>関連項目

- [Office アドインのテストとデバッグ](test-debug-office-add-ins.md)
- [WebView2 (WindowsベースWindows WebView2 Visual Studio Code Microsoft Edgeを使用して、Chromiumアドインをデバッグします](debug-desktop-using-edge-chromium.md)。
- [Internet Explorer の開発者ツールを使用してアドインをデバッグする](debug-add-ins-using-f12-tools-ie.md)
- [Edge レガシー用の開発者ツールを使用してアドインをデバッグする](debug-add-ins-using-devtools-edge-legacy.md)
- [Microsoft Edge (Chromium ベース)で開発者ツールを使用してアドインをデバッグする](debug-add-ins-using-devtools-edge-chromium.md)
- [作業ウィンドウからデバッガーをアタッチする](attach-debugger-from-task-pane.md)
