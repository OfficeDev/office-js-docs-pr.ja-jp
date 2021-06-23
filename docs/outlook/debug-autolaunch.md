---
title: イベント ベースのアドインOutlookデバッグする (プレビュー)
description: イベント ベースのアクティブ化を実装Outlookアドインをデバッグする方法について説明します。
ms.topic: article
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: 8cabbb669d9b46e047efa7e79ae4225c1fc22689
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53077093"
---
# <a name="debug-your-event-based-outlook-add-in-preview"></a>イベント ベースのアドインOutlookデバッグする (プレビュー)

この記事では、アドインでイベント ベースの [ライセンス認証を](autolaunch.md) 実装する場合のデバッグ ガイダンスを提供します。 イベント ベースのアクティブ化機能は現在プレビュー中です。

> [!IMPORTANT]
> このデバッグ機能は、サブスクリプションを使用したOutlookのWindowsプレビューでのみMicrosoft 365されます。 詳細については、この記事の「イベント ベースのアクティブ [化機能の](#preview-debugging-for-the-event-based-activation-feature) プレビュー デバッグ」セクションを参照してください。

この記事では、デバッグを有効にする重要な段階について説明します。

- [デバッグ用にアドインをマークする](#mark-your-add-in-for-debugging)
- [構成Visual Studio Code](#configure-visual-studio-code)
- [添付Visual Studio Code](#attach-visual-studio-code)
- [Debug](#debug)

アドイン プロジェクトを作成するには、いくつかのオプションがあります。 使用しているオプションによっては、手順が異なる場合があります。 このような場合は、Office アドインに Yeoman ジェネレーターを使用してアドイン プロジェクトを作成した場合 (たとえば、イベント ベースのライセンス認証のチュートリアルを実行します)、office のヨーヨー 手順に従い、それ以外の場合は、その他の手順に従います。 [](autolaunch.md) Visual Studio Codeバージョン 1.56.1 以上である必要があります。

## <a name="preview-debugging-for-the-event-based-activation-feature"></a>イベント ベースのアクティブ化機能のデバッグをプレビューする

イベント ベースのアクティブ化機能のデバッグ機能を試してみてください。 このページの最後にある「フィードバック」セクションをGitHubフィードバックを提供することで、お客様のシナリオと改善方法をお知らせします。

この機能を Outlook Windowsでプレビューするには、必要な最小ビルドは 16.0.13729.20000 です。 ベータビルドへのアクセスOffice、Insider プログラムOffice[参加してください](https://insider.office.com)。

## <a name="mark-your-add-in-for-debugging"></a>デバッグ用にアドインをマークする

1. レジストリ キーを設定します `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger` 。 `[Add-in ID]` は **、アドイン** マニフェストの ID です。

    **yo office**: コマンド ライン ウィンドウで、アドイン フォルダーのルートに移動し、次のコマンドを実行します。

    ```command&nbsp;line
    npm start
    ```

    コードを構築し、ローカル サーバーを起動する以外に、このコマンドは、このアドインのレジストリ キーを `UseDirectDebugger` に設定する必要があります `1` 。

    **その他**: の下 `UseDirectDebugger` にレジストリ キーを追加 `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]\` します。 アドイン `[Add-in ID]` マニフェスト **の Id** に置き換える。 レジストリ キーをに設定します `1` 。

    [!include[Developer registry key](../includes/developer-registry-key.md)]

1. デスクトップOutlook起動します (またはOutlook開いている場合は再起動します)。
1. 新しいメッセージまたは予定を作成します。 次のダイアログが表示されます。 ダイアログ *を* まだ操作しないでください。

    ![イベント ベースのハンドラー のデバッグ ダイアログのスクリーンショット。](../images/outlook-win-autolaunch-debug-dialog.png)

## <a name="configure-visual-studio-code"></a>構成Visual Studio Code

### <a name="yo-office"></a>yo office

1. コマンド ライン ウィンドウに戻り、コマンド ライン ウィンドウVisual Studio Code。

    ```command&nbsp;line
    code .
    ```

1. このVisual Studio Code **./.vscode/launch.js** を開き、構成の一覧に次の抜粋を追加します。 変更内容を保存します。

    ```json
    {
      "name": "Direct Debugging",
      "type": "node",
      "request": "attach",
      "port": 9229,
      "protocol": "inspector",
      "timeout": 600000,
      "trace": true
    }
    ```

### <a name="other"></a>その他

1. デバッグという名前の新しい **フォルダーを作成** します (おそらくデスクトップ フォルダー **に)。**
1. Visual Studio Code を開きます。
1. [ファイルを **開**  >  **くフォルダー]** に移動し、作成したフォルダーに移動し、[フォルダーの選択]**を選択します**。
1. [アクティビティ バー] で、[デバッグ] **アイテム** (Ctrl + Shift + D) を選択します。

    ![アクティビティ バーの [デバッグ] アイコンのスクリーンショット。](../images/vs-code-debug.png)

1. [ファイルに **対してlaunch.jsを作成する] リンクを選択** します。

    ![ページ内のファイルにlaunch.jsを作成するリンクVisual Studio Code。](../images/vs-code-create-launch.json.png)

1. [環境 **の選択] ドロップダウン** で、[ **エッジ:** 起動] を選択して、launch.jsを作成します。
1. 構成の一覧に次の抜粋を追加します。 変更内容を保存します。

    ```json
    {
      "name": "Direct Debugging",
      "type": "node",
      "request": "attach",
      "port": 9229,
      "protocol": "inspector",
      "timeout": 600000,
      "trace": true
    }
    ```

## <a name="attach-visual-studio-code"></a>添付Visual Studio Code

1. アドインのbundle.jsを見 **つける** には、Windows エクスプローラーで次のフォルダーを開き、アドインの **ID** (マニフェストにある) を検索します。

    ```text
    %LOCALAPPDATA%\Microsoft\Office\16.0\Wef
    ```

    この ID のプレフィックスが付いたフォルダーを開き、完全なパスをコピーします。 このVisual Studio Code、その **フォルダーbundle.js** を開きます。 ファイル パスのパターンは次のとおりです。

    `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\{[Outlook profile GUID]}\[encoding]\Javascript\[Add-in ID]_[Add-in Version]_[locale]\bundle.js`

1. デバッガーを停止bundle.js場所にブレークポイントを配置します。
1. [ **デバッグ] ドロップダウンで** 、[直接デバッグ] という **名前を選択し**、[実行] を **選択します**。

    ![[デバッグ] ドロップダウンの構成オプションから [直接デバッグ] を選択Visual Studio Codeスクリーンショット。](../images/outlook-win-autolaunch-debug-vsc.png)

## <a name="debug"></a>Debug

1. デバッガーが接続されているのを確認した後、Outlook に戻り、[イベント ベースのハンドラーのデバッグ] ダイアログで **[OK] を選択します**。

1. これで、イベント ベースのアクティブ化コードVisual Studio Codeデバッグを有効にすることで、ブレークポイントをヒットできます。

## <a name="stop-debugging"></a>デバッグを停止する

現在のデスクトップ セッションの残りのOutlook停止するには、[イベント ベースのハンドラーのデバッグ] ダイアログで、[キャンセル] を **選択します**。 デバッグを再び有効にするには、デスクトップOutlookします。

イベント ベースの **ハンドラー の** デバッグ ダイアログがポップアップし、後続の Outlook セッションのデバッグを停止するには、関連付けられたレジストリ キーを削除するか、その値を : に設定します `0` `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger` 。

## <a name="see-also"></a>関連項目

- [イベント ベースのOutlook用にアドインを構成する](autolaunch.md)
- [ランタイム ログを使用してアドインをデバッグする](../testing/runtime-logging.md#runtime-logging-on-windows)
