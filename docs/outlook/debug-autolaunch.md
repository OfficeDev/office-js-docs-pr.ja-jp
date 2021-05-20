---
title: イベント ベースのOutlook アドインのデバッグ (プレビュー)
description: イベント ベースのアクティブ化を実装するOutlook アドインをデバッグする方法について説明します。
ms.topic: article
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: d7621a7407db3b8e773d1534beb6c881f7b48558
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555286"
---
# <a name="debug-your-event-based-outlook-add-in-preview"></a>イベント ベースのOutlook アドインのデバッグ (プレビュー)

この記事では、アドインにイベント [ベースのアクティブ化](autolaunch.md) を実装する際のデバッグ ガイダンスを提供します。 イベントベースのアクティブ化機能は現在プレビュー段階です。

> [!IMPORTANT]
> このデバッグ機能は、Microsoft 365 サブスクリプションを使用するWindowsでOutlookでプレビューする場合にのみサポートされます。 詳細については、この記事の「 [イベントベースのアクティブ化機能のプレビュー デバッグ](#preview-debugging-for-the-event-based-activation-feature) 」セクションを参照してください。

この記事では、デバッグを有効にする主要なステージについて説明します。

- [アドインにデバッグ用のマークを付けます](#mark-your-add-in-for-debugging)
- [Visual Studio Codeの構成](#configure-visual-studio-code)
- [Visual Studio Codeを添付する](#attach-visual-studio-code)
- [Debug](#debug)

アドイン プロジェクトを作成するには、いくつかのオプションがあります。 使用するオプションによって、手順が異なる場合があります。 この場合、アドインの Office に Yeoman ジェネレーターを使用してアドイン プロジェクトを作成した場合 ([たとえば、イベント ベースのアクティブ化のチュートリアル](autolaunch.md)を実行する場合)、**ヨーオフィス** の手順に従って、それ以外の場合は「**その他** の手順」に従います。 Visual Studio Codeは、少なくともバージョン 1.56.1 である必要があります。

## <a name="preview-debugging-for-the-event-based-activation-feature"></a>イベント ベースのアクティブ化機能のプレビュー デバッグ

イベントベースのアクティブ化機能のデバッグ機能を試してみるようお勧めします。 GitHubを通じてフィードバックを提供することで、お客様のシナリオと改善方法をお知らせください(このページの最後にある **フィードバック** セクションを参照)。

WindowsでOutlookに対してこの機能をプレビューするには、最小必要なビルドは 16.0.13729.2000 です。 ベータ版ビルドOfficeアクセスするには[、Office Insider プログラム](https://insider.office.com)に参加してください。

## <a name="mark-your-add-in-for-debugging"></a>アドインにデバッグ用のマークを付けます

1. レジストリ キーを設定 `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger` する: `[Add-in ID]` はアドイン マニフェストの **ID** です。

    **yo office**: コマンド ライン ウィンドウで、アドイン フォルダーのルートに移動し、次のコマンドを実行します。

    ```command&nbsp;line
    npm start
    ```

    このコマンドでは、コードのビルドとローカル サーバーの起動に加えて、 `UseDirectDebugger` このアドインのレジストリ キーを に設定する必要があります `1` 。

    **その他**: レジストリ キーを `UseDirectDebugger` に追加 `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]\` します。 `[Add-in ID]`アドイン マニフェストの **ID** で置き換えます。 レジストリ キーを `1` に設定します。

    [!include[Developer registry key](../includes/developer-registry-key.md)]

1. デスクトップOutlook起動します (または、既に開いている場合はOutlookを再起動します)。
1. 新しいメッセージまたは予定を作成します。 次のダイアログが表示されます。 ダイアログをまだ操作 *しないでください* 。

    ![デバッグ イベント ベースのハンドラー ダイアログのスクリーンショット](../images/outlook-win-autolaunch-debug-dialog.png)

## <a name="configure-visual-studio-code"></a>Visual Studio Codeの構成

### <a name="yo-office"></a>ヨーオフィス

1. コマンド ライン ウィンドウに戻り、Visual Studio Code開きます。

    ```command&nbsp;line
    code .
    ```

1. Visual Studio Codeで **、./.vscode/launch.jsのファイルを** 開き、次の抜粋を構成のリストに追加します。 変更内容を保存します。

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

1. **デバッグ** という名前の新しいフォルダーを作成します (**デスクトップ** フォルダーの場合があります)。
1. Visual Studio Code を開きます。
1. [**ファイル**  >  **を開くフォルダ]** に移動し、作成したフォルダに移動して、[**フォルダの選択**] を選択します。
1. アクティビティ バーで、[ **デバッグ** ] 項目 (Ctrl + Shift + D) を選択します。

    ![アクティビティ バーのデバッグ アイコンのスクリーンショット](../images/vs-code-debug.png)

1. [ **ファイルにlaunch.jsを作成]リンクを選択します** 。

    ![Visual Studio Codeでファイルにlaunch.jsを作成するためのリンクのスクリーンショット](../images/vs-code-create-launch.json.png)

1. [ **環境の選択** ] ドロップダウンで、[ **エッジ: 起動** ] を選択してファイルにlaunch.jsを作成します。
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

## <a name="attach-visual-studio-code"></a>Visual Studio Codeを添付する

1. アドインの **bundle.js** を検索するには、Windows エクスプローラーで次のフォルダーを開き、アドインの **ID** (マニフェスト内にあります) を検索します。

    ```text
    %LOCALAPPDATA%\Microsoft\Office\16.0\Wef
    ```

    この ID のプレフィックスが付いたフォルダーを開き、その完全なパスをコピーします。 Visual Studio Codeで、そのフォルダから **bundle.js** 開きます。 ファイル パスのパターンは次のようになります。

    `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\{[Outlook profile GUID]}\[encoding]\Javascript\[Add-in ID]_[Add-in Version]_[locale]\bundle.js`

1. デバッガーを停止する位置bundle.jsにブレークポイントを配置します。
1. **[DEBUG]** ドロップダウンで、[**直接デバッグ**] という名前を選択し、[**実行**] を選択します。

    ![[Visual Studio Codeデバッグ] ドロップダウンの構成オプションから直接デバッグを選択するスクリーンショット](../images/outlook-win-autolaunch-debug-vsc.png)

## <a name="debug"></a>Debug

1. デバッガーがアタッチされていることを確認したら、Outlookに戻り、イベント ベースの **デバッグ ハンドラー** ダイアログで **[OK] を** クリックします。

1. Visual Studio Codeでブレークポイントにヒットして、イベントベースのアクティブ化コードをデバッグできるようになりました。

## <a name="stop-debugging"></a>デバッグを停止する

現在のOutlookデスクトップ セッションの残りの部分のデバッグを停止するには、[**イベント ベースのデバッグ ハンドラー** ] ダイアログ ボックスで [**キャンセル**] をクリックします。 デバッグを再度有効にするには、デスクトップOutlook再起動します。

**イベントに基づくデバッグ ハンドラ** ダイアログがポップアップして、後続のOutlook セッションのデバッグを停止しないようにするには、関連付けられたレジストリ キーを削除するか、その値を : に設定 `0` `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger` します。

## <a name="see-also"></a>関連項目

- [イベント ベースのアクティブ化用にOutlook アドインを構成する](autolaunch.md)
- [ランタイム ログを使用してアドインをデバッグする](../testing/runtime-logging.md#runtime-logging-on-windows)
