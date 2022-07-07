---
title: イベント ベースの Outlook アドインをデバッグする
description: イベント ベースのアクティブ化を実装する Outlook アドインをデバッグする方法について説明します。
ms.topic: article
ms.date: 04/28/2022
ms.localizationpriority: medium
ms.openlocfilehash: 8dbd74036cf56b5ff492315f928324a3aa1e7312
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/06/2022
ms.locfileid: "66659683"
---
# <a name="debug-your-event-based-outlook-add-in"></a>イベント ベースの Outlook アドインをデバッグする

この記事では、アドインに [イベント ベースのアクティブ化](autolaunch.md) を実装する際のデバッグ ガイダンスを提供します。 イベント ベースのアクティブ化機能は [要件セット 1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10) で導入され、プレビューで追加のイベントが利用可能になりました。 詳細については、「 [サポートされているイベント](autolaunch.md#supported-events)」を参照してください。

> [!IMPORTANT]
> このデバッグ機能は、Microsoft 365 サブスクリプションを使用する Outlook on Windows でのみサポートされます。

この記事では、デバッグを有効にする主要な段階について説明します。

- [デバッグ用にアドインをマークする](#mark-your-add-in-for-debugging)
- [Visual Studio Code を構成する](#configure-visual-studio-code)
- [Visual Studio Code をアタッチする](#attach-visual-studio-code)
- [Debug](#debug)

アドイン プロジェクトを作成するには、いくつかのオプションがあります。 使用しているオプションによっては、手順が異なる場合があります。 この場合、Office アドイン用 Yeoman ジェネレーターを使用してアドイン プロジェクトを作成した場合 (たとえば、 [イベント ベースのアクティブ化チュートリアル](autolaunch.md)を実行するなど)、 **yo office** の手順に従い、それ以外の場合は **他** の手順に従います。 Visual Studio Code は、少なくともバージョン 1.56.1 である必要があります。

## <a name="mark-your-add-in-for-debugging"></a>デバッグ用にアドインをマークする

1. レジストリ キーを設定します `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger`。 `[Add-in ID]`**\<Id\>** はアドイン マニフェスト内です。

    **yo office**: コマンド ライン ウィンドウで、アドイン フォルダーのルートに移動し、次のコマンドを実行します。

    ```command&nbsp;line
    npm start
    ```

    このコマンドは、コードをビルドしてローカル サーバーを起動するだけでなく、このアドイン`1`のレジストリ キーを `UseDirectDebugger` .

    **その他**: 下にレジストリ キー`HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]\`を`UseDirectDebugger`追加します。 **\<Id\>** アドイン マニフェストから置き換えます`[Add-in ID]`。 レジストリ キー `1`を .

    [!include[Developer registry key](../includes/developer-registry-key.md)]

1. Outlook デスクトップを起動します (既に開いている場合は Outlook を再起動します)。
1. 新しいメッセージまたは予定を作成します。 次のダイアログが表示されます。 ダイアログをまだ操作 *しないでください* 。

    ![[デバッグ イベント ベースのハンドラー] ダイアログのスクリーンショット。](../images/outlook-win-autolaunch-debug-dialog.png)

## <a name="configure-visual-studio-code"></a>Visual Studio Code を構成する

### <a name="yo-office"></a>yo office

1. コマンド ライン ウィンドウに戻り、Visual Studio Code を開きます。

    ```command&nbsp;line
    code .
    ```

1. Visual Studio Code で **、./.vscode/launch.json** ファイルを開き、構成の一覧に次の抜粋を追加します。 変更内容を保存します。

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

1. **[デバッグ**] という名前の新しいフォルダーを作成します (**デスクトップ** フォルダーなど)。
1. Visual Studio Code を開きます。
1. **[ファイル** > **を開くフォルダー] に** 移動し、作成したフォルダーに移動し、[**フォルダーの選択] を選択します**。
1. アクティビティ バーで、[ **デバッグ** ] 項目 (Ctrl + Shift + D) を選択します。

    ![アクティビティ バーの [デバッグ] アイコンのスクリーンショット。](../images/vs-code-debug.png)

1. **launch.json ファイルの作成リンクを選択します**。

    ![Visual Studio Code で launch.json ファイルを作成するためのリンクのスクリーンショット。](../images/vs-code-create-launch.json.png)

1. [ **環境の選択]** ドロップダウンで、[ **エッジ: 起動** ] を選択して launch.json ファイルを作成します。
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

## <a name="attach-visual-studio-code"></a>Visual Studio Code をアタッチする

1. アドインの **bundle.js** を見つけるには、Windows エクスプローラーで次のフォルダーを開き、アドイン (マニフェストにあります) を検索します **\<Id\>** 。

    ```text
    %LOCALAPPDATA%\Microsoft\Office\16.0\Wef
    ```

    この ID のプレフィックスが付いたフォルダーを開き、その完全なパスをコピーします。 Visual Studio Code で、そのフォルダーから **bundle.js** を開きます。 ファイル パスのパターンは次のようになります。

    `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\{[Outlook profile GUID]}\[encoding]\Javascript\[Add-in ID]_[Add-in Version]_[locale]\bundle.js`

1. デバッガーを停止するbundle.jsにブレークポイントを配置します。
1. **[デバッグ**] ドロップダウンで、[**Direct Debuging**] という名前を選択し、[**実行**] を選択します。

    ![Visual Studio Code Debug ドロップダウンの構成オプションから直接デバッグを選択するスクリーンショット。](../images/outlook-win-autolaunch-debug-vsc.png)

## <a name="debug"></a>デバッグ

1. デバッガーがアタッチされていることを確認したら、Outlook に戻り、[ **デバッグ イベント ベースのハンドラー** ] ダイアログで [OK] を選択 **します** 。

1. Visual Studio Code でブレークポイントにヒットし、イベント ベースのアクティブ化コードをデバッグできるようになりました。

## <a name="stop-debugging"></a>デバッグの停止

現在の Outlook デスクトップ セッションの残りの部分のデバッグを停止するには、[ **デバッグ イベント ベースのハンドラー** ] ダイアログで [ **キャンセル**] を選択します。 デバッグを再度有効にするには、Outlook デスクトップを再起動します。

**イベント ベースのハンドラーのデバッグ** ダイアログがポップアップ表示されないようにし、後続の Outlook セッションのデバッグを停止するには、関連付けられているレジストリ キーを削除するか、その値を次のように`0`設定します`HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger`。

## <a name="see-also"></a>関連項目

- [イベント ベースのアクティブ化のために Outlook アドインを構成する](autolaunch.md)
- [ランタイム ログを使用してアドインをデバッグする](../testing/runtime-logging.md#runtime-logging-on-windows)
