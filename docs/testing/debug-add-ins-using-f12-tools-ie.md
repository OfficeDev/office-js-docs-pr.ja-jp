---
title: Internet Explorer の開発者ツールを使用してアドインをデバッグする
description: アドインの開発者ツールを使用してアドインをデバッグInternet Explorer。
ms.date: 11/02/2021
ms.localizationpriority: medium
ms.openlocfilehash: bb7c328e6b1f839d5d711f81beceaf7519545fe3
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744666"
---
# <a name="debug-add-ins-using-developer-tools-in-internet-explorer"></a>アドインの開発者ツールを使用してアドインをデバッグInternet Explorer

この記事では、次の条件が満たされた場合に、アドインのクライアント側コード (JavaScript または TypeScript) をデバッグする方法を示します。

- IDE に組み込みのツールを使用してデバッグしたり、デバッグしたりすることはできません。または、アドインが IDE の外部で実行されている場合にのみ発生する問題が発生しています。
- コンピューターは、Webview コントロール Trident をWindowsバージョンOfficeバージョンと Internet Explorer組み合わせて使用しています。

コンピューターで使用されているブラウザーを確認するには、「アドインで使用Office[ブラウザー」を参照してください](../concepts/browsers-used-by-office-web-add-ins.md)。

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

> [!NOTE]
> Internet Explorer webview を使用する Office のバージョンをインストールする場合、または現在のバージョンに強制的に Internet Explorer を使用するには、「Internet Explorer [11 webview](#switch-to-the-internet-explorer-11-webview) に切り替える」を参照してください。

## <a name="debug-a-task-pane-add-in-using-the-f12-tools"></a>F12 ツールを使用して作業ウィンドウ アドインをデバッグする

Windows 10 11 には、F12 キーを押して起動されたので、"F12" という Web 開発ツールが含Internet Explorer。 F12 は、アドインが Webview コントロール Trident で実行されているときにアドインをデバッグするために使用Internet Explorerアプリケーションです。 アプリケーションは、以前のバージョンのアプリケーションでは使用Windows。

> [!NOTE]
> アドインに関数を実行するアドイン [](../design/add-in-commands.md) コマンドがある場合は、F12 ツールで検出または接続できない非表示のブラウザー プロセスで関数が実行されます。そのため、この記事で説明する手法を使用して、関数内のコードをデバッグすることはできません。

次の手順は、アドインをデバッグするための手順です。 F12 ツール自体をテストする場合は、「F12 ツールをテストするアドインの例 [」を参照してください](#example-add-in-to-test-the-f12-tools)。

1. [アドインを](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) サイドロードして実行します。
1. バージョンのバージョンに対応する F12 開発ツールを起動Office。

   - 32 ビット版の Office の場合は、C:\Windows\System32\F12\IEChooser.exe を使用します
   - 64 ビット版の Office の場合は、C:\Windows\SysWOW64\F12\IEChooser.exe を使用します

   IEChooser が開き、[デバッグするターゲットの選択 **] という名前のウィンドウが開きます**。 アドインは、アドインのホーム ページのファイル名で指定されたウィンドウに表示されます。 次のスクリーンショットでは、 です `Home.html`。 ユーザーまたは Trident で実行Internet Explorerプロセスだけが表示されます。 このツールは、他のブラウザーや Web ビューで実行されているプロセス (ブラウザーや web ビューなど) に接続Microsoft Edge。

    :::image type="content" source="../images/choose-target-to-debug.png" alt-text="IEChooser 画面で、複数のInternet Explorerトライデント プロセスが一覧表示されます。1 つは名前が Home.html。":::

1. アドインのプロセスを選択します。つまり、ホーム ページ ファイル名です。 このアクションでは、F12 ツールをプロセスに接続し、メインの F12 ユーザー インターフェイスを開きます。
1. **[デバッガー]** タブを開きます。
1. タブの左上にあるデバッガー ツール リボンの下に、小さなフォルダー アイコンがあります。 アドイン内のファイルのドロップダウン リストを開く場合は、これを選択します。 次に例を示します。

    :::image type="content" source="../images/f12-file-dropdown.png" alt-text="フォルダー ドロップダウンが開いているデバッガー タブの左上隅とファイルの一覧のスクリーンショット。":::

1. デバッグするファイルを選択し、[デバッガー] タブ **のスクリプト (** 左) ウィンドウで **開** きます。ファイルの名前を変更するトランスピラー、バンドル、またはミニファイアを使用している場合は、元のソース ファイル名ではなく、実際に読み込まれる最終的な名前が含まれます。

1. ブレークポイントを設定する行までスクロールし、行番号の左側の余白をクリックします。 行の左側に赤い点が表示され、右下ウィンドウの [ **ブレークポイント** ] タブに対応する行が表示されます。 次のスクリーンショットは、その一例を示しています。

    :::image type="content" source="../images/debugger-home-js-02.png" alt-text="ファイル内のブレークポイントを持つデバッガーhome.jsします。":::

1. 必要に応じてアドインの関数を実行して、ブレークポイントをトリガーします。 ブレークポイントがヒットすると、ブレークポイントの赤い点に右向きの矢印が表示されます。 次のスクリーンショットは、その一例を示しています。

    :::image type="content" source="../images/debugger-home-js-01.png" alt-text="トリガーされたブレークポイントからの結果を含むデバッガー。":::

> [!TIP]
> F12 ツールの使用の詳細については、「デバッガーで [JavaScript を実行する検査」を参照してください](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))。

### <a name="example-add-in-to-test-the-f12-tools"></a>F12 ツールをテストするアドインの例

次の例では、AppSource から Word と無料のアドインを使用します。

1. Word を起動し、空白の文書を選択します。
1. [挿入 **] タブ** の [アドイン] グループで、[マイ アドイン]  **を** 選択して [Office アドイン] ダイアログを開き、[STORE] タブ **を選択** します。
1. **QR4Office アドイン** を選択します。 作業ウィンドウで開きます。
1. 前のセクションで説明したように、バージョンのバージョンに対応する F12 Officeツールを起動します。
1. [F12] ウィンドウで、[ファイル] **をHome.html**。
1. [デバッガー **] タブ** で、前のセクション **Home.js** ファイル を開きます。
1. 310 行と 312 行にブレークポイントを設定します。
1. アドインで、[挿入] ボタン **を選択** します。 1 つ以上のブレークポイントがヒットします。

## <a name="debug-a-dialog-in-an-add-in"></a>アドインでダイアログをデバッグする

アドインで Office ダイアログ API を使用する場合、ダイアログは作業ウィンドウとは別のプロセス (存在する場合) で実行され、ツールは、そのプロセスに添付する必要があります。 次の手順に従ってください。

1. アドインとツールを実行します。 
1. ダイアログを開き、ツールの [ **更新** ] ボタンを選択します。 ダイアログ プロセスが表示されます。 その名前は、ダイアログで開いているファイルのファイル名です。
1. プロセスを選択して開き、F12 ツールを使用して作業ウィンドウ アドインをデバッグするセクションの説明に従 [ってデバッグします](#debug-a-task-pane-add-in-using-the-f12-tools)。

## <a name="switch-to-the-internet-explorer-11-webview"></a>11 webview Internet Explorerに切り替える

Webview に切り替える方法は 2 Internet Explorerがあります。 コマンド プロンプトで簡単なコマンドを実行するか、既定でコマンド を使用OfficeバージョンInternet Explorerインストールできます。 最初の方法をお勧めします。 ただし、次のシナリオでは 2 つ目を使用する必要があります。

- プロジェクトは、プロジェクトと IIS Visual Studio開発されました。 この機能は、node.jsに基づいて行う必要があります。
- テストで絶対に堅牢になる必要があります。
- 何らかの理由でコマンド ライン ツールが機能しない場合。

### <a name="switch-via-the-command-line"></a>コマンド ライン経由で切り替える

[!INCLUDE [Steps to switch browsers with the command line tool](../includes/use-legacy-edge-or-ie.md)]

### <a name="install-a-version-of-office-that-uses-internet-explorer"></a>アプリケーションを使用するバージョンOfficeインストールInternet Explorer

[!INCLUDE [Steps to install Office that uses Edge Legacy or Internet Explorer](../includes/install-office-that-uses-legacy-edge-or-ie.md)]

## <a name="see-also"></a>関連項目

- [デバッガーを使用して実行中の JavaScript を検査する](/previous-versions/windows/internet-explorer/ie-developer/samples/dn255007(v=vs.85))
- [F12 開発者ツールの使用](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85))
