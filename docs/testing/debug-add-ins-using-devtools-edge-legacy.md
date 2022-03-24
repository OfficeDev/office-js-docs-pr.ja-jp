---
title: 開発者向けツールを使用してアドインをデバッグMicrosoft Edge 従来版
description: アドインの開発者ツールを使用してアドインをデバッグMicrosoft Edge 従来版。
ms.date: 11/02/2021
ms.localizationpriority: medium
ms.openlocfilehash: 62f27e2ee266e3b6a92d090e8008b74bac4a9663
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744681"
---
# <a name="debug-add-ins-using-developer-tools-in-microsoft-edge-legacy"></a>アドインの開発者ツールを使用してアドインをデバッグMicrosoft Edge 従来版

この記事では、次の条件が満たされた場合に、アドインのクライアント側コード (JavaScript または TypeScript) をデバッグする方法を示します。

- IDE に組み込みのツールを使用してデバッグしたり、デバッグしたりすることはできません。または、アドインが IDE の外部で実行されている場合にのみ発生する問題が発生しています。
- コンピューターは、元の Edge webview コントロール EdgeHTML をWindowsバージョンとOfficeバージョンの組み合わせを使用しています。

> [!TIP]
> エッジ レガシを使用したデバッグの詳細Visual Studio Code、「Microsoft Office アドイン デバッガー拡張機能[」を参照](debug-with-vs-extension.md)Visual Studio Code。

使用しているブラウザーを確認するには、「アドインで使用されるブラウザー [Office参照してください](../concepts/browsers-used-by-office-web-add-ins.md)。 

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

> [!NOTE]
> エッジ レガシ Web ビューを使用する Office のバージョンをインストールする場合、または現在のバージョンの Office でエッジ レガシを使用する場合は、「Switch to the Edge [Legacy webview](#switch-to-the-edge-legacy-webview)」を参照してください。

## <a name="debug-a-task-pane-add-in-using-microsoft-edge-devtools-preview"></a>DevTools Preview を使用して作業ウィンドウ アドインMicrosoft Edgeデバッグする

1. [DevTools プレビュー Microsoft Edgeインストールします](https://www.microsoft.com/p/microsoft-edge-devtools-preview/9mzbfrmz0mnj?activetab=pivot%3Aoverviewtab)。 ("Preview" という単語は、歴史的な理由から名前に含されます。 より新しいバージョンは含めなかった。

   > [!NOTE]
   > アドインに関数を実行するアドイン [](../design/add-in-commands.md) コマンドがある場合、この関数は、Microsoft Edge DevTools が検出または接続できない非表示のブラウザー プロセスで実行されます。そのため、この記事で説明する手法を使用して、関数内のコードをデバッグすることはできません。

1. [アドインを](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) サイドロードして実行します。
1. Microsoft Edge DevTools を実行します。
1. ツールで、**[ローカル]** タブを開きます。アドインの名前が一覧表示されます。 (EdgeHTML で実行されているプロセスだけがタブに表示されます。このツールは、Microsoft Edge (WebView2) や Internet Explorer (Trident) など、他のブラウザーまたは Web ビューで実行されているプロセスに接続できません。

   :::image type="content" source="../images/edge-devtools-with-add-in-process.png" alt-text="従来のエッジ デバッグという名前のプロセスを示す Edge DevTools のスクリーンショット。":::

1. アドイン名を選択して、ツールで開きます。
1. **[デバッガー]** タブを開きます。
1. 次の手順で、デバッグするファイルを開きます。

   1. デバッガー のタスク バーで、[ファイルに検索 **を表示する] を選択します**。 これにより、検索ウィンドウが開きます。
   1. 検索ボックスに、デバッグするファイルのコード行を入力します。 これは、他のファイルに含めそうになさいものである必要があります。
   1. [更新] ボタンを選択します。
   1. 検索結果で、行を選択して、検索結果の上のウィンドウでコード ファイルを開きます。

   :::image type="content" source="../images/open-file-in-edge-devtools.png" alt-text="A から D のラベルが付いた 4 つのパーツを含む Edge DevTools デバッグ タブのスクリーンショット。":::

1. ブレークポイントを設定するには、コード ファイルの行を選択します。 ブレークポイントが [呼び出し履歴] ( **右下)** ウィンドウに登録されます。 コード ファイル内の行に赤い点が表示される場合がありますが、これは確実には表示されません。
1. 必要に応じてアドインの関数を実行して、ブレークポイントをトリガーします。

> [!TIP]
> ツールの使用の詳細については、「Microsoft Edge [(EdgeHTML) Developer Tools」を参照してください](/archive/microsoft-edge/legacy/developer/devtools-guide/)。

## <a name="debug-a-dialog-in-an-add-in"></a>アドインでダイアログをデバッグする

アドインで Office ダイアログ API を使用する場合、ダイアログは作業ウィンドウとは別のプロセス (存在する場合) で実行され、ツールは、そのプロセスに添付する必要があります。 次の手順に従ってください。

1. アドインとツールを実行します。
1. ダイアログを開き、ツールの [ **更新** ] ボタンを選択します。 ダイアログ プロセスが表示されます。 その名前は、ダイアログで `<title>` 開いている HTML ファイル内の要素から取得されます。
1. 「[DevTools Preview](#debug-a-task-pane-add-in-using-microsoft-edge-devtools-preview) を使用して作業ウィンドウ アドインをデバッグする」セクションの説明に従って、プロセスを選択してデバッグMicrosoft Edgeします。

   :::image type="content" source="../images/edge-devtools-with-add-in-and-dialog-processes.png" alt-text="My Dialog という名前のプロセスを示す Edge DevTools のスクリーンショット。":::

## <a name="switch-to-the-edge-legacy-webview"></a>エッジ レガシ Web ビューに切り替える

エッジ レガシ Web ビューを切り替える方法は 2 通りあります。 コマンド プロンプトで簡単なコマンドを実行するか、既定でエッジ レガシを使用Officeバージョンをインストールできます。 最初の方法をお勧めします。 ただし、次のシナリオでは 2 つ目を使用する必要があります。

- プロジェクトは、プロジェクトと IIS Visual Studio開発されました。 この機能は、node.jsに基づいて行う必要があります。
- テストで絶対に堅牢になる必要があります。
- 何らかの理由でコマンド ライン ツールが機能しない場合。

### <a name="switch-via-the-command-line"></a>コマンド ライン経由で切り替える

[!INCLUDE [Steps to switch browsers with the command line tool](../includes/use-legacy-edge-or-ie.md)]

### <a name="install-a-version-of-office-that-uses-edge-legacy"></a>エッジ レガシを使用Officeバージョンをインストールする

[!INCLUDE [Steps to install Office that uses Edge Legacy or Internet Explorer](../includes/install-office-that-uses-legacy-edge-or-ie.md)]
