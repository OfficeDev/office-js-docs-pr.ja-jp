---
title: WebView2 用の開発者ツールを使用してアドインMicrosoft Edgeデバッグする
description: WebView2 の開発者ツールを使用してアドインMicrosoft Edgeデバッグします (Chromiumベース)。
ms.date: 11/09/2021
ms.localizationpriority: medium
ms.openlocfilehash: d2a8aa41720eebcd15d4cb625d90af1399b87dad
ms.sourcegitcommit: 3d37c42f5e465dac52d231d31717bdbb3bfa0e30
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/10/2021
ms.locfileid: "60923542"
---
# <a name="debug-add-ins-using-developer-tools-in-microsoft-edge-chromium-based"></a>アドインの開発者ツールを使用してアドインをデバッグする (Microsoft EdgeベースChromium)

この記事では、次の条件が満たされた場合に、アドインのクライアント側コード (JavaScript または TypeScript) をデバッグする方法を示します。

- IDE に組み込みのツールを使用してデバッグしたり、デバッグしたりすることはできません。または、アドインが IDE の外部で実行されている場合にのみ発生する問題が発生しています。
- コンピューターは、エッジ (Windowsベース) webview コントロール WebView2 をOfficeするバージョンとバージョンChromiumを組み合わせて使用しています。

> [!TIP]
> Visual Studio Code 内でのエッジ WebView2 (Chromium ベース) でのデバッグの詳細については、「Visual Studio Code および Microsoft Edge [WebView2](debug-desktop-using-edge-chromium.md)(Chromium ベース) を使用して Windows でアドインをデバッグする」を参照してください。

使用しているブラウザーを確認するには、「アドインで使用されるブラウザー [Office参照してください](../concepts/browsers-used-by-office-web-add-ins.md)。

> [!TIP]
> [!INCLUDE[Identify the webview through the add-in UI](../includes/identify-webview-in-ui.md)]

## <a name="debug-a-task-pane-add-in-using-microsoft-edge-chromium-based-developer-tools"></a>作業ウィンドウ アドインを (Microsoft EdgeベースChromium) 開発者ツールを使用してデバッグする

> [!NOTE]
> アドインに関数を実行するアドイン[](../design/add-in-commands.md)コマンドがある場合は、Microsoft Edge (Chromium ベース) の開発者ツールを起動できない非表示のブラウザー プロセスで関数が実行されます。そのため、この記事で説明する手法を使用して、関数内のコードをデバッグすることはできません。

1. [アドインを](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) サイドロードして実行します。
1. 次のいずれかのMicrosoft Edge (Chromiumベースの) 開発者ツールを実行します。

   - アドインの作業ウィンドウにフォーカスが設定され **、Ctrl + Shift + I キーを押してください**。
   - 作業ウィンドウを右クリックしてコンテキスト メニューを開き、[検査] を選択するか、パーソナリティ メニューを開いて [デバッガーの [接続]](../design/task-pane-add-ins.md#personality-menu)**を選択します**。

1. [ソース] **タブを開** きます。
1. 次の手順で、デバッグするファイルを開きます。

   1. ツールのトップ メニュー バーの右にある **[...]** ボタンを選択し、[検索] を **選択します**。
   1. 検索ボックスに、デバッグするファイルのコード行を入力します。 これは、他のファイルに含めそうになさいものである必要があります。
   1. [更新] ボタンを選択します。
   1. 検索結果で、行を選択して、検索結果の上のウィンドウでコード ファイルを開きます。

   :::image type="content" source="../images/open-file-in-edge-chromium-devtools.png" alt-text="A から D Chromiumラベル付き 4 つのパーツを含む[エッジ ツールソース] タブのスクリーンショット。":::

1. ブレークポイントを設定するには、コード ファイルの行番号を選択します。 コード ファイルの行に赤い点が表示されます。 右側のデバッガー ウィンドウで、ブレークポイントが [ブレークポイント] ドロップダウン **に** 登録されます。
1. 必要に応じてアドインの関数を実行して、ブレークポイントをトリガーします。

> [!TIP]
> ツールの使用の詳細については、「開発者ツール[の概要」Microsoft Edgeを参照してください](/microsoft-edge/devtools-guide-chromium/)。

## <a name="debug-a-dialog-in-an-add-in"></a>アドインでダイアログをデバッグする

アドインが Office Dialog API を使用している場合、ダイアログは作業ウィンドウ (存在する場合) とは別のプロセスで実行され、ツールは別のプロセスから開始する必要があります。 次の手順に従ってください。

1. アドインを実行します。
1. ダイアログを開き、フォーカスが設定されている必要があります。
1. 次のいずれかのMicrosoft Edge (Chromiumベースの) 開発者ツールを開きます。

   - Ctrl **+ Shift + I または** **F12 キーを押します**。
   - ダイアログを右クリックしてコンテキスト メニューを開き、[検査] を **選択します**。

1. 作業ウィンドウ内のコードと同じツールを使用します。 この[記事の前の「](#debug-a-task-pane-add-in-using-microsoft-edge-chromium-based-developer-tools)作業ウィンドウ アドインのデバッグ」をMicrosoft Edge (Chromiumベースの) 開発者ツールを使用して参照してください。
