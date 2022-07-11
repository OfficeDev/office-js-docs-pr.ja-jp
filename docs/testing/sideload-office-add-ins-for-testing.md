---
title: テスト用に Office on the web で Office アドインをサイドロードする
description: サイドローディングを使用して、Office on the webで Office アドインをテストします。
ms.date: 02/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 32d80a10ccddab93fc8d41151be6a2842d3732cb
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/11/2022
ms.locfileid: "66713029"
---
# <a name="sideload-office-add-ins-in-office-on-the-web-for-testing"></a>テスト用に Office on the web で Office アドインをサイドロードする

アドインをサイドロードすると、アドイン カタログに最初に配置せずにアドインをインストールできます。 これは、アドインの表示方法と機能を確認できるため、アドインをテストして開発するときに便利です。

アドインをサイドロードすると、アドインのマニフェストはブラウザーのローカル ストレージに格納されるため、ブラウザーのキャッシュをクリアするか、別のブラウザーに切り替える場合は、アドインをもう一度サイドロードする必要があります。

サイドローディングは、ホスト アプリケーション (Excel など) によって異なります。

> [!NOTE]
> この記事で説明するようにサイドローディングは、Excel、OneNote、PowerPoint、Word でサポートされています。 Outlook アドインをサイドロードするには、「[テストのために Outlook アドインをサイドロードする](../outlook/sideload-outlook-add-ins-for-testing.md)」をご参照ください。

## <a name="sideload-an-office-add-in-in-office-on-the-web"></a>Office on the web で Office アドインをサイドロードする

このプロセスは、 **Excel**、 **OneNote**、 **PowerPoint**、 **Word** でのみサポートされています。 その他のホスト アプリケーションについては、次のセクションの手動サイドローディング手順を参照してください。 このプロジェクト例では、 [Office アドイン用の Yeoman ジェネレーター](../develop/yeoman-generator-overview.md)で作成されたプロジェクトを使用していることを前提としています。

1. [Office on the web](https://office.live.com/)を開きます。 **[作成**] オプションを使用して、**Excel**、**OneNote**、**PowerPoint**、または Word でドキュメントを作成 **します**。 この新しいドキュメントで、リボンで **[共有** ] を選択し、[ **リンクのコピー**] を選択して URL をコピーします。

1. office プロジェクト ファイルのルート ディレクトリで、 **package.json** ファイルを開きます。 このファイルの **構成** セクション内で、プロパティを `"document"` 作成します。 コピーした URL をプロパティの値 `"document"` として貼り付けます。 たとえば、次のようになります。

    ```json
      "config": {
        "document": "<YOUR URL>",
        ...
      }
    ```

    > [!TIP]
    > Yeoman ジェネレーターを使用しないアドインを作成する場合は、既存の URL に次を追加して、ドキュメントの URL にクエリ パラメーターを追加できます。
    >
    > - 開発サーバー のポート (例: `&wdaddindevserverport=3000`.
    > - マニフェスト ファイル名 (例: `&wdaddinmanifestfile=manifest1.xml`.
    > - マニフェスト GUID (例: `&wdaddinmanifestguid=05c2e1c9-3e1d-406e-9a91-e9ac64854143`.
    >
    > Yeoman ジェネレーターを使用している場合は、Yeoman ツールによってこの情報が自動的に追加されるため、この情報を追加する必要はありません。
    > ただし、どちらの場合も、localhost からのみマニフェストを読み込むことができます。

1. プロジェクトのルート ディレクトリからコマンド ラインで、次のコマンドを実行します。 "{url}" は、OneDrive またはアクセス許可を持つ SharePoint ライブラリ上の Office ドキュメントの URL に置き換えます。

    [!INCLUDE [npm start:web command syntax](../includes/start-web-sideload-instructions.md)]

1. このメソッドを初めて使用してアドインを Web にサイドロードすると、開発者モードを有効にするように求めるダイアログが表示されます。 [ **今すぐ開発者モードを有効にする]** チェック ボックスをオンにし、[OK] を選択 **します**。

1. 2 つ目のダイアログ ボックスが表示され、コンピューターから Office アドイン マニフェストを登録するかどうかを確認するメッセージが表示されます。 **[はい**] を選択する必要があります。

1. アドインがインストールされています。 アドイン コマンドの場合は、リボンまたはコンテキスト メニューに表示されます。 作業ウィンドウ アドインの場合は、作業ウィンドウが表示されます。

## <a name="sideload-an-office-add-in-in-office-on-the-web-manually"></a>Office アドインを手動でサイドロードOffice on the web

このメソッドはコマンド ラインを使用せず、ホスト アプリケーション (Excel など) 内でのみコマンドを使用して実行できます。

1. [Office on the web](https://office.com/)を開きます。 **Excel**、**OneNote**、**PowerPoint**、または **Word** でドキュメントを開きます。 [アドイン] セクションのリボンの [ **挿入** ] タブ **で** 、[ **Office アドイン**] を選択します。

1. **[Office アドイン**] ダイアログで、[**MY ADD-INS**] タブを選択し、[**マイ アドインの管理**] を選択して、[**マイ アドインのアップロード**] を選択します。

    ![[Office アドイン] ダイアログで、右上にドロップダウンが [アドインの管理] と表示され、その下に [自分のアドインのアップロード] オプションが表示されます。](../images/office-add-ins-my-account.png)

1. アドイン マニフェスト ファイルを **参照** して、**[アップロード]** を選択します。

    ![[参照]、[アップロード]、[キャンセル] のボタンがある [アドインのアップロード] ダイアログ。](../images/upload-add-in.png)

1. アドインがインストールされていることを確認します。たとえば、アドイン コマンドである場合は、リボンまたはコンテキスト メニューのいずれかに表示されます。作業ウィンドウ アドインである場合は、ウィンドウが表示されます。

> [!NOTE]
> 元の WebView (EdgeHTML) を使用して Microsoft Edge で Office アドインをテストするには、追加の構成手順が必要です。 Windows コマンド プロンプトで、次の行 `npx office-addin-dev-settings appcontainer EdgeWebView --loopback --yes`を実行します。 これは、Office が Chromium ベースの Edge WebView2 を使用している場合は必要ありません。 詳細については、「[Office アドインによって使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md)」を参照してください。

[!INCLUDE[Office settings tool not supported on Mac](../includes/tool-nonsupport-mac-note.md)]

## <a name="sideload-an-office-add-in"></a>Office アドインをサイドロードする

1. Microsoft 365 アカウントにサインインします。

1. ツール バーの左端にあるアプリ 起動ツールを開き、 **Excel**、 **PowerPoint**、または **Word** を選択して、新しいドキュメントを作成します。

1. 手順 3 から 6 は、前のセクション「**Office on the web で Office アドインをサイドロードする**」のものと同じです。

## <a name="sideload-an-add-in-when-using-visual-studio"></a>Visual Studio の使用時にアドインをサイドロードする

Visual Studio を使用してアドインを開発する場合、サイドロードするプロセスは Web への手動サイドローディングに似ています。 アドインの開発に Visual Studio を使用している場合、サイドロードするプロセスは似ています。唯一の違いは、マニフェストの **SourceURL** 要素の値を更新して、アドインが展開されている完全な URL を含める必要がある点です。

> [!NOTE]
> アドインは Visual Studio から Office on the web にサイドロードできますが、Visual Studio からはデバッグできません。 デバッグするには、ブラウザー デバッグ ツールを使用する必要があります。 詳細については、「[Office on the web でアドインをデバッグする](debug-add-ins-in-office-online.md)」を参照してください。

1. Visual Studio で、[**表示**]  >  [**プロパティ ウィンドウ**] の順に選択して [**プロパティ**] ウィンドウを表示させます。
1. [**ソリューション エクスプローラー**] で Web プロジェクトを選択します。 プロジェクトのプロパティが [**プロパティ**] ウィンドウに表示されます。
1. [プロパティ] ウィンドウで、[**SSL URL**] をコピーします。
1. アドイン プロジェクトで、マニフェスト XML ファイルを開きます。 編集しているのがソース XML であることを確認します。 一部の種類のプロジェクトでは、Visual Studio は XML のビジュアル ビューを開きますが、これは次の手順で使用できません。
1. **~remoteAppUrl/** のすべてのインスタンスを検索し、先ほどコピーした SSL URL と置き換えます。 プロジェクトの種類に応じていくつかの置換が表示され、新しい URL の表示は `https://localhost:44300/Home.html` に似たものになりま。
1. XML ファイルを保存します。
1. Web プロジェクトを右クリックして、[**デバッグ**]  >  [**新しいインスタンスを開始**] の順に選択します。 これにより、Office を起動することなく Web プロジェクトが実行されます。
1. 前述の「[Office on the web で Office アドインをサイドロードする](#sideload-an-office-add-in-in-office-on-the-web)」で説明した手順を使用して、Office on the web からアドインをサイドロードします。

## <a name="remove-a-sideloaded-add-in"></a>サイドロードされたアドインを削除する

ブラウザーのキャッシュをクリアすることで、以前にサイドロードされたアドインを削除できます。 アドインのマニフェストに変更を加えた場合 (たとえば、アイコンのファイル名やアドイン コマンドのテキストを更新するなど)、ブラウザーのキャッシュをクリアしてから、更新されたマニフェストを使用してアドインを再サイドロードする必要がある場合があります。 これにより、更新されたマニフェストで説明されているように、アドインをレンダリングするOffice on the webが可能になります。

## <a name="see-also"></a>関連項目

- [Mac で Office アドインをサイドロードする](sideload-an-office-add-in-on-mac.md)
- [iPad で Office アドインをサイドロードする](sideload-an-office-add-in-on-ipad.md)
- [テスト用に Outlook アドインをサイドロードする](../outlook/sideload-outlook-add-ins-for-testing.md)
- [Office のキャッシュをクリアする](clear-cache.md)
