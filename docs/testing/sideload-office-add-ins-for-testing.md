---
title: Office アドインをOffice on the webにサイドロードする
description: サイドローディングを使用して、Office on the webで Office アドインをテストします。
ms.date: 09/02/2022
ms.localizationpriority: medium
ms.openlocfilehash: 128e3537ac0ece5b7574dfec6d9d5c67b8d95a7b
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810385"
---
# <a name="sideload-office-add-ins-to-office-on-the-web"></a>Office アドインをOffice on the webにサイドロードする

アドインをサイドロードする場合は、まずアドイン カタログにアドインを配置せずにアドインをインストールできます。 これは、アドインの表示方法と機能を確認できるため、アドインのテストと開発に役立ちます。

Web でアドインをサイドロードする場合、アドインのマニフェストはブラウザーのローカル ストレージに格納されるため、ブラウザーのキャッシュをクリアするか、別のブラウザーに切り替える場合は、もう一度アドインをサイドロードする必要があります。

Web 上でアドインをサイドロードする手順は、次の要因によって異なります。

- ホスト アプリケーション (Excel、Word、Outlook など)
- アドイン プロジェクトを作成したツール (たとえば、Visual Studio、Office アドイン用 Yeoman ジェネレーター、またはどちらも作成しません)
- Microsoft アカウントを使用してOffice on the webにサイドローディングする場合も、Microsoft 365 テナント内のアカウントを使用する場合も

次の一覧で、シナリオに一致するセクションまたは記事に移動します。 一覧の最初のシナリオは Outlook アドインに適用されることに注意してください。残りのシナリオは、Outlook 以外のアドインに適用されます。

- Outlook アドインをサイドロードする場合は、「 [テスト用の Outlook アドインをサイドロードする](../outlook/sideload-outlook-add-ins-for-testing.md)」の記事を参照してください。
- [Office アドイン用 Yeoman ジェネレーター](../develop/yeoman-generator-overview.md)を使用してアドインを作成した場合は、「[Yeoman で作成されたアドインをOffice on the webにサイドロードする](#sideload-a-yeoman-created-add-in-to-office-on-the-web)」を参照してください。
- Visual Studio を使用してアドインを作成した場合は、「Visual Studio を [使用するときに Web でアドインをサイドロード](#sideload-an-add-in-on-the-web-when-using-visual-studio)する」を参照してください。
- その他のすべてのケースについては、次のセクションのいずれかを参照してください。

  - Microsoft アカウントを使用してOffice on the webにサイドローディングする場合は、「[アドインを手動でOffice on the webにサイドロードする](#manually-sideload-an-add-in-to-office-on-the-web)」を参照してください。
  - Microsoft 365 テナント内のアカウントを使用してOffice on the webにサイドローディングする場合は、「[Microsoft 365 へのアドインのサイドロード](#sideload-an-add-in-to-microsoft-365)」を参照してください。

## <a name="sideload-a-yeoman-created-add-in-to-office-on-the-web"></a>Yeoman で作成されたアドインをOffice on the webにサイドロードする

このプロセスは、 **Excel**、 **OneNote**、 **PowerPoint**、 **Word** でのみサポートされています。 このプロジェクト例では、 [Office アドイン用 Yeoman ジェネレーター](../develop/yeoman-generator-overview.md)で作成されたプロジェクトを使用していることを前提としています。

1. [Office on the web](https://office.live.com/) または OneDrive を開きます。 **[作成**] オプションを使用して、**Excel**、**OneNote**、**PowerPoint**、または **Word** でドキュメントを作成します。 この新しいドキュメントで、[ **共有**] を選択し、[ **リンクのコピー**] を選択し、URL をコピーします。

1. プロジェクトのルート ディレクトリから始まるコマンド ラインで、次のコマンドを実行します。 "{url}" をコピーした URL に置き換えます。

    [!INCLUDE [npm start:web command syntax](../includes/start-web-sideload-instructions.md)]

1. このメソッドを初めて使用して Web 上のアドインをサイドロードすると、開発者モードを有効にするように求めるダイアログが表示されます。 [ **今すぐ開発者モードを有効にする]** のチェック ボックスをオンにし、[ **OK] を選択します**。

1. コンピューターから Office アドイン マニフェストを登録するかどうかを確認する 2 つ目のダイアログ ボックスが表示されます。 [**はい**] を選択します。

1. アドインがインストールされています。 アドイン コマンドがある場合は、リボンまたはコンテキスト メニューに表示されます。 アドイン コマンドのない作業ウィンドウ アドインの場合は、作業ウィンドウが表示されます。

## <a name="sideload-an-add-in-on-the-web-when-using-visual-studio"></a>Visual Studio を使用するときにアドインを Web にサイドロードする

Visual Studio を使用してアドインを開発している場合は、 **F5** キーを押して *デスクトップ* Office で Office ドキュメントを開き、空白のドキュメントを作成し、アドインをサイドロードします。 *Office on the web* にサイドロードする場合、サイドロードするプロセスは、Web への手動サイドローディングに似ています。 唯一の違いは、マニフェスト内の **SourceURL** 要素と場合によっては他の要素の値を更新して、アドインがデプロイされる完全な URL を含める必要があるということです。

1. Visual Studio で、[**プロパティ ウィンドウ****の表示** > ] を選択します。

1. [**ソリューション エクスプローラー**] で Web プロジェクトを選択します。 プロジェクトのプロパティが [ **プロパティ** ] ウィンドウに表示されます。

1. [プロパティ] ウィンドウで、[**SSL URL**] をコピーします。

1. アドイン プロジェクトで、マニフェスト XML ファイルを開きます。 ソース XML を編集していることを確認します。 一部のプロジェクトの種類では、次の手順では機能しない XML のビジュアル ビューが Visual Studio によって開きます。

1. **~remoteAppUrl/** のすべてのインスタンスを検索し、先ほどコピーした SSL URL と置き換えます。 プロジェクトの種類に応じていくつかの置換が表示され、新しい URL は のように `https://localhost:44300/Home.html`表示されます。

1. XML ファイルを **保存します**。

1. **ソリューション エクスプローラー** で、Web プロジェクトのコンテキスト メニュー (右クリックなど) を開き、[**デバッグ** > **] [新しいインスタンスの開始**] の順に選択します。 これにより、Office を起動せずに Web プロジェクトが実行されます。

1. Office on the webから、「手動でアドインをOffice on the webにサイドロードする」で説明されている手順を使用して[アドインをサイドロードします](#manually-sideload-an-add-in-to-office-on-the-web)。

## <a name="manually-sideload-an-add-in-to-office-on-the-web"></a>アドインを手動でサイドロードしてOffice on the web

このメソッドはコマンド ラインを使用せず、ホスト アプリケーション (Excel など) 内でのみコマンドを使用して実行できます。

1. [Office on the web](https://office.com/)を開きます。 **Excel**、**OneNote**、**PowerPoint**、または **Word** でドキュメントを開きます。 

1. [ **挿入** ] タブの [アドイン] セクション **で** 、[ **Office アドイン**] を選択します。

1. [ **Office アドイン** ] ダイアログで、[ **MY ADD-INS** ] タブを選択し、[ **マイ アドインの管理**] を選択し、[ **マイ アドインのアップロード**] を選択します。

    ![右上の [アドインの管理] というドロップダウンと、[マイ アドインのアップロード] オプションが表示されたドロップダウンが表示された [Office アドイン] ダイアログ。](../images/office-add-ins-my-account.png)

1. アドイン マニフェスト ファイルを **参照** して、**[アップロード]** を選択します。

    ![[参照]、[アップロード]、[キャンセル] のボタンがある [アドインのアップロード] ダイアログ。](../images/upload-add-in.png)

1. アドインがインストールされていることを確認します。 たとえば、アドイン コマンドがある場合は、リボンまたはコンテキスト メニューに表示されます。 アドイン コマンドがない作業ウィンドウ アドインの場合は、作業ウィンドウが表示されます。

> [!NOTE]
> 元の WebView (EdgeHTML) を使用して Microsoft Edge で Office アドインをテストするには、追加の構成手順が必要です。 Windows コマンド プロンプトで、次の行を実行します。 `npx office-addin-dev-settings appcontainer EdgeWebView --loopback --yes` これは、Office が Chromium ベースの Edge WebView2 を使用している場合は必要ありません。 詳細については、「[Office アドインによって使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md)」を参照してください。

[!INCLUDE[Office settings tool not supported on Mac](../includes/tool-nonsupport-mac-note.md)]

## <a name="sideload-an-add-in-to-microsoft-365"></a>アドインを Microsoft 365 にサイドロードする

1. Microsoft 365 アカウントにサインインします。

1. ツール バーの左端にあるアプリ起動ツールを開き、[ **Excel**]、[ **OneNote**]、[ **PowerPoint**]、または **[Word**] を選択し、新しいドキュメントを作成します。

1. [ **挿入** ] タブ **で、[アドイン** ] ボタンを選択します。

1. 「[手動でアドインをOffice on the webにサイドロードする](#manually-sideload-an-add-in-to-office-on-the-web)」セクションの手順 3 から 5 に従います。

## <a name="remove-a-sideloaded-add-in"></a>サイドロードされたアドインを削除する

Office on the webにサイドロードされたアドインを削除するには、ブラウザーのキャッシュをクリアするだけです。 アドインのマニフェストを変更する場合 (たとえば、アイコンのファイル名の更新やアドイン コマンドのテキストの更新)、ブラウザーのキャッシュをクリアしてから、更新されたマニフェストを使用してアドインを再サイドロードする必要がある場合があります。 これにより、更新されたマニフェストで説明されているように、Office on the webアドインをレンダリングできます。

## <a name="see-also"></a>関連項目

- [Office アドインと Mac で Office アドインをサイドロードする](sideload-an-office-add-in-on-mac.md)
- [iPad と Office アドイン で Office アドインをサイドロードする](sideload-an-office-add-in-on-ipad.md)
- [テスト用に Outlook アドインをサイドロードする](../outlook/sideload-outlook-add-ins-for-testing.md)
- [Office のキャッシュをクリアする](clear-cache.md)
