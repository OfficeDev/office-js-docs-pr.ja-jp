---
title: テスト用に Office on the web で Office アドインをサイドロードする
description: サイドローディングOfficeして、Office on the webアドインをテストします。
ms.date: 08/02/2021
localization_priority: Normal
ms.openlocfilehash: b128bf26c185a51ef9b8235b4439739bc35253bd
ms.sourcegitcommit: e570fa8925204c6ca7c8aea59fbf07f73ef1a803
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/05/2021
ms.locfileid: "53773651"
---
# <a name="sideload-office-add-ins-in-office-on-the-web-for-testing"></a>テスト用に Office on the web で Office アドインをサイドロードする

アドインをサイドロードすると、アドインを最初にアドイン カタログに含めずにアドインをインストールできます。 これは、アドインの表示方法と機能を確認できるので、アドインをテストおよび開発する場合に便利です。

アドインをサイドロードすると、アドインのマニフェストはブラウザーのローカル ストレージに格納されます。そのため、ブラウザーのキャッシュをクリアするか、別のブラウザーに切り替える場合は、アドインを再度サイドロードする必要があります。

サイドローディングは、ホスト アプリケーションによって異なります (たとえば、Excel)。

> [!NOTE]
> この記事で説明するサイドローディングは、Excel、OneNote、PowerPoint、および Word でサポートされています。 Outlook アドインをサイドロードするには、「[テストのために Outlook アドインをサイドロードする](../outlook/sideload-outlook-add-ins-for-testing.md)」をご参照ください。

## <a name="sideload-an-office-add-in-in-office-on-the-web"></a>Office on the web で Office アドインをサイドロードする

このプロセスは、word、Excel、OneNote、PowerPointでのみ **サポート** されます。  他のホスト アプリケーションについては、次のセクションの手動サイドローディング手順を参照してください。 このサンプル プロジェクトでは、Yeoman ジェネレーターを使用して作成されたプロジェクトを、アドインに使用[Office想定しています](https://github.com/OfficeDev/generator-office)。

1. [ファイル[Office on the web] を開きます](https://office.live.com/)。 [作成 **] オプション** を使用して、Excel、OneNote、PowerPoint、**または** Word **で** ドキュメントを作成 **します**。  この新しいドキュメントで、リボン **で [共有** ] を選択し、[リンクのコピー] **を** 選択して URL をコピーします。

1. yo office プロジェクト ファイルのルート ディレクトリで、ファイルのpackage.js **開** きます。 このファイル **の構成** セクション内に、プロパティを作成 `"document"` します。 コピーした URL をプロパティの値として貼り付 `"document"` けます。 たとえば、次のようになります。

    ```json
      "config": {
        "document": "<YOUR URL>",
        ...
      }
    ```

    > [!TIP]
    > Yeoman ジェネレーターを使用しないアドインを作成する場合は、次の項目を既存の URL に追加して、ドキュメントの URL にクエリ パラメーターを追加できます。
    >
    > - など、開発サーバー ポート `&wdaddindevserverport=3000` 。
    > - マニフェスト ファイル名 ( `&wdaddinmanifestfile=manifest1.xml` など)。
    > - マニフェスト GUID など `&wdaddinmanifestguid=05c2e1c9-3e1d-406e-9a91-e9ac64854143` 。
    >
    > Yeoman ジェネレーターを使用している場合は、Yeoman ツールによってこの情報が自動的に追加されるので、この情報を追加する必要はありません。
    > ただし、どちらの場合も、localhost からのみマニフェストを読み込み可能です。

1. プロジェクトのルート ディレクトリから始まるコマンド ラインで、次のコマンドを実行します `npm run start:web` 。

1. このメソッドを初めて使用して、Web 上にアドインをサイドロードすると、開発者モードを有効にしてくださいというダイアログが表示されます。 [今すぐ開発者モードを **有効にする] のチェック ボックスをオンにして****、[OK] を選択します**。

1. 2 番目のダイアログ Office ボックスが表示されます。 [はい] を **選択する必要があります**。

1. アドインがインストールされています。 アドイン コマンドの場合は、リボンまたはコンテキスト メニューに表示されます。 作業ウィンドウ アドインの場合は、作業ウィンドウが表示されます。

## <a name="sideload-an-office-add-in-in-office-on-the-web-manually"></a>手動でOfficeアドインをサイドOffice on the webする

このメソッドはコマンド ラインを使用しないので、ホスト アプリケーション内のコマンド (コマンド など) を使用Excel。

1. [ファイル[Office on the web] を開きます](https://office.live.com/)。 [Word] 、または **[Excel]** で **ドキュメントを開** PowerPoint。 [アドイン **]** セクションのリボンの [挿入] タブ **で、[アドイン**] Office **を選択します**。

1. [アドイン **Office]** ダイアログで **、[MY ADD-INS]** タブを選択し、[自分のアドインの管理] を選択し、[マイ アップロード] をクリック **します**。

    ![右上Officeの [アドインの管理] というドロップダウンが表示された [Office アドイン] ダイアログボックスと、その下に [アップロード マイ アドイン] というオプションが表示されます。](../images/office-add-ins-my-account.png)

1. アドイン マニフェスト ファイルを **参照** して、**[アップロード]** を選択します。

    ![[参照]、[アップロード]、[キャンセル] のボタンがある [アドインのアップロード] ダイアログ。](../images/upload-add-in.png)

1. アドインがインストールされていることを確認します。たとえば、アドイン コマンドである場合は、リボンまたはコンテキスト メニューのいずれかに表示されます。作業ウィンドウ アドインである場合は、ウィンドウが表示されます。

> [!NOTE]
> 元の WebView Officeを使用Microsoft Edgeアドインをテストするには、追加の構成手順が必要です。 コマンド プロンプトWindows、次の行を実行します `npx office-addin-dev-settings appcontainer EdgeWebView --loopback --yes` 。 この機能は、Officeベースのエッジ WebView2 をChromium場合は必要ありません。 詳細については、「アドインで使用Office[ブラウザー」を参照してください](../concepts/browsers-used-by-office-web-add-ins.md)。

## <a name="sideload-an-office-add-in"></a>アドインをサイドOfficeする

1. アカウントにサインインMicrosoft 365します。

1. ツール バーの起動ツールで [App 起動ツール] を開き、[Excel]、[Word]、または **[PowerPoint]** を選択し、新しいドキュメントを作成します。 

1. 手順 3 から 6 は、前のセクション「**Office on the web で Office アドインをサイドロードする**」のものと同じです。

## <a name="sideload-an-add-in-when-using-visual-studio"></a>Visual Studio の使用時にアドインをサイドロードする

アドインの開発にVisual Studio場合、サイドロードするプロセスは、Web への手動サイドローディングに似ています。 アドインの開発に Visual Studio を使用している場合、サイドロードするプロセスは似ています。唯一の違いは、マニフェストの **SourceURL** 要素の値を更新して、アドインが展開されている完全な URL を含める必要がある点です。

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

ブラウザーのキャッシュをクリアすると、以前にサイドロードされたアドインを削除できます。 アドインのマニフェストに変更を加えた場合 (たとえば、アイコンのファイル名やアドイン コマンドのテキストを更新する) 場合は、ブラウザーのキャッシュをクリアしてから、更新されたマニフェストを使用してアドインを再読み込みする必要があります。 これにより、Office on the webマニフェストで説明されているとおりにアドインをレンダリングできます。

## <a name="see-also"></a>関連項目

- [iPad と Mac で Office アドインをサイドロードする](sideload-an-office-add-in-on-ipad-and-mac.md)
- [テスト用に Outlook アドインをサイドロードする](../outlook/sideload-outlook-add-ins-for-testing.md)
- [Office のキャッシュをクリアする](clear-cache.md)
