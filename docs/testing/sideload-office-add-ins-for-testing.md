---
title: テスト用に Office on the web で Office アドインをサイドロードする
description: サイドロードOffice Web 上のOfficeアドインをテストします。
ms.date: 02/11/2021
localization_priority: Normal
ms.openlocfilehash: f81fbc163135be5a616e7b44e604cb842da9870b
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/14/2021
ms.locfileid: "50238065"
---
# <a name="sideload-office-add-ins-in-office-on-the-web-for-testing"></a>テスト用に Office on the web で Office アドインをサイドロードする

アドインをサイドロードすると、最初にアドイン カタログにアドインを置かずにアドインをインストールできます。 これは、アドインの表示方法と機能を確認できるアドインをテストおよび開発するときに役立ちます。

アドインをサイドロードすると、アドインのマニフェストはブラウザーのローカル ストレージに格納されます。そのため、ブラウザーのキャッシュをクリアするか、別のブラウザーに切り替える場合は、アドインを再びサイドロードする必要があります。

サイドロードは、ホスト アプリケーション (Excel など) によって異なります。

> [!NOTE]
> この記事で説明するようにサイドロードは、Excel、OneNote、PowerPoint、および Word でサポートされています。Outlook アドインをサイドロードするには、「テスト用に Outlook アドイン [をサイドロードする」を参照してください](../outlook/sideload-outlook-add-ins-for-testing.md)。

## <a name="sideload-an-office-add-in-in-office-on-the-web"></a>Office on the web で Office アドインをサイドロードする

このプロセスは、Excel、OneNote、PowerPoint、**および** **Word でのみサポート** されています。   その他のホスト アプリケーションについては、次のセクションの手動サイドロード手順を参照してください。 このサンプル プロジェクトでは、アドイン用の Yeoman ジェネレーターで作成された [プロジェクトOffice想定しています](https://github.com/OfficeDev/generator-office)。

1. Web [Officeを開きます](https://office.live.com/)。 [作成 **] オプション** を使用して、Excel、OneNote、PowerPoint、または Word **でドキュメント** を作成 **します**。  この新しいドキュメントで、リボンで **[共有**]を選択し、[リンクのコピー] を選択して URL をコピーします。

2. yo office プロジェクト ファイルのルート ディレクトリで、ファイルのpackage.js **開** きます。 このファイル **のスクリプト** セクション内にプロパティを作成 `"document"` します。 コピーした URL をプロパティの値として貼り付 `"document"` けます。 たとえば、次のようになります。

    ```json
      "scripts": {
        "document": "<YOUR URL>",
        ...
      }
    ```

    > [!TIP]
    > Yeoman ジェネレーターを使用しないアドインを作成する場合は、既存の URL に次を追加して、ドキュメントの URL にクエリ パラメーターを追加できます。

    - 開発サーバーのポート。次に例を示します `&wdaddindevserverport=3000` 。
    - マニフェスト ファイル名。次に例を示します `&wdaddinmanifestfile=manifest1.xml` 。
    - マニフェスト GUID。次に例を示します `&wdaddinmanifestguid=05c2e1c9-3e1d-406e-9a91-e9ac64854143` 。

    > Yeoman ジェネレーターを使用している場合、Yeoman ツールによってこの情報が自動的に追加されるので、この情報を追加する必要はありません。
    > ただし、どちらの場合も、localhost からしかマニフェストを読み込めない点に注意してください。

3. プロジェクトのルート ディレクトリから始まるコマンド ラインで、次のコマンドを実行します `npm run start:web` 。

4. このメソッドを初めて使用して Web 上にアドインをサイドロードすると、開発者モードを有効にしてくださいというダイアログが表示されます。 [開発者モードを今すぐ有効 **にする] チェック ボックスをオンにして****、[OK] を選択します**。

5. 2 番目のダイアログ ボックスが表示されます。コンピューターからアドイン マニフェストOffice登録を求めるダイアログ ボックスが表示されます。 [はい] を選択 **する必要があります**。

6. アドインがインストールされている。 アドイン コマンドの場合は、リボンまたはコンテキスト メニューに表示されます。 作業ウィンドウ アドインの場合は、作業ウィンドウが表示されます。

## <a name="sideload-an-office-add-in-in-office-on-the-web-manually"></a>Web 上のOfficeアドインを手動Officeサイドロードする

このメソッドはコマンド ラインを使用し、ホスト アプリケーション (Excel など) 内でのみコマンドを使用して実行できます。

1. Web [Officeを開きます](https://office.live.com/)。 **Excel、Word、または** PowerPoint でドキュメント **を開きます**。 [アドイン **]** セクションのリボンの[挿入] タブで、[アドイン] Office **選択します**。

1. [アドインOffice] ダイアログで、[**マイ** アドイン] タブを選択し、[マイ アドインの管理] を選択して、[マイ アドインのアップロード] を **選択します**。

    ![右上に [個人用アドインの管理] というドロップダウンがあり、その下に [マイ アドインのアップロード] オプションのドロップダウンがある [Office アドイン] ダイアログ](../images/office-add-ins-my-account.png)

1. アドイン マニフェスト ファイルを **参照** して、**[アップロード]** を選択します。

    ![[参照]、[アップロード]、[キャンセル] のボタンがある [アドインのアップロード] ダイアログ。](../images/upload-add-in.png)

1. アドインがインストールされていることを確認します。たとえば、アドイン コマンドである場合は、リボンまたはコンテキスト メニューのいずれかに表示されます。作業ウィンドウ アドインである場合は、ウィンドウが表示されます。

> [!NOTE]
> Microsoft Edge Office元の WebView (EdgeHTML) を使用してアドインをテストするには、追加の構成手順が必要です。 Windows コマンド プロンプトで、次の行を実行します `npx office-addin-dev-settings appcontainer EdgeWebView --loopback --yes` 。 Chromium ベースの Edge WebView2 Office使用している場合、これは必要ありません。 詳細については、「アドイン [で使用されるブラウザー」Office参照してください](../concepts/browsers-used-by-office-web-add-ins.md)。

## <a name="sideload-an-office-add-in"></a>新しいアドインOfficeサイドロードする

1. Microsoft 365 アカウントにサインインします。

2. ツールバーの左側起動ツールアプリ アプリを開き **、Excel、Word、** または **PowerPoint** を選択して、新しいドキュメントを作成します。

3. 手順 3 から 6 は、前のセクション「**Office on the web で Office アドインをサイドロードする**」のものと同じです。

## <a name="sideload-an-add-in-when-using-visual-studio"></a>Visual Studio の使用時にアドインをサイドロードする

アドインの開発に Visual Studioを使用している場合、サイドロードするプロセスは Web への手動サイドロードに似ています。 アドインの開発に Visual Studio を使用している場合、サイドロードするプロセスは似ています。唯一の違いは、マニフェストの **SourceURL** 要素の値を更新して、アドインが展開されている完全な URL を含める必要がある点です。

> [!NOTE]
> アドインは Visual Studio から Office on the web にサイドロードできますが、Visual Studio からはデバッグできません。 デバッグするには、ブラウザー デバッグ ツールを使用する必要があります。 詳細については、「[Office on the web でアドインをデバッグする](debug-add-ins-in-office-online.md)」を参照してください。

1. Visual Studio で、[**表示**]  >  [**プロパティ ウィンドウ**] の順に選択して [**プロパティ**] ウィンドウを表示させます。
2. [**ソリューション エクスプローラー**] で Web プロジェクトを選択します。 プロジェクトのプロパティが [**プロパティ**] ウィンドウに表示されます。
3. [プロパティ] ウィンドウで、[**SSL URL**] をコピーします。
4. アドイン プロジェクトで、マニフェスト XML ファイルを開きます。 編集しているのがソース XML であることを確認します。 一部の種類のプロジェクトでは、Visual Studio は XML のビジュアル ビューを開きますが、これは次の手順で使用できません。
5. **~remoteAppUrl/** のすべてのインスタンスを検索し、先ほどコピーした SSL URL と置き換えます。 プロジェクトの種類に応じていくつかの置換が表示され、新しい URL の表示は `https://localhost:44300/Home.html` に似たものになりま。
6. XML ファイルを保存します。
7. Web プロジェクトを右クリックして、[**デバッグ**]  >  [**新しいインスタンスを開始**] の順に選択します。 これにより、Office を起動することなく Web プロジェクトが実行されます。
8. 前述の「[Office on the web で Office アドインをサイドロードする](#sideload-an-office-add-in-in-office-on-the-web)」で説明した手順を使用して、Office on the web からアドインをサイドロードします。

## <a name="remove-a-sideloaded-add-in"></a>サイドロードされたアドインを削除する

ブラウザーのキャッシュをクリアすることで、以前にサイドロードされたアドインを削除できます。 アドインのマニフェストに変更を加えた場合 (たとえば、アイコンのファイル名やアドイン コマンドのテキストを更新する場合 [)、Office](clear-cache.md) キャッシュをクリアしてから、更新されたマニフェストを使用してアドインを再サイドロードする必要があります。 これを実行することにより、アドインは更新されたマニフェストの記載どおりに Office で表示されるようになります。

## <a name="see-also"></a>関連項目

- [iPad と Mac で Office アドインをサイドロードする](sideload-an-office-add-in-on-ipad-and-mac.md)
- [テスト用に Outlook アドインをサイドロードする](../outlook/sideload-outlook-add-ins-for-testing.md)
- [Office のキャッシュをクリアする](clear-cache.md)
