---
title: テスト用に Outlook アドインをサイドロードする
description: サイドロードを使用して、最初にアドイン カタログに置かずに、テスト用に Outlook アドインをインストールします。
ms.date: 10/22/2021
ms.localizationpriority: medium
ms.openlocfilehash: 43007ece67d85f584a682b7503f1b59e0d19ad5b
ms.sourcegitcommit: e4d98eb90e516b9c90e3832f3212caf48691acf6
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/22/2021
ms.locfileid: "60537507"
---
# <a name="sideload-outlook-add-ins-for-testing"></a>テスト用に Outlook アドインをサイドロードする

サイドロードを使用すると、最初にアドイン カタログに置かなくても、テスト用に Outlook アドインをインストールすることができます。

> [!IMPORTANT]
> Outlook アドインがモバイルをサポートしている場合は、web、Windows、または Mac の Outlook クライアントに関するこの記事の指示に従ってマニフェストをサイドロードし[、「Outlook Mobile](outlook-mobile-addins.md#testing-your-add-ins-on-mobile)用アドイン」の「モバイルでアドインをテストする」の記事のガイダンスに従います。

## <a name="sideload-automatically"></a>サイドロードが自動的に実行される

Office アドイン用の[Yeoman](https://github.com/OfficeDev/generator-office)ジェネレーターを使用して Outlook アドインを作成した場合は、サイドローディングは Windows のコマンド ラインを通じて行うのが最善です。 これにより、1 つのコマンドでサポートされているすべてのデバイスでツールとサイドロードを利用できます。

1. このWindowsコマンド プロンプトを開き、Yeoman が生成したアドイン プロジェクトのルート ディレクトリに移動します。 コマンド`npm start`を実行します。

1. ユーザー Outlookは、デスクトップ コンピューター上のOutlookに自動的にサイドロードされます。 アドインをサイドロードしようとして、マニフェスト ファイルの名前と場所を一覧に表示するダイアログが表示されます。 **[OK]** を選択し、マニフェストを登録します。

    > [!IMPORTANT]
    > マニフェストにエラーが含まれているか、マニフェストへのパスが無効な場合は、エラー メッセージが表示されます。

1. マニフェストにエラーが含まれるが、パスが有効な場合、アドインはサイドロードされ、デスクトップと Outlook on the web の両方で使用できます。 また、サポートされているすべてのデバイスにインストールされます。

## <a name="sideload-manually"></a>サイドロードを手動で実行する

前のセクションで説明したコマンド ラインから自動的にサイドローディングすることを強く推奨しますが、Outlook クライアントに基づいて Outlook アドインを手動でサイドロードすることもできます。

### <a name="outlook-on-the-web"></a>Outlook on the web

新しいバージョンまたはクラシック バージョンを使用Outlook on the webアドインをサイドローディングするプロセスは異なります。

- メールボックスのツールバーが次の図のような場合、「[新しい Outlook on the web のアドインをサイドロードする](#new-outlook-on-the-web)」を参照してください。

    ![新しいツール バーの部分的なOutlook on the webです。](../images/outlook-on-the-web-new-toolbar.png)

- メールボックスのツールバーが次の図のような場合、「[従来の Outlook on the web のアドインをサイドロードする](#classic-outlook-on-the-web)」を参照してください。

    ![従来のツール バーの一部Outlook on the webスクリーンショット。](../images/outlook-on-the-web-classic-toolbar.png)

> [!NOTE]
> 組織のメールボックスのツールバーにロゴが含まれている場合、上の図に示されるものと表示が少し異なる場合があります。

### <a name="new-outlook-on-the-web"></a>新しいOutlook on the web

1. [[Outlook on the web]](https://outlook.office.com) に進みます。

1. 新しいメッセージを作成します。

1. 新しいメッセージの下部で [**...**] を選択し、表示されるメニューから [**アドインを取得**] を選択します。

    ![[アドインの取得] オプションが強調表示Outlook on the web新しいウィンドウのメッセージ作成ウィンドウ。](../images/outlook-on-the-web-new-get-add-ins.png)

1. [**Outlook のアドイン**] ダイアログ ボックスで、[**個人用アドイン**] を選択します。

    ![[自分のOutlook] を選択した新しいOutlook on the webダイアログ ボックスのアドイン。](../images/outlook-on-the-web-new-my-add-ins.png)

1. ダイアログ ボックスの下部にある [**カスタム アドイン**] セクションに移動します。 [**カスタム アドインを追加**] リンクを選択し、[**ファイルから追加**] を選択します。

    ![[ファイルから追加] オプションをポイントするアドインのスクリーンショットを管理します。](../images/outlook-sideload-desktop-add-from-file.png)

1. カスタム アドインのマニフェスト ファイルを探してインストールします。インストール中にすべてのプロンプトを受け入れます。

### <a name="classic-outlook-on-the-web"></a>クラシック Outlook on the web

1. [[Outlook on the web]](https://outlook.office.com) に進みます。

1. ツールバー右上のセクションにあるギア アイコンを選択し、[**アドインの管理**] を選択します。

    ![Outlook on the webアドインの管理オプションをポイントするスクリーンショットを作成します。](../images/outlook-sideload-web-manage-integrations.png)

1. **アドインの管理** ページで、**[アドイン]** を選択してから、**[個人用アドイン]** を選択します。

    ![Outlook on the web[マイ アドイン] が選択されている場合は、[ストア] ダイアログボックスを開きます。](../images/outlook-sideload-store-select-add-ins.png)

1. ダイアログ ボックスの下部にある [**カスタム アドイン**] セクションに移動します。 [**カスタム アドインを追加**] リンクを選択し、[**ファイルから追加**] を選択します。

    ![[ファイルから追加] オプションをポイントするアドインのスクリーンショットを管理します。](../images/outlook-sideload-desktop-add-from-file.png)

1. カスタム アドインのマニフェスト ファイルを探してインストールします。インストール中にすべてのプロンプトを受け入れます。

### <a name="outlook-on-the-desktop"></a>Outlookの設定

### <a name="outlook-2016-or-later"></a>Outlook 2016以降

1. [Outlook 2016または Mac で、Windows以降を開きます。

1. リボンで [**アドインを取得**] ボタンを選択します。

    ![Outlook 2016アドインの取得] ボタンをポイントするリボンを選択します。](../images/outlook-sideload-desktop-store.png)

    > [!IMPORTANT]
    > バージョンの [アドインの取得] ボタンが表示されていない場合は、次Outlook選択します。
    >
    > - **リボン** の [ストア] ボタン (使用可能な場合)。
    >
    >   または
    >
    > - **[** ファイル] メニューの [情報] タブの[アドインの管理]ボタンを選択して、[アドイン] ダイアログボックスを開Outlook on the web。 <br>Web エクスペリエンスの詳細については、前のセクションの「アドインをサイドロードする」を参照[Outlook on the web。](#outlook-on-the-web)

1. ダイアログの上部にタブがある場合は、[アドイン] タブが **選択** されている必要があります。 [ **個人用アドイン**] を選びます。

    ![Outlook 2016[マイ アドイン] が選択されている場合は、[ストア] ダイアログボックスを開きます。](../images/outlook-sideload-store-select-add-ins.png)

1. ダイアログ ボックスの下部にある **[カスタム アドイン]** セクションに移動します。 **[カスタム アドインを追加]** リンクを選択し、**[ファイルから追加]** を選択します。

    ![[ファイルから追加] オプションをポイントするスクリーンショットを保存します。](../images/outlook-sideload-desktop-add-from-file.png)

1. カスタム アドインのマニフェスト ファイルを探してインストールします。インストール中にすべてのプロンプトを受け入れます。

### <a name="outlook-2013"></a>Outlook 2013

1. 2013 Outlook 2013 を開Windows。

1. [ファイル **] メニューを** 選択し、[情報] タブの [アドインの管理] ボタンを選択しますOutlookブラウザーで Web バージョンが開きます。

1. [アドインのサイドロード][セクションの](#outlook-on-the-web)手順に従って、Outlook on the webのバージョンに従Outlook on the web。

## <a name="remove-a-sideloaded-add-in"></a>サイドロードされたアドインを削除する

すべてのバージョンの Outlook、サイドロードされたアドインを削除するキーは、インストールされているアドインを一覧表示する[マイ アドイン] ダイアログです。アドインの省略記号 ( `...` ) を選択し、[削除] を **選択します**。

Outlook クライアントの[マイ アドイン] ダイアログ ボックスに移動するには、この記事の前のセクションで手動[](#sideload-manually)サイドローディングの最後の手順を使用します。

サイドロードされたアドインを Outlook から削除するには、この記事で説明した手順を使用して、インストールされているアドインを一覧表示するダイアログボックスの [カスタム アドイン] セクションでアドインを検索します。アドインの省略記号 ( ) を選択し、[削除] を選択して、その特定の `...` アドインを削除します。  ダイアログを閉じます。

## <a name="see-also"></a>関連項目

- [Outlook Mobile のアドイン](outlook-mobile-addins.md)