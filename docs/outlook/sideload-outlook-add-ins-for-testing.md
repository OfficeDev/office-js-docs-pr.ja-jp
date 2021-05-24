---
title: テスト用に Outlook アドインをサイドロードする
description: サイドロードを使用して、最初にアドイン カタログに置かずに、テスト用に Outlook アドインをインストールします。
ms.date: 05/13/2021
localization_priority: Normal
ms.openlocfilehash: 9d0fb246f6522c745658a09fce6934ee44d5079a
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555193"
---
# <a name="sideload-outlook-add-ins-for-testing"></a>テスト用に Outlook アドインをサイドロードする

サイドロードを使用すると、最初にアドイン カタログに置かなくても、テスト用に Outlook アドインをインストールすることができます。

## <a name="sideload-automatically"></a>サイドロードが自動的に実行される

Office アドイン用の[Yeoman](https://github.com/OfficeDev/generator-office)ジェネレーターを使用して Outlook アドインを作成した場合は、コマンド ラインを使用してサイドローディングを行うのが最適です。 これにより、1 つのコマンドでサポートされているすべてのデバイスでツールとサイドロードを利用できます。

1. コマンド ラインを使用して、Yeoman が生成したアドイン プロジェクトのルート ディレクトリに移動します。 コマンド`npm start`を実行します。

1. ユーザー Outlookは、デスクトップ コンピューター上のOutlookに自動的にサイドロードされます。 アドインをサイドロードしようとして、マニフェスト ファイルの名前と場所を一覧に表示するダイアログが表示されます。 **[OK]** を選択し、マニフェストを登録します。

    > [!IMPORTANT]
    > マニフェストにエラーが含まれているか、マニフェストへのパスが無効な場合は、エラー メッセージが表示されます。

1. マニフェストにエラーが含まれるが、パスが有効な場合、アドインはサイドロードされ、デスクトップと web 上の Outlookで使用できます。 また、サポートされているすべてのデバイスにインストールされます。

## <a name="sideload-manually"></a>サイドロードを手動で実行する

前のセクションで説明したコマンド ラインから自動的にサイドローディングすることを強く推奨しますが、Outlook クライアントに基づいて Outlook アドインを手動でサイドロードすることもできます。

### <a name="outlook-on-the-web"></a>Outlook on the web

Web 上でアドインをサイドローディングOutlookプロセスは、新しいバージョンとクラシック バージョンの使用によって異なります。

- メールボックスのツールバーが次の図のような場合、「[新しい Outlook on the web のアドインをサイドロードする](#new-outlook-on-the-web)」を参照してください。

    ![新しい Outlook on the web の部分的なスクリーンショット](../images/outlook-on-the-web-new-toolbar.png)

- メールボックスのツールバーが次の図のような場合、「[従来の Outlook on the web のアドインをサイドロードする](#classic-outlook-on-the-web)」を参照してください。

    ![従来の Outlook on the web の部分的なスクリーンショット](../images/outlook-on-the-web-classic-toolbar.png)

> [!NOTE]
> 組織のメールボックスのツールバーにロゴが含まれている場合、上の図に示されるものと表示が少し異なる場合があります。

### <a name="new-outlook-on-the-web"></a>Web Outlookの新しい情報

1. [[Outlook on the web]](https://outlook.office.com) に進みます。

1. 新しいメッセージを作成します。

1. 新しいメッセージの下部で [**...**] を選択し、表示されるメニューから [**アドインを取得**] を選択します。

    ![[アドインを取得] オプションが強調表示された Outlook on the web のメッセージ作成ウィンドウ](../images/outlook-on-the-web-new-get-add-ins.png)

1. [**Outlook のアドイン**] ダイアログ ボックスで、[**個人用アドイン**] を選択します。

    ![[個人用アドイン] が選択された 新しい Outlook on the web の [Outlook のアドイン] ダイアログ ボックス](../images/outlook-on-the-web-new-my-add-ins.png)

1. ダイアログ ボックスの下部にある [**カスタム アドイン**] セクションに移動します。 [**カスタム アドインを追加**] リンクを選択し、[**ファイルから追加**] を選択します。

    ![ファイル オプションからの追加を示すアドイン スクリーンショットの管理](../images/outlook-sideload-desktop-add-from-file.png)

1. カスタム アドインのマニフェスト ファイルを探してインストールします。インストール中にすべてのプロンプトを受け入れます。

### <a name="classic-outlook-on-the-web"></a>Web Outlookクラシック コンテンツ

1. [[Outlook on the web]](https://outlook.office.com) に進みます。

1. ツールバー右上のセクションにあるギア アイコンを選択し、[**アドインの管理**] を選択します。

    ![[アドインの管理] オプションを示す Outlook on the web のスクリーンショット](../images/outlook-sideload-web-manage-integrations.png)

1. **アドインの管理** ページで、**[アドイン]** を選択してから、**[個人用アドイン]** を選択します。

    ![Outlook on the web の [ストア] ダイアログ ボックスで [個人用アドイン] を選択しているところ](../images/outlook-sideload-store-select-add-ins.png)

1. ダイアログ ボックスの下部にある [**カスタム アドイン**] セクションに移動します。 [**カスタム アドインを追加**] リンクを選択し、[**ファイルから追加**] を選択します。

    ![ファイル オプションからの追加を示すアドイン スクリーンショットの管理](../images/outlook-sideload-desktop-add-from-file.png)

1. カスタム アドインのマニフェスト ファイルを探してインストールします。インストール中にすべてのプロンプトを受け入れます。

### <a name="outlook-on-the-desktop"></a>Outlookの設定

#### <a name="outlook-2016-or-later"></a>Outlook 2016以降

1. [Outlook 2016または Mac で、Windows以降を開きます。

1. リボンで [**アドインを取得**] ボタンを選択します。

    ![Outlook 2016アドインの取得] ボタンをポイントするリボン](../images/outlook-sideload-desktop-store.png)

    > [!IMPORTANT]
    > バージョンの [アドインの取得] ボタンが表示されていない場合は、次Outlook選択します。
    >
    > - **リボン** の [ストア] ボタン (使用可能な場合)。
    >
    >   OR
    >
    > - **[** ファイル] メニューの [情報] タブの[アドインの管理]ボタンを選択して、Web 上の [アドイン] ダイアログ Outlook開きます。 <br>Web エクスペリエンスの詳細については、前のセクション「Web 上でアドインをサイドロードOutlook[参照してください](#outlook-on-the-web)。

1. ダイアログの上部にタブがある場合は、[アドイン] タブが **選択** されている必要があります。 [ **個人用アドイン**] を選びます。

    ![Outlook 2016 の [ストア] ダイアログ ボックスで [個人用アドイン] を選択しているところ](../images/outlook-sideload-store-select-add-ins.png)

1. ダイアログ ボックスの下部にある **[カスタム アドイン]** セクションに移動します。 **[カスタム アドインを追加]** リンクを選択し、**[ファイルから追加]** を選択します。

    ![[ファイルから追加] オプションを示す [ストア] のスクリーンショット](../images/outlook-sideload-desktop-add-from-file.png)

1. カスタム アドインのマニフェスト ファイルを探してインストールします。インストール中にすべてのプロンプトを受け入れます。

#### <a name="outlook-2013"></a>Outlook 2013

1. 2013 Outlook 2013 を開Windows。

1. [ファイル **] メニューを** 選択し、[情報] タブの [アドインの管理] ボタンを選択しますOutlookブラウザーで Web バージョンが開きます。

1. [Web 上の[](#outlook-on-the-web)アドインをサイドロードする] セクションOutlookに従って、Web 上のアドインのバージョンOutlook実行します。

## <a name="remove-a-sideloaded-add-in"></a>サイドロードされたアドインを削除する

すべてのバージョンの Outlook、サイドロードされたアドインを削除するキーは、インストールされているアドインを一覧表示する[マイ アドイン] ダイアログです。アドインの省略記号 ( `...` ) を選択し、[削除] を **選択します**。

Outlook クライアントの[マイ アドイン] ダイアログ ボックスに移動するには、この記事の前のセクションで手動[](#sideload-manually)サイドローディングの最後の手順を使用します。

サイドロードされたアドインを Outlook から削除するには、この記事で説明した手順を使用して、インストールされているアドインを一覧表示するダイアログボックスの [カスタム アドイン] セクションでアドインを検索します。アドインの省略記号 ( ) を選択し、[削除] を選択して、その特定の `...` アドインを削除します。  ダイアログを閉じます。
