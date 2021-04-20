---
title: テスト用に Outlook アドインをサイドロードする
description: サイドロードを使用して、最初にアドイン カタログに置かずに、テスト用に Outlook アドインをインストールします。
ms.date: 02/10/2021
localization_priority: Normal
ms.openlocfilehash: b783b815af84a7fd8b4abd52cdd8e0925bfb9ecf
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234248"
---
# <a name="sideload-outlook-add-ins-for-testing"></a>テスト用に Outlook アドインをサイドロードする

サイドロードを使用すると、最初にアドイン カタログに置かなくても、テスト用に Outlook アドインをインストールすることができます。

## <a name="sideload-automatically"></a>自動的にサイドロードする

Office アドイン用 [の Yeoman](https://github.com/OfficeDev/generator-office)ジェネレーターを使用して Outlook アドインを作成した場合は、コマンド ラインを使用してサイドローディングを行うのが最適です。 これにより、ツールを利用し、サポートされているデバイス全体にサイドロードするコマンドが 1 つになります。

1. コマンド ラインを使用して、Yeoman によって生成されたアドイン プロジェクトのルート ディレクトリに移動します。 コマンド`npm start`を実行します。

2. Outlook アドインは、デスクトップ コンピューター上の Outlook に自動的にサイドロードします。 アドインをサイドロードしようとして、マニフェスト ファイルの名前と場所を一覧で示すダイアログが表示されます。 **[OK]** を選択します。マニフェストが登録されます。

> [!IMPORTANT]
> マニフェストにエラーが含まれているか、マニフェストへのパスが無効な場合は、エラー メッセージが表示されます。

3. マニフェストにエラーが含まれているのにパスが有効な場合、アドインはサイドロードされ、デスクトップと Outlook on the web の両方で利用できます。 また、サポートされているデバイスすべてにもインストールされます。

## <a name="sideload-manually"></a>手動でのサイドロード

前のセクションで説明したコマンド ラインを使用して自動的にサイドロードすることを強く推奨しますが、Outlook クライアントに基づいて Outlook アドインを手動でサイドロードすることもできます。

### <a name="outlook-on-the-web"></a>Outlook on the web

Outlook on the web でアドインをサイドロードするプロセスは、新しいバージョンを使用しているか、クラシック バージョンを使用しているかによって異なります。

- メールボックスのツールバーが次の図のような場合、「[新しい Outlook on the web のアドインをサイドロードする](#new-outlook-on-the-web)」を参照してください。

    ![新しい Outlook on the web の部分的なスクリーンショット](../images/outlook-on-the-web-new-toolbar.png)

- メールボックスのツールバーが次の図のような場合、「[従来の Outlook on the web のアドインをサイドロードする](#classic-outlook-on-the-web)」を参照してください。

    ![従来の Outlook on the web の部分的なスクリーンショット](../images/outlook-on-the-web-classic-toolbar.png)

> [!NOTE]
> 組織のメールボックスのツールバーにロゴが含まれている場合、上の図に示されるものと表示が少し異なる場合があります。

### <a name="new-outlook-on-the-web"></a>新しい Outlook on the web

1. [[Outlook on the web]](https://outlook.office.com) に進みます。

1. 新しいメッセージを作成します。

1. 新しいメッセージの下部で [**...**] を選択し、表示されるメニューから [**アドインを取得**] を選択します。

    ![[アドインを取得] オプションが強調表示された Outlook on the web のメッセージ作成ウィンドウ](../images/outlook-on-the-web-new-get-add-ins.png)

1. [**Outlook のアドイン**] ダイアログ ボックスで、[**個人用アドイン**] を選択します。

    ![[個人用アドイン] が選択された 新しい Outlook on the web の [Outlook のアドイン] ダイアログ ボックス](../images/outlook-on-the-web-new-my-add-ins.png)

1. ダイアログ ボックスの下部にある [**カスタム アドイン**] セクションに移動します。 [**カスタム アドインを追加**] リンクを選択し、[**ファイルから追加**] を選択します。

    ![ファイル オプションからの追加を示すアドイン スクリーンショットの管理](../images/outlook-sideload-desktop-add-from-file.png)

1. カスタム アドインのマニフェスト ファイルを探してインストールします。インストール中にすべてのプロンプトを受け入れます。

### <a name="classic-outlook-on-the-web"></a>従来の Outlook on the web

1. [[Outlook on the web]](https://outlook.office.com) に進みます。

1. ツールバー右上のセクションにあるギア アイコンを選択し、[**アドインの管理**] を選択します。

    ![[アドインの管理] オプションを示す Outlook on the web のスクリーンショット](../images/outlook-sideload-web-manage-integrations.png)

1. **アドインの管理** ページで、**[アドイン]** を選択してから、**[個人用アドイン]** を選択します。

    ![Outlook on the web の [ストア] ダイアログ ボックスで [個人用アドイン] を選択しているところ](../images/outlook-sideload-store-select-add-ins.png)

1. ダイアログ ボックスの下部にある [**カスタム アドイン**] セクションに移動します。 [**カスタム アドインを追加**] リンクを選択し、[**ファイルから追加**] を選択します。

    ![ファイル オプションからの追加を示すアドイン スクリーンショットの管理](../images/outlook-sideload-desktop-add-from-file.png)

1. カスタム アドインのマニフェスト ファイルを探してインストールします。インストール中にすべてのプロンプトを受け入れます。

### <a name="outlook-on-the-desktop"></a>デスクトップ上の Outlook

#### <a name="outlook-2016-or-later"></a>Outlook 2016 以降

1. Windows または Mac で Outlook 2016 以降を開きます。

1. リボンで [**アドインを取得**] ボタンを選択します。

    ![[アドインの取得] ボタンを指す Outlook 2016 リボン](../images/outlook-sideload-desktop-store.png)

    > [!IMPORTANT]
    > Outlook のバージョンに [アドインの取得] ボタンが表示されていない場合は、次を選択します。
    >
    > - **リボン** の [ストア] ボタン (使用可能な場合)。
    >
    >   または
    >
    > - **[** ファイル] メニューの [情報] タブで[アドインの管理]ボタンを選択し、Outlook on the web で [アドイン] ダイアログを開きます。 <br>Web エクスペリエンスの詳細については、前のセクションの「Outlook on the web でアドインをサイドロードする [」を参照してください](#outlook-on-the-web)。

1. ダイアログの上部付近にタブがある場合は、[アドイン] タブ **が選択されている** 必要があります。 [ **個人用アドイン**] を選びます。

    ![Outlook 2016 の [ストア] ダイアログ ボックスで [個人用アドイン] を選択しているところ](../images/outlook-sideload-store-select-add-ins.png)

1. ダイアログ ボックスの下部にある **[カスタム アドイン]** セクションに移動します。 **[カスタム アドインを追加]** リンクを選択し、**[ファイルから追加]** を選択します。

    ![[ファイルから追加] オプションを示す [ストア] のスクリーンショット](../images/outlook-sideload-desktop-add-from-file.png)

1. カスタム アドインのマニフェスト ファイルを探してインストールします。インストール中にすべてのプロンプトを受け入れます。

#### <a name="outlook-2013"></a>Outlook 2013

1. Windows で Outlook 2013 を開きます。

1. [ファイル **] メニュー** を選択し、[情報] タブの [アドインの管理] ボタン **を選択** します。Outlook はブラウザーで Web バージョンを開きます。

1. Outlook on the web の [バージョンに応じて、Web](#outlook-on-the-web) 上の Outlook でアドインをサイドロードするセクションの手順に従います。

## <a name="remove-a-sideloaded-add-in"></a>サイドロードされたアドインを削除する

すべてのバージョンの Outlook では、サイドロードされたアドインを削除するための鍵は、インストールされているアドインを一覧表示する [マイ アドイン] ダイアログです。アドインの省略記号 ( ) を選択 `...` し、[削除] を選択 **します**。

Outlook クライアントの [**マイ** アドイン] ダイアログ ボックスに移動するには、この記事の前の [](#sideload-manually)セクションで手動サイドロードの最後の手順を使用します。

サイドロードされたアドインを Outlook から削除するには、この記事で前述した手順を使用して、インストールされているアドインを一覧表示するダイアログ ボックスの [カスタム アドイン] セクションでアドインを検索します。アドインの省略記号 ( ) を選択し、[削除] を選択してその特定のアドイン `...` を削除します。 

