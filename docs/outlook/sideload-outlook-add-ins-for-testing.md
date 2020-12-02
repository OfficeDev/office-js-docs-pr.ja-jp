---
title: テスト用に Outlook アドインをサイドロードする
description: サイドロードを使用して、最初にアドイン カタログに置かずに、テスト用に Outlook アドインをインストールします。
ms.date: 12/01/2020
localization_priority: Normal
ms.openlocfilehash: dea2125ccd64eba2e3f1695c8ca1111a710321a4
ms.sourcegitcommit: c2fd7f982f3da748ef6be5c3a7434d859f8b46b9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/02/2020
ms.locfileid: "49530928"
---
# <a name="sideload-outlook-add-ins-for-testing"></a>テスト用に Outlook アドインをサイドロードする

サイドロードを使用すると、最初にアドイン カタログに置かなくても、テスト用に Outlook アドインをインストールすることができます。

## <a name="sideload-an-add-in-in-outlook-on-the-web"></a>Outlook on the web のアドインをサイドロードする

Web 上の Outlook でアドインをサイドロードするためのプロセスは、新しいバージョンとクラシックバージョンのどちらを使用しているかによって異なります。

- メールボックスのツールバーが次の図のような場合、「[新しい Outlook on the web のアドインをサイドロードする](#sideload-an-add-in-in-the-new-outlook-on-the-web)」を参照してください。

    ![新しい Outlook on the web の部分的なスクリーンショット](../images/outlook-on-the-web-new-toolbar.png)

- メールボックスのツールバーが次の図のような場合、「[従来の Outlook on the web のアドインをサイドロードする](#sideload-an-add-in-in-classic-outlook-on-the-web)」を参照してください。

    ![従来の Outlook on the web の部分的なスクリーンショット](../images/outlook-on-the-web-classic-toolbar.png)

> [!NOTE]
> 組織のメールボックスのツールバーにロゴが含まれている場合、上の図に示されるものと表示が少し異なる場合があります。

### <a name="sideload-an-add-in-in-the-new-outlook-on-the-web"></a>新しい Outlook on the web のアドインをサイドロードする

1. [Office 365 の Outlook](https://outlook.office.com) に移動します。

1. Outlook on the web で新しいメッセージを作成します。

1. 新しいメッセージの下部で [**...**] を選択し、表示されるメニューから [**アドインを取得**] を選択します。

    ![[アドインを取得] オプションが強調表示された Outlook on the web のメッセージ作成ウィンドウ](../images/outlook-on-the-web-new-get-add-ins.png)

1. [**Outlook のアドイン**] ダイアログ ボックスで、[**個人用アドイン**] を選択します。

    ![[個人用アドイン] が選択された 新しい Outlook on the web の [Outlook のアドイン] ダイアログ ボックス](../images/outlook-on-the-web-new-my-add-ins.png)

1. ダイアログ ボックスの下部にある [**カスタム アドイン**] セクションに移動します。 [**カスタム アドインを追加**] リンクを選択し、[**ファイルから追加**] を選択します。

    ![ファイル オプションからの追加を示すアドイン スクリーンショットの管理](../images/outlook-sideload-desktop-add-from-file.png)

1. カスタム アドインのマニフェスト ファイルを探してインストールします。インストール中にすべてのプロンプトを受け入れます。

### <a name="sideload-an-add-in-in-classic-outlook-on-the-web"></a>従来の Outlook on the web のアドインをサイドロードする

1. [Office 365 の Outlook](https://outlook.office.com) に移動します。

1. ツールバー右上のセクションにあるギア アイコンを選択し、[**アドインの管理**] を選択します。

    ![[アドインの管理] オプションを示す Outlook on the web のスクリーンショット](../images/outlook-sideload-web-manage-integrations.png)

1. **アドインの管理** ページで、**[アドイン]** を選択してから、**[個人用アドイン]** を選択します。

    ![Outlook on the web の [ストア] ダイアログ ボックスで [個人用アドイン] を選択しているところ](../images/outlook-sideload-store-select-add-ins.png)

1. ダイアログ ボックスの下部にある [**カスタム アドイン**] セクションに移動します。 [**カスタム アドインを追加**] リンクを選択し、[**ファイルから追加**] を選択します。

    ![ファイル オプションからの追加を示すアドイン スクリーンショットの管理](../images/outlook-sideload-desktop-add-from-file.png)

1. カスタム アドインのマニフェスト ファイルを探してインストールします。インストール中にすべてのプロンプトを受け入れます。

## <a name="sideload-an-add-in-in-outlook-on-the-desktop"></a>Outlook on the desktop のアドインをサイドロードする

### <a name="outlook-2016-or-later"></a>Outlook 2016 以降

1. Windows または Mac で Outlook 2016 以降を開きます。

1. リボンで [**アドインを取得**] ボタンを選択します。

    ![[アドインの取得] ボタンをポイントする Outlook 2016 リボン](../images/outlook-sideload-desktop-store.png)

    > [!IMPORTANT]
    > 使用している Outlook のバージョンで [アドインの **取得** ] ボタンが表示されない場合は、次のように選択します。
    >
    > - リボン上の [**ストア**] ボタン (使用可能な場合)。
    >
    >   OR
    >
    > - [**ファイル**] メニューの [**情報**] タブで [アドインの **管理**] をクリックして、Outlook on the web の **[アドイン] ダイアログを** 開きます。<br>Web の詳細については、「 [Outlook on the web in a サイドロード](#sideload-an-add-in-in-outlook-on-the-web)in the web」を参照してください。

1. ダイアログボックスの上部付近にタブがある場合は、[ **アドイン** ] タブが選択されていることを確認します。 [ **個人用アドイン**] を選びます。

    ![Outlook 2016 の [ストア] ダイアログ ボックスで [個人用アドイン] を選択しているところ](../images/outlook-sideload-store-select-add-ins.png)

1. ダイアログ ボックスの下部にある **[カスタム アドイン]** セクションに移動します。 **[カスタム アドインを追加]** リンクを選択し、**[ファイルから追加]** を選択します。

    ![[ファイルから追加] オプションを示す [ストア] のスクリーンショット](../images/outlook-sideload-desktop-add-from-file.png)

1. カスタム アドインのマニフェスト ファイルを探してインストールします。インストール中にすべてのプロンプトを受け入れます。

### <a name="outlook-2013"></a>Outlook 2013

1. Windows で Outlook 2013 を開きます。

1. [**ファイル**] メニューを選択し、[**情報**] タブの [アドインの **管理**] をクリックします。Outlook は、ブラウザーで web バージョンを開きます。

1. Web 上の Outlook のバージョンに応じて、「 [web 上の outlook でアドインをサイドロード](#sideload-an-add-in-in-outlook-on-the-web) する」セクションの手順を実行します。

## <a name="remove-a-sideloaded-add-in"></a>サイドロードアドインを削除する

Outlook からサイドロードアドインを削除するには、この記事で前述した手順を使用して、インストールされているアドインを一覧表示するダイアログボックスの [ **カスタムアドイン** ] セクションでアドインを見つけます。アドインの省略記号 () を選択 `...` し、[ **削除** ] を選択して、その特定のアドインを削除します。