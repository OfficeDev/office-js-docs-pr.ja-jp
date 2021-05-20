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

## <a name="sideload-automatically"></a>サイドロードは自動的に行われます

Office アドイン[用の Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)を使用して Outlook アドインを作成した場合は、コマンド ラインを使用してサイドローディングを実行することをお勧めします。 これは、1つのコマンドでサポートされているすべてのデバイス全体で私たちのツールとサイドロードを利用します。

1. コマンド ラインを使用して、Yeoman によって生成されたアドイン プロジェクトのルート ディレクトリに移動します。 コマンド`npm start`を実行します。

1. Outlook アドインは、デスクトップ コンピュータ上のOutlookに自動的にサイドロードされます。 アドインのサイドロードが試行され、マニフェスト ファイルの名前と場所が一覧表示されたダイアログが表示されます。 **[OK]** を選択すると、マニフェストが登録されます。

    > [!IMPORTANT]
    > マニフェストにエラーが含まれている場合、またはマニフェストへのパスが無効な場合は、エラー メッセージが表示されます。

1. マニフェストにエラーがなく、パスが有効な場合、アドインはサイドロードされ、デスクトップと Web 上のOutlookの両方で使用できるようになります。 また、サポートされているすべてのデバイスにインストールされます。

## <a name="sideload-manually"></a>手動でサイドロード

前のセクションで説明したように、コマンド ラインを使用して自動的にサイドローディングを行うことを強くお勧めしますが、Outlook クライアントに基づいてOutlookアドインを手動でサイドロードすることもできます。

### <a name="outlook-on-the-web"></a>Outlook on the web

Web 上のアドインをサイドロードOutlookプロセスは、新しいバージョンとクラシック バージョンのどちらを使用しているかによって異なります。

- メールボックスのツールバーが次の図のような場合、「[新しい Outlook on the web のアドインをサイドロードする](#new-outlook-on-the-web)」を参照してください。

    ![新しい Outlook on the web の部分的なスクリーンショット](../images/outlook-on-the-web-new-toolbar.png)

- メールボックスのツールバーが次の図のような場合、「[従来の Outlook on the web のアドインをサイドロードする](#classic-outlook-on-the-web)」を参照してください。

    ![従来の Outlook on the web の部分的なスクリーンショット](../images/outlook-on-the-web-classic-toolbar.png)

> [!NOTE]
> 組織のメールボックスのツールバーにロゴが含まれている場合、上の図に示されるものと表示が少し異なる場合があります。

### <a name="new-outlook-on-the-web"></a>Web 上の新しいOutlook

1. [[Outlook on the web]](https://outlook.office.com) に進みます。

1. 新しいメッセージを作成します。

1. 新しいメッセージの下部で [**...**] を選択し、表示されるメニューから [**アドインを取得**] を選択します。

    ![[アドインを取得] オプションが強調表示された Outlook on the web のメッセージ作成ウィンドウ](../images/outlook-on-the-web-new-get-add-ins.png)

1. [**Outlook のアドイン**] ダイアログ ボックスで、[**個人用アドイン**] を選択します。

    ![[個人用アドイン] が選択された 新しい Outlook on the web の [Outlook のアドイン] ダイアログ ボックス](../images/outlook-on-the-web-new-my-add-ins.png)

1. ダイアログ ボックスの下部にある [**カスタム アドイン**] セクションに移動します。 [**カスタム アドインを追加**] リンクを選択し、[**ファイルから追加**] を選択します。

    ![ファイル オプションからの追加を示すアドイン スクリーンショットの管理](../images/outlook-sideload-desktop-add-from-file.png)

1. カスタム アドインのマニフェスト ファイルを探してインストールします。インストール中にすべてのプロンプトを受け入れます。

### <a name="classic-outlook-on-the-web"></a>ウェブ上の古典的なOutlook

1. [[Outlook on the web]](https://outlook.office.com) に進みます。

1. ツールバー右上のセクションにあるギア アイコンを選択し、[**アドインの管理**] を選択します。

    ![[アドインの管理] オプションを示す Outlook on the web のスクリーンショット](../images/outlook-sideload-web-manage-integrations.png)

1. **アドインの管理** ページで、**[アドイン]** を選択してから、**[個人用アドイン]** を選択します。

    ![Outlook on the web の [ストア] ダイアログ ボックスで [個人用アドイン] を選択しているところ](../images/outlook-sideload-store-select-add-ins.png)

1. ダイアログ ボックスの下部にある [**カスタム アドイン**] セクションに移動します。 [**カスタム アドインを追加**] リンクを選択し、[**ファイルから追加**] を選択します。

    ![ファイル オプションからの追加を示すアドイン スクリーンショットの管理](../images/outlook-sideload-desktop-add-from-file.png)

1. カスタム アドインのマニフェスト ファイルを探してインストールします。インストール中にすべてのプロンプトを受け入れます。

### <a name="outlook-on-the-desktop"></a>デスクトップ上のOutlook

#### <a name="outlook-2016-or-later"></a>Outlook 2016以降

1. Windowsまたは Mac でOutlook 2016以降で開きます。

1. リボンで [**アドインを取得**] ボタンを選択します。

    ![[アドインの取得] ボタンをポイントするリボンをOutlook 2016します。](../images/outlook-sideload-desktop-store.png)

    > [!IMPORTANT]
    > 使用しているバージョンのOutlookに [**アドインの取得**] ボタンが表示されない場合は、次のオプションを選択します。
    >
    > - リボン上の **[ストア**] ボタン (可能な場合)
    >
    >   OR
    >
    > - **[ファイル]** メニューをクリックし、[**情報**] タブの [**アドインの管理**] ボタンを選択して、Web のOutlookで [**アドイン**] ダイアログを開きます。<br>Web エクスペリエンスの詳細については、前のセクション「web 上の[Outlookアドインをサイドロード](#outlook-on-the-web)する」を参照してください。

1. ダイアログの上部にタブがある場合は、[ **アドイン** ] タブが選択されていることを確認します。 [ **個人用アドイン**] を選びます。

    ![Outlook 2016 の [ストア] ダイアログ ボックスで [個人用アドイン] を選択しているところ](../images/outlook-sideload-store-select-add-ins.png)

1. ダイアログ ボックスの下部にある **[カスタム アドイン]** セクションに移動します。 **[カスタム アドインを追加]** リンクを選択し、**[ファイルから追加]** を選択します。

    ![[ファイルから追加] オプションを示す [ストア] のスクリーンショット](../images/outlook-sideload-desktop-add-from-file.png)

1. カスタム アドインのマニフェスト ファイルを探してインストールします。インストール中にすべてのプロンプトを受け入れます。

#### <a name="outlook-2013"></a>Outlook 2013

1. Windows Outlook 2013をオープンします。

1. [**ファイル]** メニューを選択し、[**情報**] タブ **の [アドインの管理**] ボタンを選択Outlook、ブラウザで Web バージョンを開きます。

1. Web 上のOutlookのバージョン[に従って、web セクションの [Outlookのアドインをサイドロード](#outlook-on-the-web)するの手順に従います。

## <a name="remove-a-sideloaded-add-in"></a>サイドローデッド アドインを削除する

Outlookのすべてのバージョンで、サイドローデッド アドインを削除するキーは、インストールされている **アドイン** を一覧表示する [マイ アドイン] ダイアログです。アドインの省略記号 ( `...` ) を選択し、[**削除**] を選択します。

Outlook クライアントの [**個人用アドイン**] ダイアログ ボックスに移動するには、この記事の前のセクションで [説明した手動サイドローディング](#sideload-manually)の最後の手順を使用します。

サイドローデッド アドインをOutlookから削除するには、この記事で説明した手順を使用して、インストールされているアドインの一覧が表示されるダイアログ ボックスの **[カスタム アドイン**] セクションでアドインを検索します。アドインの省略記号 ( `...` ) を選択し、[**削除**] を選択して、特定のアドインを削除します。 ダイアログを閉じます。
