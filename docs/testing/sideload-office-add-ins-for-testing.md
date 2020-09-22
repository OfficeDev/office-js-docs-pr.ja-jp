---
title: テスト用に Office on the web で Office アドインをサイドロードする
description: サイドロードを使用して、office で office アドインをテストします。
ms.date: 09/21/2020
localization_priority: Normal
ms.openlocfilehash: 709461d19fbf4602db3ba5bd9c40f495d0dbbd52
ms.sourcegitcommit: 4a03d8b3f676ee2d91114813cb81bce5da3c8d6b
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/22/2020
ms.locfileid: "48175536"
---
# <a name="sideload-office-add-ins-in-office-on-the-web-for-testing"></a>テスト用に Office on the web で Office アドインをサイドロードする

サイドロードを使用することで、最初にアドイン カタログに置かなくても、テスト用に Office アドインをインストールすることができます。 サイドローディングは、Microsoft 365 または web 上の Office のどちらかで実行できます。 2 つのプラットフォームで手順が少し異なります。

アドインをサイドロードするとき、アドイン マニフェストはブラウザーのローカル ストレージに格納されます。そのため、ブラウザーのキャッシュを消去したり、別のブラウザーに切り替えたりする場合、アドインを再びサイドロードする必要があります。

> [!NOTE]
> この記事で説明したようにサイドロードは、Word、Excel、および PowerPoint でサポートされています。Outlook アドインをサイドロードするには、「[テストのために Outlook アドインをサイドロードする](../outlook/sideload-outlook-add-ins-for-testing.md)」をご参照ください。

次のビデオでは、Office on the web またはデスクトップでアドインをサイドロードする手順について説明しています。

> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]

## <a name="sideload-an-office-add-in-in-office-on-the-web"></a>Office on the web で Office アドインをサイドロードする

1. [Web 上の Office を](https://office.live.com/)開きます。

2. [ **オンラインアプリを今すぐ開始する**] で、 **Excel**、 **Word**、または **PowerPoint**を選択します。新しいドキュメントを開きます。

3. リボンの [ **挿入** ] タブを開き、 **[アドイン] セクションで** 、[ **Office アドイン**] を選択します。

4. [ **Office アドイン** ] ダイアログボックスで、[ **個人用アドイン** ] タブ、[ **個人用アドインの管理**]、[ **個人用アドインのアップロード**] の順に選択します。

    ![右上に [個人用アドインの管理] というドロップダウンがあり、その下に [マイ アドインのアップロード] オプションのドロップダウンがある [Office アドイン] ダイアログ](../images/office-add-ins-my-account.png)

5. アドイン マニフェスト ファイルを**参照**して、**[アップロード]** を選択します。

    ![[参照]、[アップロード]、[キャンセル] のボタンがある [アドインのアップロード] ダイアログ。](../images/upload-add-in.png)

6. アドインがインストールされていることを確認します。たとえば、アドイン コマンドである場合は、リボンまたはコンテキスト メニューのいずれかに表示されます。作業ウィンドウ アドインである場合は、ウィンドウが表示されます。

> [!NOTE]
> Microsoft Edge で Office アドインをテストするには、追加の構成手順が必要です。 Windows コマンド プロンプトで、次のコマンドを実行します: `npx office-addin-dev-settings appcontainer EdgeWebView --loopback --yes`

## <a name="sideload-an-office-add-in-in-office-365"></a>Office 365 で Office アドインをサイドロードする

1. Microsoft 365 アカウントにサインインします。

2. ツールバーの左端にあるアプリ起動ツールを開き、 **Excel**、 **Word**、または **PowerPoint**を選択して、新しいドキュメントを作成します。

3. 手順 3 から 6 は、前のセクション「**Office on the web で Office アドインをサイドロードする**」のものと同じです。

## <a name="sideload-an-add-in-when-using-visual-studio"></a>Visual Studio の使用時にアドインをサイドロードする

アドインの開発に Visual Studio を使用している場合、サイドロードするプロセスは似ています。 アドインの開発に Visual Studio を使用している場合、サイドロードするプロセスは似ています。唯一の違いは、マニフェストの **SourceURL** 要素の値を更新して、アドインが展開されている完全な URL を含める必要がある点です。

> [!NOTE]
> アドインは Visual Studio から Office on the web にサイドロードできますが、Visual Studio からはデバッグできません。 デバッグするには、ブラウザー デバッグ ツールを使用する必要があります。 詳細については、「[Office on the web でアドインをデバッグする](debug-add-ins-in-office-online.md)」を参照してください。

1. Visual Studio で、[**表示**]  ->  [**プロパティ ウィンドウ**] の順に選択して [**プロパティ**] ウィンドウを表示させます。
2. [**ソリューション エクスプローラー**] で Web プロジェクトを選択します。 プロジェクトのプロパティが [**プロパティ**] ウィンドウに表示されます。
3. [プロパティ] ウィンドウで、[**SSL URL**] をコピーします。
4. アドイン プロジェクトで、マニフェスト XML ファイルを開きます。 編集しているのがソース XML であることを確認します。 一部の種類のプロジェクトでは、Visual Studio は XML のビジュアル ビューを開きますが、これは次の手順で使用できません。
5. **~remoteAppUrl/** のすべてのインスタンスを検索し、先ほどコピーした SSL URL と置き換えます。 プロジェクトの種類に応じていくつかの置換が表示され、新しい URL の表示は `https://localhost:44300/Home.html` に似たものになりま。
6. XML ファイルを保存します。
7. Web プロジェクトを右クリックして、[**デバッグ**]  ->  [**新しいインスタンスを開始**] の順に選択します。 これにより、Office を起動することなく Web プロジェクトが実行されます。
8. 前述の「[Office on the web で Office アドインをサイドロードする](#sideload-an-office-add-in-in-office-on-the-web)」で説明した手順を使用して、Office on the web からアドインをサイドロードします。

## <a name="remove-a-sideloaded-add-in"></a>サイドロードアドインを削除する

以前のサイドロードアドインを削除するには、ブラウザーのキャッシュをクリアする必要があります。 また、アドインのマニフェストを変更した場合 (たとえば、アイコンの更新ファイル名やアドインコマンドのテキスト)、キャッシュをクリアし、更新されたマニフェストを使用してアドインを再サイドロードする必要がある場合があります。 これを実行することにより、アドインは更新されたマニフェストの記載どおりに Office で表示されるようになります。
