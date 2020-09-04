---
title: テスト用に iPad と Mac で Office アドインをサイドロードする
description: サイドロードを使用して、iPad と Mac で Office アドインをテストします。
ms.date: 09/02/2020
localization_priority: Normal
ms.openlocfilehash: 7c5e9542c6e6f9abc96defde389b9543421b8529
ms.sourcegitcommit: 604361e55dee45c7a5d34c2fa6937693c154fc24
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/03/2020
ms.locfileid: "47364058"
---
# <a name="sideload-office-add-ins-on-ipad-and-mac-for-testing"></a>テスト用に iPad と Mac で Office アドインをサイドロードする

Office on iOS でアドインの実行状態を確認するには、iTunes を利用してアドインのマニフェストを iPad にサイドロードするか、Office on Mac でアドインのマニフェストを直接サイドロードします。このアクションでは、実行中にブレークポイントを設定したり、アドインのコードをデバッグしたりできませんが、その動作を確認したり、UI が使いやすいかどうかや、適切にレンダリングされているかどうかを確認できます。

## <a name="prerequisites-for-office-on-ios"></a>Office on iOS の前提条件

- [iTunes](https://www.apple.com/itunes/download/) がインストールされた Windows または Mac コンピューター。
  > [!IMPORTANT]
  > MacOS Catalina を実行している場合 [は、iTunes を使用でき](https://support.apple.com/HT210200) なくなりました。この記事の後半の「 [macos Catalina を使用して、Excel または Word でのアドインのサイドロード](#sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina) 」の手順に従ってください。

- [Excel](https://apps.apple.com/app/microsoft-excel/id586683407)または[Word](https://apps.apple.com/app/microsoft-word/id586447913)がインストールされた iOS 8.2 以降を実行している iPad と、同期ケーブル。

- テスト対象アドインのマニフェスト .xml ファイル。

## <a name="prerequisites-for-office-on-mac"></a>Office on Mac の前提条件

- [Office on Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) がインストールされていて OS X v10.10 "Yosemite" を実行している Mac。

- Word on Mac バージョン 15.18 (160109)。

- Excel on Mac バージョン 15.19 (160206)。

- PowerPoint on Mac バージョン 15.24 (160614)

- テスト対象アドインのマニフェスト .xml ファイル。

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad-using-itunes"></a>サイドロードを iTunes を使用して Excel または Word で iPad に追加する

1. 同期ケーブルを使用し、iPad をコンピューターに接続します。 初めて iPad をコンピューターに接続している場合は、 **このコンピューターを信頼するかどうか**を確認するメッセージが表示されます。 **[信頼する]** を選択して続行します。

2. iTunes で、メニュー バーの下にある **[iPad]** のアイコンをクリックします。

3. iTunes の左側の **[設定]** で、**[App]** をクリックします。

4. iTunes の右側で、**[ファイル共有]** までスクロールしてから、**[アドイン]** 列で **[Excel]** または **[Word]** をクリックします。

5. [ **Excel** ] 列または [ **Word ドキュメント** ] 列の下部で、[ **ファイルの追加**] を選択し、サイドロードするアドインの manifest.xml ファイルを選択します。

6. iPad で Excel または Word アプリを開きます。 Excel または Word アプリが既に実行されている場合は、[ **ホーム** ] ボタンを選択し、アプリを閉じて再起動します。

7. ドキュメントを開きます。

8. [**挿入**] タブで [**アドイン**] を選択します。 ([**挿入**] タブで、[**アドイン**] ボタンが表示されるまで、横にスクロールする必要がある場合があります)。サイドロードアドインは **、アドインの UI の**[**開発者**] 見出しの下に挿入できます。

    ![Excel アプリでアドインを挿入](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina"></a>MacOS Catalina を使用して、Excel または Word でアドインをサイドロードします。

> [!IMPORTANT]
> MacOS Catalina の導入により、 [Apple で廃止](https://support.apple.com/HT210200) された ITunes を Mac に、サイドロードアプリを **Finder**にするために必要な統合機能を使用しています。

1. 同期ケーブルを使用し、iPad をコンピューターに接続します。 初めて iPad をコンピューターに接続している場合は、 **このコンピューターを信頼するかどうか**を確認するメッセージが表示されます。 **[信頼する]** を選択して続行します。 また、これが新しい iPad であるかどうか、または1つを復元しているかどうかを尋ねられる場合もあります。

2. [Finder] の [ **場所**] で、メニューバーの下にある [ **iPad** ] アイコンを選択します。

3. ファインダーウィンドウの上部で、[ **ファイル**] をクリックし、[ **Excel** ] または [ **Word**] を見つけます。

4. 別のファインダーウィンドウから、最初のファインダーウィンドウで、サイドロードするアドインの manifest.xml ファイルを **Excel** または **Word** ファイルにドラッグアンドドロップします。

5. iPad で Excel または Word アプリを開きます。 Excel または Word アプリが既に実行されている場合は、[ **ホーム** ] ボタンを選択し、アプリを閉じて再起動します。

6. ドキュメントを開きます。

7. [**挿入**] タブで [**アドイン**] を選択します。 ([**挿入**] タブで、[**アドイン**] ボタンが表示されるまで、横にスクロールする必要がある場合があります)。サイドロードアドインは **、アドインの UI の**[**開発者**] 見出しの下に挿入できます。

    ![Excel アプリでアドインを挿入](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-in-office-on-mac"></a>Office on Mac にアドインをサイドロードする

> [!NOTE]
> Mac に Outlook アドインをサイドロードするには、「[テストのために Outlook アドインをサイドロードする](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-an-add-in-in-outlook-on-the-desktop)」をご参照ください。

1. **ターミナル**を開き、次のいずれかのフォルダーに移動して、アドインのマニフェストファイルを保存します。 `wef` フォルダーがコンピューター上に存在しない場合は、作成します。

    - Word の場合: `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`
    - Excel の場合: `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`
    - PowerPoint の場合: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`

2. コマンド**Finder** `open .` (ピリオドまたはドットを含む) を使用して、Finder でフォルダーを開きます。 アドインのマニフェスト ファイルをこのフォルダーにコピーします。

    ![Office on Mac の Wef フォルダー](../images/all-my-files.png)

3. Word を起動し、ドキュメントを開きます。既に起動している場合は、Word を再起動します。

4. Word で、[アドインの**挿入**  >  **Add-ins**  >  **My Add-ins** ] (ドロップダウンメニュー) を選択し、アドインを選択します。

    ![Office on Mac の個人用アドイン](../images/my-add-ins-wikipedia.png)

    > [!IMPORTANT]
    > サイドロードしたアドインは [個人用アドイン] ダイアログには表示されません。ドロップダウン メニュー内にのみ表示されます (**[挿入]** タブの [個人用アドイン] の右にある小さい下向き矢印)。サイドロードしたアドインは、このメニューの見出し **[開発者向けアドイン]** の下に一覧表示されます。

5. アドインが Word に表示されることを確認します。

    ![Office on Mac に表示された Office アドイン](../images/lorem-ipsum-wikipedia.png)

## <a name="remove-a-sideloaded-add-in"></a>サイドロードアドインを削除する

コンピューター上の Office キャッシュをクリアすることによって、以前のサイドロードアドインを削除することができます。 各プラットフォームとアプリケーションのキャッシュをクリアする方法については、記事「 [Office キャッシュをクリア](clear-cache.md)する」を参照してください。

## <a name="see-also"></a>関連項目

- [iPad と Mac で Office アドインをデバッグする](debug-office-add-ins-on-ipad-and-mac.md)
