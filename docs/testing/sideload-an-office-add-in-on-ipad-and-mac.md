---
title: テスト用に iPad と Mac で Office アドインをサイドロードする
description: サイドロードOffice iPad と Mac でアドインをテストします。
ms.date: 09/02/2020
localization_priority: Normal
ms.openlocfilehash: 22271409cdacd8f3e32039743b8916b1fb87252f
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/14/2021
ms.locfileid: "50238072"
---
# <a name="sideload-office-add-ins-on-ipad-and-mac-for-testing"></a>テスト用に iPad と Mac で Office アドインをサイドロードする

Office on iOS でアドインの実行状態を確認するには、iTunes を利用してアドインのマニフェストを iPad にサイドロードするか、Office on Mac でアドインのマニフェストを直接サイドロードします。このアクションでは、実行中にブレークポイントを設定したり、アドインのコードをデバッグしたりできませんが、その動作を確認したり、UI が使いやすいかどうかや、適切にレンダリングされているかどうかを確認できます。

## <a name="prerequisites-for-office-on-ios"></a>Office on iOS の前提条件

- [iTunes](https://www.apple.com/itunes/download/) がインストールされた Windows または Mac コンピューター。
  > [!IMPORTANT]
  > macOS Catalina を実行している場合 [、iTunes](https://support.apple.com/HT210200) は使用できなくなったので、この記事で後に説明する [「macOS Catalina](#sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina) を使用して、Excel または iPad 上の Word でアドインをサイドロードする」の手順に従う必要があります。

- [Excel](https://apps.apple.com/app/microsoft-excel/id586683407)または[Word](https://apps.apple.com/app/microsoft-word/id586447913)がインストールされた iOS 8.2 以降を実行している iPad と同期ケーブル。

- テスト対象アドインのマニフェスト .xml ファイル。

## <a name="prerequisites-for-office-on-mac"></a>Office on Mac の前提条件

- [Office on Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) がインストールされていて OS X v10.10 "Yosemite" を実行している Mac。

- Word on Mac バージョン 15.18 (160109)。

- Excel on Mac バージョン 15.19 (160206)。

- PowerPoint on Mac バージョン 15.24 (160614)

- テスト対象アドインのマニフェスト .xml ファイル。

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad-using-itunes"></a>iTunes を使用して Excel または iPad 上の Word でアドインをサイドロードする

1. 同期ケーブルを使用し、iPad をコンピューターに接続します。 If you're connecting the iPad to your computer for the first time, you'll be prompted with **Trust This Computer?**. **[信頼する]** を選択して続行します。

2. iTunes で、メニュー バーの下にある **[iPad]** のアイコンをクリックします。

3. iTunes の左側の **[設定]** で、**[App]** をクリックします。

4. iTunes の右側で、**[ファイル共有]** までスクロールしてから、**[アドイン]** 列で **[Excel]** または **[Word]** をクリックします。

5. [Excel またはWord **ドキュメント**] 列の下部で、[ファイルの追加] を選択し、サイドロードするアドインのマニフェスト .xml ファイルを選択します。

6. iPad で Excel または Word アプリを開きます。 Excel または Word アプリが既に実行されている場合は、[ホーム] ボタンを選択し、アプリを閉じて再起動します。

7. ドキュメントを開きます。

8. [**挿入] タブ** で[アドイン] を選択します ([挿入] タブでは、[アドイン] ボタンが表示されるまで水平方向にスクロールする必要がある **場合** があります)。サイドロードしたアドインは、アドイン UI の **[** 開発者] 見出し **の下に挿入** できます。

    ![Excel アプリでアドインを挿入](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina"></a>macOS Catalina を使用して Excel または iPad 上の Word でアドインをサイドロードする

> [!IMPORTANT]
> macOS Catalina の導入により [、Apple](https://support.apple.com/HT210200) は Mac で iTunes を廃止し、Finder にアプリをサイドロードするために必要な統合された機能 **を提供しました**。

1. 同期ケーブルを使用し、iPad をコンピューターに接続します。 If you're connecting the iPad to your computer for the first time, you'll be prompted with **Trust This Computer?**. **[信頼する]** を選択して続行します。 また、これが新しい iPad か、復元中か尋ねらる場合があります。

2. Finder の [場所] **で**、メニュー バーの **下にある iPad** アイコンを選択します。

3. [Finder] ウィンドウの上部にある [ファイル] をクリックし **、Excel** または Word を探 **します**。

4. 別の [Finder] ウィンドウで、サイド ロードするアドインの manifest.xml ファイルを最初の **Finder** ウィンドウの Excel または **Word** ファイルにドラッグ アンド ドロップします。

5. iPad で Excel または Word アプリを開きます。 Excel または Word アプリが既に実行されている場合は、[ホーム] ボタンを選択し、アプリを閉じて再起動します。

6. ドキュメントを開きます。

7. [**挿入] タブ** で[アドイン] を選択します ([挿入] タブでは、[アドイン] ボタンが表示されるまで水平方向にスクロールする必要がある **場合** があります)。サイドロードしたアドインは、アドイン UI の **[** 開発者] 見出し **の下に挿入** できます。

    ![Excel アプリでアドインを挿入](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-in-office-on-mac"></a>Office on Mac にアドインをサイドロードする

> [!NOTE]
> Mac に Outlook アドインをサイドロードするには、「[テストのために Outlook アドインをサイドロードする](../outlook/sideload-outlook-add-ins-for-testing.md)」をご参照ください。

1. ターミナル **を** 開き、アドインのマニフェスト ファイルを保存する次のいずれかのフォルダーに移動します。 `wef` フォルダーがコンピューター上に存在しない場合は、作成します。

    - Word の場合: `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`
    - Excel の場合: `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`
    - PowerPoint の場合: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`

2. Finder でコマンド ( **ピリオド** またはドットを含 `open .` む) を使用してフォルダーを開きます。 アドインのマニフェスト ファイルをこのフォルダーにコピーします。

    ![Office on Mac の Wef フォルダー](../images/all-my-files.png)

3. Word を起動し、ドキュメントを開きます。既に起動している場合は、Word を再起動します。

4. Word で、[アドインの挿入] (ドロップダウン メニュー) を選択し、アドイン  >    >  を選択します。

    ![Office on Mac の個人用アドイン](../images/my-add-ins-wikipedia.png)

    > [!IMPORTANT]
    > サイドロードしたアドインは [個人用アドイン] ダイアログには表示されません。ドロップダウン メニュー内にのみ表示されます (**[挿入]** タブの [個人用アドイン] の右にある小さい下向き矢印)。サイドロードしたアドインは、このメニューの見出し **[開発者向けアドイン]** の下に一覧表示されます。

5. アドインが Word に表示されることを確認します。

    ![Office on Mac に表示された Office アドイン](../images/lorem-ipsum-wikipedia.png)

## <a name="remove-a-sideloaded-add-in"></a>サイドロードされたアドインを削除する

以前にサイドロードしたアドインを削除するには、コンピューター上のOfficeキャッシュをクリアします。 各プラットフォームとアプリケーションのキャッシュをクリアする方法の詳細については、「キャッシュのクリア」Office [覧ください](clear-cache.md)。

## <a name="see-also"></a>関連項目

- [iPad と Mac で Office アドインをデバッグする](debug-office-add-ins-on-ipad-and-mac.md)
