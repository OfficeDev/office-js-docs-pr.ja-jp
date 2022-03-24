---
title: テスト用に iPad と Mac で Office アドインをサイドロードする
description: サイドローディングOfficeして、iPad Mac でアドインをテストします。
ms.date: 09/24/2021
ms.localizationpriority: medium
ms.openlocfilehash: b57b072df1fa7c55e709f4ed4045cece8b95aa7e
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746612"
---
# <a name="sideload-office-add-ins-on-ipad-and-mac-for-testing"></a>テスト用に iPad と Mac で Office アドインをサイドロードする

Office on iOS でアドインの実行状態を確認するには、iTunes を利用してアドインのマニフェストを iPad にサイドロードするか、Office on Mac でアドインのマニフェストを直接サイドロードします。このアクションでは、実行中にブレークポイントを設定したり、アドインのコードをデバッグしたりできませんが、その動作を確認したり、UI が使いやすいかどうかや、適切にレンダリングされているかどうかを確認できます。

> [!NOTE]
> Outlook アドインをサイドロードするには、「[テストのために Outlook アドインをサイドロードする](../outlook/sideload-outlook-add-ins-for-testing.md)」をご参照ください。

## <a name="prerequisites-for-office-on-ios"></a>Office on iOS の前提条件

- [iTunes](https://www.apple.com/itunes/download/) がインストールされた Windows または Mac コンピューター。
  > [!IMPORTANT]
  > macOS Catalina を実行している場合、[iTunes](https://support.apple.com/HT210200) は使用できなくなったので、この記事の「Excel または word on [iPad で macOS Catalina](#sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina) を使用してアドインをサイドロードする」セクションの指示に従ってください。

- iOS iPad 8.2 以降を実行し、Excel [Word](https://apps.apple.com/app/microsoft-word/id586447913) がインストールされている[](https://apps.apple.com/app/microsoft-excel/id586683407)場合、同期ケーブルを使用します。

- テスト対象アドインのマニフェスト .xml ファイル。

## <a name="prerequisites-for-office-on-mac"></a>Office on Mac の前提条件

- [Office on Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) がインストールされていて OS X v10.10 "Yosemite" を実行している Mac。

- Word on Mac バージョン 15.18 (160109)。

- Excel on Mac バージョン 15.19 (160206)。

- PowerPoint on Mac バージョン 15.24 (160614)

- テスト対象アドインのマニフェスト .xml ファイル。

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad-using-itunes"></a>iTunes を使用して、Excelまたは Word のアドインiPad読み込む

1. 同期ケーブルを使用し、iPad をコンピューターに接続します。 コンピューターに初めてiPad接続する場合は、このコンピューターの信頼を求める **メッセージが表示されます**。 **[信頼する]** を選択して続行します。

2. iTunes で、メニュー バーの下にある **[iPad]** のアイコンをクリックします。

3. iTunes の左側の **[設定]** で、**[App]** をクリックします。

4. iTunes の右側で、**[ファイル共有]** までスクロールしてから、**[アドイン]** 列で **[Excel]** または **[Word]** をクリックします。

5. [Excel または **Word ドキュメント**] **列の** 下部にある [ファイルの追加] を選択し、サイドロードするアドインのマニフェスト .xml ファイルを選択します。

6. iPad で Excel または Word アプリを開きます。 アプリまたは word Excelが既に実行されている場合は、[ホーム] ボタンを選択し、アプリを閉じて再起動します。

7. ドキュメントを開きます。

8. [**挿入] タブで** [アドイン] を選択します([挿入] タブで、[アドイン] ボタンが表示されるまで水平方向にスクロール **する必要** がある場合があります)。サイドロードされたアドインは、アドイン UI の **[開発者**] 見出しの下 **に挿入** できます。

    ![アドインをアプリに挿入Excelします。](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina"></a>macOS Catalina を使用して、Excelまたは Word のiPadをサイドロードする

> [!IMPORTANT]
> macOS Catalina の導入により、Apple は Mac 上 [の iTunes](https://support.apple.com/HT210200) を廃止し、アプリを Finder にサイドロードするために必要な統合された **機能を廃止しました**。

1. 同期ケーブルを使用し、iPad をコンピューターに接続します。 コンピューターに初めてiPad接続する場合は、このコンピューターの信頼を求める **メッセージが表示されます**。 **[信頼する]** を選択して続行します。 また、これが新しいファイルか、iPad復元中かという問い合いを受け取る場合があります。

2. Finder の [場所] **で**、メニュー バー **iPadアイコンを** 選択します。

3. [Finder] ウィンドウの上部にある [ファイル] を **クリックし、[** ファイル] または [Word] **Excel** 探 **します**。

4. 別の Finder ウィンドウから、サイド ロードするアドインの manifest.xml ファイルを最初の **Finder** ウィンドウの Excel または **Word** ファイルにドラッグ アンド ドロップします。

5. iPad で Excel または Word アプリを開きます。 アプリまたは word Excelが既に実行されている場合は、[ホーム] ボタンを選択し、アプリを閉じて再起動します。

6. ドキュメントを開きます。

7. [**挿入] タブで** [アドイン] を選択します([挿入] タブで、[アドイン] ボタンが表示されるまで水平方向にスクロール **する必要** がある場合があります)。サイドロードされたアドインは、アドイン UI の **[開発者**] 見出しの下 **に挿入** できます。

    ![アドインをアプリに挿入Excelします。](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-in-office-on-mac"></a>Office on Mac にアドインをサイドロードする

1. ターミナル **を** 開き、アドインのマニフェスト ファイルを保存する次のいずれかのフォルダーに移動します。 `wef` フォルダーがコンピューター上に存在しない場合は、作成します。

    - Word の場合: `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`
    - Excel の場合: `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`
    - PowerPoint の場合: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`

2. コマンド (ピリオドまたはドットを含む) を使用して `open .` **Finder** でフォルダーを開きます。 アドインのマニフェスト ファイルをこのフォルダーにコピーします。

    ![Mac 上の Office Wef フォルダー。](../images/all-my-files.png)

3. Word を起動し、ドキュメントを開きます。既に起動している場合は、Word を再起動します。

4. Word で、[**InsertAdd-insMy** >  >  アドイン] (ドロップダウン メニュー) を選択し、アドインを選択します。

    ![Mac 上の [Officeアドイン]](../images/my-add-ins-wikipedia.png)

    > [!IMPORTANT]
    > サイドロードしたアドインは [個人用アドイン] ダイアログには表示されません。ドロップダウン メニュー内にのみ表示されます (**[挿入]** タブの [個人用アドイン] の右にある小さい下向き矢印)。サイドロードしたアドインは、このメニューの見出し **[開発者向けアドイン]** の下に一覧表示されます。

5. アドインが Word に表示されることを確認します。

    ![Officeに表示されるアドインOffice Mac 上に表示されます。](../images/lorem-ipsum-wikipedia.png)

## <a name="remove-a-sideloaded-add-in"></a>サイドロードされたアドインを削除する

以前にサイドロードされたアドインを削除するには、コンピューター上Officeキャッシュをクリアします。 各プラットフォームとアプリケーションのキャッシュをクリアする方法の詳細については、「キャッシュのクリア」[をOfficeしてください](clear-cache.md)。

## <a name="see-also"></a>関連項目

- [Mac で Office アドインをデバッグする](debug-office-add-ins-on-ipad-and-mac.md)
- [テスト用に Outlook アドインをサイドロードする](../outlook/sideload-outlook-add-ins-for-testing.md)
