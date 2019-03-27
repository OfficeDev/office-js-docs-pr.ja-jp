---
title: テスト用に iPad と Mac で Office アドインをサイドロードする
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: e2f9ee912395e0f54130f0e78109cab4479b6567
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871977"
---
# <a name="sideload-office-add-ins-on-ipad-and-mac-for-testing"></a>テスト用に iPad と Mac で Office アドインをサイドロードする

Office for iOS でアドインを実行するしくみを確認するには、iTunes を利用し、アドインのマニフェストを iPad にサイドロードするか、Office for Mac で直接、アドインのマニフェストをサイドロードします。このアクションでは、実行中、ブレークポイントを設定したり、アドインのコードをデバッグしたりできませんが、その動作を確認したり、UI が使えることと適切にレンダリングされることを確認できます。 

## <a name="prerequisites-for-office-for-ios"></a>Office for iOS の前提条件

- [iTunes](https://www.apple.com/itunes/download/) がインストールされた Windows または Mac コンピューター。
    
- [iPad 用 Excel](https://itunes.apple.com/us/app/microsoft-excel/id586683407?mt=8) がインストールされた iOS 8.2 以上の iPad と同期ケーブル。
    
- テスト対象アドインのマニフェスト .xml ファイル。
    

## <a name="prerequisites-for-office-for-mac"></a>Office for Mac の前提条件

- OS X v10.10 "Yosemite" 以降が動作し、 [Office for Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) がインストールされている Mac。
    
- Word for Mac バージョン 15.18 (160109)。
   
- Excel for Mac バージョン 15.19 (160206)。

- PowerPoint for Mac バージョン 15.24 (160614)
    
- テスト対象アドインのマニフェスト .xml ファイル。
    

## <a name="sideload-an-add-in-on-excel-or-word-for-ipad"></a>iPad 用 Excel または Word のアドインをサイドロードする

1. 同期ケーブルを使用し、iPad をコンピューターに接続します。iPad を初めてコンピューターに接続する場合、**[このコンピューターを信頼しますか?]** と問われます。**[信頼する]** を選択して続行します。

2. iTunes で、メニュー バーの下にある **[iPad]** のアイコンをクリックします。

3. iTunes の左側の  **[設定]** で、 **[App]** をクリックします。

4. iTunes の右側で、 **[ファイル共有]** までスクロールしてから、 **[アドイン]** 列で **[Excel]** または **[Word]** をクリックします。

5. 
            **[Excel]** 列または **[Word ドキュメント]** 列の下部で、 **[ファイルの追加]** をクリックしてから、サイドロードするアドインのマニフェスト .xml ファイルを選択します。 
    
6. iPad で Excel または Word アプリを開きます。Excel または Word アプリがすでに実行されている場合は、 **[ホーム]** ボタンを選択して、アプリを閉じて再起動します。
    
7. ドキュメントを開きます。
    
8. **[挿入]** タブで **[アドイン]** をクリックします。 **[アドイン]** UI の **[開発者]** という見出しの下に、サイドロードしたアドインが表示され、挿入のために選択できるようになっています。
    
    ![Excel アプリでアドインを挿入](../images/excel-insert-add-in.png)


## <a name="sideload-an-add-in-on-office-for-mac"></a>Office for Mac でアドインをサイドロードする

> [!NOTE]
> Outlook for Mac アドインをサイドロードするには、「[テスト用に Outlook アドインをサイドロードする](/outlook/add-ins/sideload-outlook-add-ins-for-testing)」を参照してください。

1. **Terminal** を開き、次のフォルダーの 1 つに移動します。そこにアドインのマニフェスト ファイルを保存します。`wef` フォルダーがコンピューターにない場合、作成します。
    
    - Word の場合: `/Users/<username>/Library/Containers/com.microsoft.Word/Data/documents/wef`    
    - Excel の場合: `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/documents/wef`
    - PowerPoint の場合: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/documents/wef`
    
2. **Finder** で `open .` コマンドを使用してフォルダーを開きます (ピリオドまたはドットを含みます)。アドインのマニフェスト ファイルをこのフォルダーにコピーします。
    
    ![Office for Mac の Wef フォルダー](../images/all-my-files.png)

3. Word を起動し、ドキュメントを開きます。既に起動している場合は、Word を再起動します。
    
4. Word で、**[挿入]** > **[アドイン]** > **[個人用アドイン]** (ドロップダウン メニュー) を選択し、アドインを選択します。
    
    ![Office for Mac のマイ アドイン](../images/my-add-ins-wikipedia.png)

    > [!IMPORTANT]
    > サイドロードしたアドインは [個人用アドイン] ダイアログには表示されません。ドロップダウン メニュー内にのみ表示されます (**[挿入]** タブの [個人用アドイン] の右にある小さい下向き矢印)。サイドロードしたアドインは、このメニューの見出し **[開発者向けアドイン]** の下に一覧表示されます。 
    
5. アドインが Word に表示されることを確認します。
    
    ![Office for Mac に表示される Office アドイン](../images/lorem-ipsum-wikipedia.png)
    
    > [!NOTE]
    > Office for Mac では、パフォーマンス上の理由でアドインがよくキャッシュされます。 アドインの開発中にそれを再読み込みする必要がある場合は、`Users/<usr>/Library/Containers/com.Microsoft.OsfWebHost/Data/` フォルダーをクリアにすることが可能です。 フォルダーが存在しない場合、`com.microsoft.Office365ServiceV2/Data/Caches/com.microsoft.Office365ServiceV2/` フォルダーのファイルをクリアします。

## <a name="see-also"></a>関連項目

- [iPad と Mac で Office アドインをデバッグする](debug-office-add-ins-on-ipad-and-mac.md)
