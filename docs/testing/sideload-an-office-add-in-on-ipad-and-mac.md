---
title: テスト用に iPad と Mac で Office アドインをサイドロードする
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: e5ec6924917f2351da77c8b9a84eb8de77b3864e
ms.sourcegitcommit: fdf7f4d686700edd6e6b04b2ea1bd43e59d4a03a
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/28/2018
ms.locfileid: "25348129"
---
# <a name="sideload-office-add-ins-on-ipad-and-mac-for-testing"></a>テスト用に iPad と Mac で Office アドインをサイドロードする

Office for iOS でのアドインの実行状況を確認するには、iTunes を 使用してアドインのマニフェストを iPad にサイドロードするか、アドインのマニフェストを Office for Mac で直接サイドロードします。このアクションでは、ブレークポイントを設定したり、アドインのコードを実行中にデバッグすることはできませんが、その動作を確認したり、UI が使用可能であり適切にレンダリングされることを確認できます。 

## <a name="prerequisites-for-office-for-ios"></a>Office for iOS の前提条件

- [iTunes](http://www.apple.com/itunes/download/) がインストールされた Windows または Mac コンピュータ。
    
- [Excel for iPad](https://itunes.apple.com/us/app/microsoft-excel/id586683407?mt=8) がインストールされた iOS 8.2 以上の iPad と同期ケーブル。
    
- テスト対象アドインのマニフェスト .xml ファイル。
    

## <a name="prerequisites-for-office-for-mac"></a>Office for Mac の前提条件

- [Office for Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) をインストールし、OS X v10.10 "Yosemite" 以降で動作している Mac。
    
- Word for Mac バージョン 15.18 (160109)。
   
- Excel for Mac バージョン 15.19 (160206)。

- PowerPoint for Mac バージョン 15.24 (160614)
    
- テスト対象アドインのマニフェスト .xml ファイル。
    

## <a name="sideload-an-add-in-on-excel-or-word-for-ipad"></a>Excel for iPad または Word for iPad のアドインをサイドロードする

1. 同期ケーブルを使用し、iPad をコンピューターに接続します。iPad を初めてコンピューターに接続する場合、**[このコンピューターを信頼しますか?]** というメッセージが表示されるので、**[信頼する]** を選択して続行します。

2. iTunes のメニュー バーの下にある **[iPad]** アイコンをクリックします。
    
    ![iTunes の iPad アイコン](../images/ipad.png)

3. iTunes の左側にある  **[設定]** で、 **[App]** をクリックします。
    
    ![iTunes App 設定](../images/file-settings-apps.png)

4. iTunes の右側で、 **[ファイル共有]** までスクロールしてから、 **[アドイン]** 列で **[Excel]** または **[Word]** をクリックします。
    
    ![iTunes ファイル共有](../images/file-sharing.png)

5. **[Excel]** 列または **[Word ドキュメント]** 列の下部で、 **[ファイルの追加]** をクリックしてから、サイドロードするアドインのマニフェスト .xml ファイルを選択します。 
    
6. iPad で Excel または Word アプリを開きます。Excel または Word アプリがすでに実行されている場合は、 **[ホーム]** ボタンを選択し、アプリを閉じてから再起動します。
    
7. ドキュメントを開きます。
    
8. **[挿入]** タブで **[アドイン]** をクリックします。 **[アドイン]** UI の **[開発者]** という見出しの下に、サイドロードしたアドインが表示され、挿入するために選択できます。
    
    ![Excel アプリでアドインを挿入](../images/excel-insert-add-in.png)


## <a name="sideload-an-add-in-on-office-for-mac"></a>Office for Mac でアドインをサイドロードする

> [!NOTE]
> Outlook for Mac アドインをサイドロードするには、「[テスト用に Outlook アドインをサイドロードする](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing)」をご参照ください。

1. **ターミナル**を開き、アプリに対応するフォルダーに移動しアドインのマニフェスト ファイルを保存します。`wef` フォルダーがコンピューターにない場合には作成します。
    
    - Word の場合:  `/Users/<username>/Library/Containers/com.microsoft.Word/Data/documents/wef`    
    - Excel の場合:  `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/documents/wef`
    - PowerPoint の場合: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/documents/wef`
    
2. **Finder** で `open .` コマンド (ピリオドまたはドットを含む) を使用してフォルダーを開き、アドインのマニフェスト ファイルをこのフォルダーにコピーします。
    
    ![Office for Mac の Wef フォルダー](../images/all-my-files.png)

3. Word を起動し、ドキュメントを開きます。既に起動している場合は、Word を再起動します。
    
4. Word で、**[挿入]** > **[アドイン]** > **[個人用アドイン]** (ドロップダウン メニュー) を選択し、アドインを選択します。
    
    ![Office for Mac のマイ アドイン](../images/my-add-ins-wikipedia.png)

    > [!IMPORTANT]
    > サイドロードしたアドインは [個人用アドイン] ダイアログには表示されず、ドロップダウン メニュー (**[挿入]** タブの [個人用アドイン] の右にある小さな下向き矢印) 内にのみ表示されます。サイドロードしたアドインは、このメニューの見出し **[開発者向けアドイン]** の下に一覧表示されます。 
    
5. Word にアドインが表示されることを確認します。
    
    ![Office for Mac で示される Office アドイン](../images/lorem-ipsum-wikipedia.png)
    
    > [!NOTE]
    > Office for Mac は、パフォーマンス上の理由からアドインを頻繁にキャッシュします。アドイン開発中に強制的にリロードする必要がある場合は、`Users/<usr>/Library/Containers/com.Microsoft.OsfWebHost/Data/` フォルダーを消去します。このフォルダーが存在しない場合は、`com.microsoft.Office365ServiceV2/Data/Caches/com.microsoft.Office365ServiceV2/` フォルダーのファイルを消去します。

## <a name="see-also"></a>関連項目

- [iPad と Mac で Office アドインをデバッグする](debug-office-add-ins-on-ipad-and-mac.md)
    
