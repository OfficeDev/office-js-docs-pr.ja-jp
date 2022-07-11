---
title: iPad でテスト用の Office アドインをサイドロードする
description: サイドローディングによって iPad で Office アドインをテストします。
ms.date: 06/29/2022
ms.localizationpriority: medium
ms.openlocfilehash: 0ba52ae78bed36c4eb8130c714577a1b0899aeb6
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/11/2022
ms.locfileid: "66713208"
---
# <a name="sideload-office-add-ins-on-ipad-for-testing"></a>iPad でテスト用の Office アドインをサイドロードする

iOS 上の Office でアドインがどのように実行されるかを確認するには、iTunes を使用してアドインのマニフェストを iPad にサイドロードします。 このアクションでは、実行中、ブレークポイントを設定したり、アドインのコードをデバッグしたりできませんが、その動作を確認したり、UI が使えることと適切にレンダリングされることを確認できます。

> [!NOTE]
> Outlook アドインをサイドロードするには、「[テストのために Outlook アドインをサイドロードする](../outlook/sideload-outlook-add-ins-for-testing.md)」をご参照ください。

## <a name="prerequisites-for-office-on-ios"></a>Office on iOS の前提条件

- [iTunes](https://www.apple.com/itunes/download/) がインストールされた Windows または Mac コンピューター。
  > [!IMPORTANT]
  > macOS Catalina を実行している場合、 [iTunes は使用できなくなりました](https://support.apple.com/HT210200) ので、この記事の後半の「 [macOS Catalina を使用して Excel または Word on iPad でアドインをサイドロード](#sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina) する」セクションの手順に従う必要があります。

- [Excel](https://apps.apple.com/app/microsoft-excel/id586683407) または [Word が](https://apps.apple.com/app/microsoft-word/id586447913)インストールされた iOS 8.2 以降を実行する iPad と同期ケーブル。

- テスト対象アドインのマニフェスト .xml ファイル。

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad-using-itunes"></a>iTunes を使用して Excel または Word on iPad でアドインをサイドロードする

1. 同期ケーブルを使用し、iPad をコンピューターに接続します。 iPad を初めてコンピューターに接続する場合は、このコンピューターの **信頼** を求めるメッセージが表示されます。 **[信頼する]** を選択して続行します。

2. iTunes で、メニュー バーの下にある **[iPad]** のアイコンをクリックします。

3. iTunes の左側の **[設定]** で、**[App]** をクリックします。

4. iTunes の右側で、**[ファイル共有]** までスクロールしてから、**[アドイン]** 列で **[Excel]** または **[Word]** をクリックします。

5. **Excel** または **Word Documents** 列の下部にある **[ファイルの追加**] を選択し、サイドロードするアドインのマニフェスト .xml ファイルを選択します。

6. iPad で Excel または Word アプリを開きます。 Excel または Word アプリが既に実行されている場合は、[ **ホーム** ] ボタンを選択し、アプリを閉じて再起動します。

7. ドキュメントを開きます。

8. [**挿入**] タブ **で [アドイン**] を選択します([**挿入**] タブで、[**アドイン**] ボタンが表示されるまで水平方向にスクロールする必要がある場合があります)。サイドロードされたアドインは、アドイン UI の **[開発者**] 見出しの下 **に** 挿入できます。

    ![Excel アプリにアドインを挿入します。](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina"></a>macOS Catalina を使用して Excel または Word on iPad でアドインをサイドロードする

> [!IMPORTANT]
> macOS Catalina の導入により、 [Apple は Mac 上の iTunes を廃止](https://support.apple.com/HT210200) し、 **Finder** にアプリをサイドローディングするために必要な機能を統合しました。

1. 同期ケーブルを使用し、iPad をコンピューターに接続します。 iPad を初めてコンピューターに接続する場合は、このコンピューターの **信頼** を求めるメッセージが表示されます。 **[信頼する]** を選択して続行します。 また、これが新しい iPad かどうか、または復元しようとしているかどうかも尋ねられる場合があります。

2. Finder の [ **場所]** で、メニュー バーの下にある **iPad** アイコンを選択します。

3. Finder ウィンドウの上部にある [ **ファイル**] をクリックし、 **Excel** または **Word** を探します。

4. 別の Finder ウィンドウで、サイド ロードするアドインのmanifest.xml ファイルを最初の Finder ウィンドウの **Excel** または **Word** ファイルにドラッグ アンド ドロップします。

5. iPad で Excel または Word アプリを開きます。 Excel または Word アプリが既に実行されている場合は、[ **ホーム** ] ボタンを選択し、アプリを閉じて再起動します。

6. ドキュメントを開きます。

7. [**挿入**] タブ **で [アドイン**] を選択します([**挿入**] タブで、[**アドイン**] ボタンが表示されるまで水平方向にスクロールする必要がある場合があります)。サイドロードされたアドインは、アドイン UI の **[開発者**] 見出しの下 **に** 挿入できます。

    ![Excel アプリにアドインを挿入します。](../images/excel-insert-add-in.png)

## <a name="remove-a-sideloaded-add-in"></a>サイドロードされたアドインを削除する

以前にサイドロードされたアドインを削除するには、コンピューター上の Office キャッシュをクリアします。 プラットフォームとアプリケーションごとにキャッシュをクリアする方法の詳細については、「 [Office キャッシュをクリアする」](clear-cache.md)を参照してください。

## <a name="see-also"></a>関連項目

- [テスト用の Mac 上の Office アドインをサイドロードする](sideload-an-office-add-in-on-mac.md)
- [Mac で Office アドインをデバッグする](debug-office-add-ins-on-ipad-and-mac.md)
- [テスト用に Outlook アドインをサイドロードする](../outlook/sideload-outlook-add-ins-for-testing.md)
