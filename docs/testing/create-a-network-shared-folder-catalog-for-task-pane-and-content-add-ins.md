---
title: テスト用に Office アドインをサイドロードする
description: ''
ms.date: 01/25/2018
ms.openlocfilehash: b143999422866dba9b43432359c12f3607261c60
ms.sourcegitcommit: e094aaa06d9aff3d13f8ffd3429d4a31f0b65b81
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/03/2018
ms.locfileid: "21782813"
---
# <a name="sideload-office-add-ins-for-testing"></a>テスト用に Office アドインをサイドロードする

マニフェストをネットワーク ファイル共有に公開することで、Windows 上で実行されている Office クライアントにテスト用の Office アドインをインストールできます（以下の手順）。

> [!NOTE]
> [**yo office **ツール](https://github.com/OfficeDev/generator-office)を使用してアドイン プロジェクトを作成した場合、お客様に適した別のサイドロードの方法があります。 詳細は、 [sideload コマンドを使用した Sideload Office アドイン](sideload-office-addin-using-sideload-command.md)を参照してください。

この記事は、Windows 上の Word、Excel、または PowerPoint アドインのテストにのみ適用されます。 別のプラットフォームでテストする場合、または Outlook アドインをテストする場合は、次のトピックのいずれかを参照してアドインをサイドロードします。

- [テスト用に Office Online で Office アドインをサイドロードする](sideload-office-add-ins-for-testing.md)
- [テスト用に iPad と Mac で Office アドインをサイドロードする](sideload-an-office-add-in-on-ipad-and-mac.md)
- [テスト用に Outlook アドインをサイドロードする](../../../../outlook/add-ins/sideload-outlook-add-ins-for-testing)


次のビデオでは、共有フォルダ カタログを使用して Office デスクトップまたは Office Online のアドインをサイドロードする手順について説明します。  


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]


## <a name="share-a-folder"></a>フォルダーの共有

1. アドインをホストさせようとしている Windows コンピューターで、共有フォルダー カタログとして使用するつもりのフォルダーの親フォルダーまたはドライブ文字に移動します。

2. フォルダーのコンテキスト メニューを (右クリックして) 開き、**[プロパティ]** を選びます。

3. **[共有]** タブを開きます。

4. 「**相手を選んでください**」ページで、自分自身とアドインを共有する相手を追加します。相手がセキュリティ グループのメンバー全員の場合は、そのグループを追加できます。少なくとも、フォルダーへの**読み取り/書き込み**アクセス許可が必要です。 

5. **[共有]** > **[完了]** > **[閉じる]** の順に選択します。


## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a>信頼できるカタログとしてその共有フォルダーを指定します。
      
1. Excel、Word、または PowerPoint で新しいドキュメントを開きます。
    
2. **[ファイル]** タブを選び、**[オプション]** を選びます。
    
3. **[セキュリティ センター]** を選び、**[セキュリティ センターの設定]** ボタンを選びます。
    
4. **[信頼されているアドイン カタログ]** を選びます。
    
5. **[カタログの URL]** ボックスで、共有フォルダー カタログへの完全なネットワーク パスを入力し、**[カタログの追加]** を選びます。
    
6. **[メニューに表示する]** チェック ボックスをオンにし、**[OK]** を選びます。

7. Office アプリケーションを閉じると変更内容が有効になります。
    

## <a name="sideload-your-add-in"></a>アドインのサイドロード

1. テストするアドインのマニフェスト ファイルを共有フォルダー カタログに置きます。なお、Web サーバーに Web アプリケーション自体を展開します。必ずマニフェスト ファイルの **SourceLocation** 要素で URL を指定してください。

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

2. Excel、Word、または PowerPoint で、リボンの **[挿入]** タブにある **[個人用アドイン]** を選びます。

3. **[Office アドイン]** ダイアログ ボックスの上部にある **[共有フォルダー]** を選びます。

4. アドインの名前を選び、**[OK]** を選択して、アドインを挿入します。


## <a name="see-also"></a>関連項目

- [マニフェストの問題を検証し、トラブルシューティングする](troubleshoot-manifest.md)
- [Office アドインを発行する](../publish/publish.md)
    
