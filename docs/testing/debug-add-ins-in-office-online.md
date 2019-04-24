---
title: Office Online でアドインをデバッグする
description: Office Online を使用してアドインのテストとデバッグを行う方法
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: ff77f3d8b3e332288d4ccb3e2d2305d1b1c4a825
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451529"
---
# <a name="debug-add-ins-in-office-online"></a>Office Online でアドインをデバッグする


Windows、Office 2013、または Office 2016 デスクトップ クライアントを実行していないコンピューター (たとえば、Mac で開発を行っている場合) でアドインの作成とデバッグを行えます。この記事では、Office Online を使用してアドインのテストとデバッグを行う方法について説明します。 この記事では、Office Online を使用してアドインのテストとデバッグを行う方法を説明します。 

## <a name="prerequisites"></a>前提条件

開始するには

- Office 365 の開発者アカウントをまだお持ちでない場合はこれを取得します。または SharePoint サイトにアクセスできるようにします。
    
  > [!NOTE]
  > 無料の Office 365 開発者サブスクリプションにサインアップするには、[Office 365 開発者プログラム](https://developer.microsoft.com/office/dev-program)にご参加ください。 Office 365 開発者プログラムに参加し、サブスクリプションにサインアップして構成する方法についての詳しい手順については、[Office 365 開発者プログラムのドキュメント](/office/developer-program/office-365-developer-program)を参照してください。
     
- Office 365 (SharePoint Online) 上でアドイン カタログをセットアップするアドイン カタログとは、Office アドイン用のドキュメント ライブラリをホストする SharePoint Online の専用サイト コレクションです。独自の SharePoint サイトを所有している場合は、アドイン カタログのドキュメント ライブラリをセットアップすることができます。詳細については、「[作業ウィンドウ アドインとコンテンツ アドインを SharePoint カタログに発行する](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)」をご覧ください。
    

## <a name="debug-your-add-in-from-excel-online-or-word-online"></a>Excel Online または Word Online からアドインをデバッグする

Office Online を使用してアドインをデバッグするには、

1. SSL をサポートするサーバーにアドインを展開します。
    
    > [!NOTE]
    > [Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)を使用して、アドインを作成し、ホストすることをお勧めします。
     
2. [アドイン マニフェスト ファイル](../develop/add-in-manifests.md)で、相対 URI ではなく絶対 URI を含めるように **SourceLocation** 要素の値を更新します。たとえば次のようにします。
      
    ```xml
    <SourceLocation DefaultValue="https://localhost:44300/App/Home/Home.html" />
    ```
    
3. SharePoint のアドイン カタログにある Office アドイン ライブラリにマニフェストをアップロードします。
    
4. Office 365 のアプリ起動ツールから Excel Online または Word Online を起動し、新しいドキュメントを開きます。
    
5. [挿入] タブで、 **[個人用アドイン]** または **[Office アドイン]** をクリックし、アプリにアドインを挿入してテストします。
    
6. お気に入りのブラウザーのツール デバッガーを使用してアドインをデバッグします。

## <a name="potential-issues"></a>潜在的な問題    

以下は、デバッグ時に発生する可能性がある問題です。
    
- 表示される JavaScript エラーのいくつかは Office Online に起因している可能性があります。
      
- ブラウザーが、バイパスが必要になる、無効な証明書エラーを表示することがあります。
      
- コードにブレークポイントを設定する場合、Office Online から、保存できないというエラーがスローされることがあります。

## <a name="see-also"></a>関連項目

- [Office アドイン開発のベスト プラクティス](../concepts/add-in-development-best-practices.md)
- [AppSource の検証ポリシー](/office/dev/store/validation-policies)  
- [効率的な AppSource アプリおよびアドインを作成する](/office/dev/store/create-effective-office-store-listings)  
- [Office アドインでのユーザー エラーのトラブルシューティング](testing-and-troubleshooting.md)
    
