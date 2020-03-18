---
title: Office on the web でアドインをデバッグする
description: Office on the web を使用してアドインをテストおよびデバッグする方法。
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: c840d0a16e2a4cdf0bb9f4b213099cb74c2aa815
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719813"
---
# <a name="debug-add-ins-in-office-on-the-web"></a>Office on the web でアドインをデバッグする


Windows、Office 2013、または Office 2016 デスクトップ クライアントを実行していないコンピューター (たとえば、Mac で開発を行っている場合) でアドインの作成とデバッグを行えます。この記事では、Office Online を使用してアドインのテストとデバッグを行う方法について説明します。 この記事では、Office on the web を使用してアドインをテストおよびデバッグする方法について説明します。 

## <a name="prerequisites"></a>前提条件

開始するには

- Office 365 の開発者アカウントをまだお持ちでない場合はこれを取得します。または SharePoint サイトにアクセスできるようにします。

  > [!NOTE]
  > 90 日間の更新可能な無料の Office 365 開発者サブスクリプションを取得するには、[Office 365 開発者プログラム](https://developer.microsoft.com/office/dev-program)にご参加ください。Office 365 開発者プログラムに参加し、サブスクリプションを構成する方法についての詳しい手順については、[Office 365 開発者プログラムのドキュメント](/office/developer-program/office-365-developer-program)を参照してください。

- Office 365 (SharePoint Online) 上でアプリ カタログをセットアップします。アプリ カタログとは、Office アドイン用のドキュメント ライブラリをホストする SharePoint Online の専用サイト コレクションです。独自の SharePoint サイトを所有している場合は、アプリ カタログのドキュメント ライブラリをセットアップできます。詳細については、「[作業ウィンドウ アドインとコンテンツ アドインを SharePoint のアプリ カタログに発行する](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)」を参照してください。


## <a name="debug-your-add-in-from-excel-or-word-on-the-web"></a>Excel または Word on the web からアドインをデバッグする

Word on the web を使用してアドインをデバッグするには: 

1. SSL をサポートするサーバーにアドインを展開します。

    > [!NOTE]
    > [Yeoman ジェネレーター](https://github.com/OfficeDev/generator-office)を使用して、アドインを作成し、ホストすることをお勧めします。

2. [アドイン マニフェスト ファイル](../develop/add-in-manifests.md)で、相対 URI ではなく絶対 URI を含めるように **SourceLocation** 要素の値を更新します。たとえば次のようにします。

    ```xml
    <SourceLocation DefaultValue="https://localhost:44300/App/Home/Home.html" />
    ```

3. SharePoint のアプリ カタログにある Office アドイン ライブラリにマニフェストをアップロードします。

4. Office 365 のアプリ起動ツールから Excel または Word on the web を起動して、新しいドキュメントを開きます。

5. [挿入] タブで、**[個人用アドイン]** または **[Office アドイン]** をクリックし、アプリにアドインを挿入してテストします。

6. お気に入りのブラウザーのツール デバッガーを使用してアドインをデバッグします。

## <a name="potential-issues"></a>潜在的な問題

以下は、デバッグ時に発生する可能性がある問題です。

- 表示される JavaScript エラーのいくつかは Office on the web に起因している可能性があります。

- ブラウザーに無効な証明書エラーが表示されることがありますが、このエラーはバイパスする必要があります。 これを行うプロセスは、ブラウザおよびこの変更を定期的に行うさまざまなブラウザの UI によって異なります。 詳細については、ブラウザーのヘルプを検索するか、オンラインで検索してください。 (たとえば、「Microsoft Edge の無効な証明書警告」を検索します。) ほとんどのブラウザーには、警告ページにリンクがあり、このリンクをクリックするとアドイン ページにアクセスされます。 たとえば、Microsoft Edge には「Web ページへ移動 (推奨しません)」というリンクがあります。 ただし、通常はアドインが再び読み込まれるたびに、このリンクを経由する必要があります。 継続的なバイパスについては、お勧めのヘルプを参照してください。

- コードにブレークポイントを設定すると、保存できないというエラーが Office on the web からスローされることがあります。

## <a name="see-also"></a>関連項目

- [Office アドイン開発のベスト プラクティス](../concepts/add-in-development-best-practices.md)
- [AppSource の検証ポリシー](/office/dev/store/validation-policies)  
- [効率的な AppSource アプリおよびアドインを作成する](/office/dev/store/create-effective-office-store-listings)  
- [Office アドインでのユーザー エラーのトラブルシューティング](testing-and-troubleshooting.md)
    
