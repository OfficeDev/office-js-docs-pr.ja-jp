---
title: Office on the web でアドインをデバッグする
description: Office on the web を使用してアドインをテストおよびデバッグする方法。
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: f7ef3fa3d6389629e28b428b9bdbe3b128896b1f
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094492"
---
# <a name="debug-add-ins-in-office-on-the-web"></a>Office on the web でアドインをデバッグする

Windows、Office 2013、または Office 2016 デスクトップ クライアントを実行していないコンピューター (たとえば、Mac で開発を行っている場合) でアドインの作成とデバッグを行えます。この記事では、Office Online を使用してアドインのテストとデバッグを行う方法について説明します。 この記事では、Office on the web を使用してアドインをテストおよびデバッグする方法について説明します。 

## <a name="prerequisites"></a>前提条件

開始するには

- Microsoft 365 開発者アカウントを持っていない場合、または SharePoint サイトにアクセスできない場合は、Microsoft 開発者アカウントを取得します。

  > [!NOTE]
  > 90日更新プログラムの Microsoft 365 開発者向けサブスクリプションを無料で入手するには、 [microsoft 365 developer プログラム](https://developer.microsoft.com/office/dev-program)にご参加ください。Microsoft 365 開発者プログラムに参加し、サブスクリプションを構成する方法の詳細な手順については、 [microsoft 365 開発者向けプログラムのドキュメント](/office/developer-program/office-365-developer-program)を参照してください。

- SharePoint Online でアプリカタログを設定します。アプリカタログは、Office アドインのドキュメントライブラリをホストする SharePoint Online の専用サイトコレクションです。独自の SharePoint サイトがある場合は、アプリカタログドキュメントライブラリをセットアップすることができます。詳細については、「 [SharePoint のアプリカタログに作業ウィンドウアドインとコンテンツアドインを発行する](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)」を参照してください。


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

4. Microsoft 365 のアプリ起動ツールから web 上で Excel または Word を起動し、新しいドキュメントを開きます。

5. [挿入] タブで、**[個人用アドイン]** または **[Office アドイン]** をクリックし、アプリにアドインを挿入してテストします。

6. お気に入りのブラウザーのツール デバッガーを使用してアドインをデバッグします。

## <a name="potential-issues"></a>潜在的な問題

以下は、デバッグ時に発生する可能性がある問題です。

- 表示される JavaScript エラーのいくつかは Office on the web に起因している可能性があります。

- ブラウザーに無効な証明書エラーが表示されることがありますが、このエラーはバイパスする必要があります。 これを行うプロセスは、ブラウザおよびこの変更を定期的に行うさまざまなブラウザの UI によって異なります。 詳細については、ブラウザーのヘルプを検索するか、オンラインで検索してください。 (たとえば、「Microsoft Edge の無効な証明書警告」を検索します。) ほとんどのブラウザーには、警告ページにリンクがあり、このリンクをクリックするとアドイン ページにアクセスされます。 たとえば、Microsoft Edge には「Web ページへ移動 (推奨しません)」というリンクがあります。 ただし、通常はアドインが再び読み込まれるたびに、このリンクを経由する必要があります。 継続的なバイパスについては、お勧めのヘルプを参照してください。

- コードにブレークポイントを設定すると、保存できないというエラーが Office on the web からスローされることがあります。

## <a name="see-also"></a>関連項目

- [Office アドイン開発のベスト プラクティス](../concepts/add-in-development-best-practices.md)
- [AppSource の検証ポリシー](/legal/marketplace/certification-policies)  
- [効率的な AppSource アプリおよびアドインを作成する](/office/dev/store/create-effective-office-store-listings)  
- [Office アドインでのユーザー エラーのトラブルシューティング](testing-and-troubleshooting.md)
