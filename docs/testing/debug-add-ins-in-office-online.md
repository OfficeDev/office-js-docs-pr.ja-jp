---
title: Office on the web でアドインをデバッグする
description: Office on the web を使用してアドインをテストおよびデバッグする方法。
ms.date: 07/07/2020
ms.localizationpriority: medium
ms.openlocfilehash: 255826f8925ea35d25cf228e80de6774c9917cea
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152993"
---
# <a name="debug-add-ins-in-office-on-the-web"></a>Office on the web でアドインをデバッグする

Windows、Office 2013、または Office 2016 デスクトップ クライアントを実行していないコンピューター (たとえば、Mac で開発を行っている場合) でアドインの作成とデバッグを行えます。この記事では、Office Online を使用してアドインのテストとデバッグを行う方法について説明します。 この記事では、Office on the web を使用してアドインをテストおよびデバッグする方法について説明します。 

## <a name="prerequisites"></a>前提条件

開始するには

- 開発者アカウントMicrosoft 365持ってない場合、またはサイトにアクセスできる場合は、SharePointします。

  > [!NOTE]
  > 開発者向けの無料の 90 日間の更新プログラムをMicrosoft 365、開発者向けプログラムMicrosoft 365[参加してください](https://developer.microsoft.com/office/dev-program)。 開発者プログラム[Microsoft 365](/office/developer-program/office-365-developer-program)に参加してサブスクリプションを構成する方法の詳細については、開発者プログラムのドキュメントMicrosoft 365参照してください。

- [オンライン] でアプリ カタログをSharePointします。 アプリ カタログは、SharePoint アドイン用のドキュメント ライブラリをホストする Office サイト コレクションです。サイトに独自のSharePoint場合は、アプリ カタログ ドキュメント ライブラリを設定できます。 詳細については、「タスク ウィンドウと[コンテンツ](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)アドインをアプリ カタログに公開する」を参照SharePoint。


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

4. アプリExcel起動Word on the web起動して、新しいMicrosoft 365を開きます。

5. [挿入] タブで、**[個人用アドイン]** または **[Office アドイン]** をクリックし、アプリにアドインを挿入してテストします。

6. お気に入りのブラウザーのツール デバッガーを使用してアドインをデバッグします。

## <a name="potential-issues"></a>潜在的な問題

デバッグ時に発生する可能性があるいくつかの問題を次に示します。

- 表示される JavaScript エラーのいくつかは Office on the web に起因している可能性があります。

- ブラウザーに無効な証明書エラーが表示されることがありますが、このエラーはバイパスする必要があります。 これを行うプロセスは、ブラウザおよびこの変更を定期的に行うさまざまなブラウザの UI によって異なります。 詳細については、ブラウザーのヘルプを検索するか、オンラインで検索してください。 (たとえば、「Microsoft Edge の無効な証明書警告」を検索します。) ほとんどのブラウザーには、警告ページにリンクがあり、このリンクをクリックするとアドイン ページにアクセスされます。 たとえば、Microsoft Edge には「Web ページへ移動 (推奨しません)」というリンクがあります。 ただし、通常はアドインが再び読み込まれるたびに、このリンクを経由する必要があります。 継続的なバイパスについては、お勧めのヘルプを参照してください。

- コードにブレークポイントを設定すると、保存できないというエラーが Office on the web からスローされることがあります。

## <a name="see-also"></a>関連項目

- [Office アドイン開発のベスト プラクティス](../concepts/add-in-development-best-practices.md)
- [AppSource の検証ポリシー](/legal/marketplace/certification-policies)  
- [効率的な AppSource アプリおよびアドインを作成する](/office/dev/store/create-effective-office-store-listings)  
- [Office アドインでのユーザー エラーのトラブルシューティング](testing-and-troubleshooting.md)
