---
title: テスト用に Office Online で Office アドインをサイドロードする
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 69b255545525ff667618c9f8bd1e1b7953592967
ms.sourcegitcommit: 58af795c3d0393a4b1f6425fa1cbdca1e48fb473
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/29/2018
ms.locfileid: "20138850"
---
# <a name="sideload-office-add-ins-in-office-online-for-testing"></a>テスト用に Office Online で Office アドインをサイドロードする

まずアドイン カタログに置かなくても、サイドロードを使用すると、テスト用に Office アドインをインストールすることができます。サイドロードは、Office 365 または Office Online 上のいずれかで実行できます。2 つのプラットフォームで手順が少し異なります。 

アドインをサイドロードするとき、アドイン マニフェストはブラウザーのローカル ストレージに格納されます。そのため、ブラウザーのキャッシュを消去したり、別のブラウザーに切り替えたりする場合、アドインを再びサイドロードする必要があります。


> [!NOTE]
> この記事で説明したようにサイドロードは、Word、Excel、および PowerPoint でサポートされています。Outlook アドインをサイドロードするには、「[テストのために Outlook アドインをサイドロードする](https://docs.microsoft.com/en-us/outlook/add-ins/sideload-outlook-add-ins-for-testing)」をご参照ください。

次のビデオでは、Office デスクトップまたは Office Online のアドインをサイドロードする手順について説明します。  


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]

## <a name="sideload-an-office-add-in-on-office-365"></a>Office アドインを Office 365 にサイドロードする


1. Office 365 サイトにサインインします。
    
2. ツールバーの左端にあるアプリ起動ツールを開き、**Excel**、**Word**、または **PowerPoint** を選択して、新しいドキュメントを作成します。
    
3. リボンの  **[挿入]** タブを開き、 **[アドイン]** セクションで、 **Office [アドイン]** を選択します。
    
4. **[Office アドイン]** ダイアログ ボックスで、**[自分の所属組織]** タブ、**[個人用アドインのアップロード]** の順に選択します。
    
    ![左上隅近くの、リンクが付いている Office アドインのダイアログ。タイトルは、[マイ アドインのアップロード]](../images/office-add-ins.png)

5.  アドイン マニフェスト ファイルを**参照**して、**[アップロード]** を選択します。
    
    ![[参照]、[アップロード]、[キャンセル] のボタンがある [アドインのアップロード] ダイアログ](../images/upload-add-in.png)

6. アドイン がインストールされていることを確認します。たとえば、アドイン コマンドである場合は、リボンまたはコンテキスト メニューのいずれかに表示されます。作業ウィンドウ アドインである場合は、ウィンドウが表示されます。
    

## <a name="sideload-an-office-add-in-on-office-online"></a>Office アドイン を Office Online にサイドロードする


1. [Microsoft Office Online](https://office.live.com/) を開きます。
    
2. **[オンライン アプリを今すぐ開始する]** で、 **Excel**、 **Word**、または  **PowerPoint** を選択して、新しいドキュメントを開きます。
    
3. リボンの  **[挿入]** タブを開き、 **[アドイン]** セクションで、 **Office [アドイン]** を選択します。
    
4. **[Office アドイン]** ダイアログ ボックスで、**[個人用アドイン]** タブ、**[個人用アドインの管理]**、**[個人用アドインのアップロード]** の順に選択します。
    
    ![右上に [個人用アドインの管理] というドロップダウンがあり、その下に [マイ アドインのアップロード] オプションのドロップダウンがある [Office アドイン] ダイアログ](../images/office-add-ins-my-account.png)

5.  アドイン マニフェスト ファイルを**参照**して、**[アップロード]** を選択します。
    
    ![[参照]、[アップロード]、[キャンセル] のボタンがある [アドインのアップロード] ダイアログ。](../images/upload-add-in.png)

6. アドインがインストールされていることを確認します。たとえば、アドイン コマンドである場合は、リボンまたはコンテキスト メニューのいずれかに表示されます。作業ウィンドウ アドインである場合は、ウィンドウが表示されます。

> [!NOTE]
>Office アドインを Edge でテストするには、Edge の検索バーに "**abou:flags**" を入力し、[開発者設定] オプションを表示します。  "**ローカルホスト ループバックを許可する**" オプションにチェックを入れ、Edgeを再起動します。

>    ![Edge の [ローカルホスト ループバックを許可する] オプションにチェックを入れます。](../images/allow-localhost-loopback.png)

## <a name="sideload-an-add-in-when-using-visual-studio"></a>Visual Studio の使用時にアドインをサイドロードする

アドインの開発に Visual Studio を使用している場合、サイドロードするプロセスは似ています。唯一の違いは、マニフェストの **SourceURL** 要素の値を更新して、アドインが展開されている完全な URL を含める必要がある点です。 

現在アドインを開発している場合、アドイン manifest.xml ファイルを見つけて、**SourceLocation** 要素の値を更新することにより、絶対 URI を含めます。Visual Studio は、localhost を展開するためのトークンを配置します。

例: 

```xml
<SourceLocation DefaultValue="https://localhost:44300/App/Home/Home.html" />
```
