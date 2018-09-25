---
title: 発行のための準備として Visual Studio を使用してアドインをパッケージ化する| Microsoft Docs
description: この記事では、Visual Studio 2015 を使用して、Web プロジェクトを展開し、アドインをパッケージ化する方法について説明します。
ms.date: 01/25/2018
ms.openlocfilehash: d74ead03b8ac5b7652c7c98851e7e082f4b31ba8
ms.sourcegitcommit: eb74e94d3e1bc1930a9c6582a0a99355d0da34f2
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/25/2018
ms.locfileid: "25004918"
---
# <a name="package-your-add-in-using-visual-studio-to-prepare-for-publishing"></a>発行のための準備として Visual Studio を使用してアドインをパッケージ化する

Office アドイン パッケージには、アドインの発行に使用する XML [マニフェスト ファイル](../develop/add-in-manifests.md)が含まれています。 プロジェクトの Web アプリケーション ファイルは個別に発行する必要があります。 この記事では、Visual Studio 2015 を使用して、Web プロジェクトを展開し、アドインをパッケージ化する方法について説明します。

## <a name="to-deploy-your-web-project-using-visual-studio-2015"></a>Visual Studio 2015 を使用して Web プロジェクトを展開するには

次に示す、Visual Studio 2015 を使用して Web プロジェクトを展開する手順を実行します。

1. **[ソリューション エクスプローラー]** で、アドイン プロジェクトのショートカット メニューを開き、**[発行]** を選択します。
    
    [**アドインの発行**] ページが表示されます。
    
2. **[現在のプロファイル]** ドロップダウン リストで、プロファイルを選択するか、**[新規…]** を選択して新しいプロファイルを作成します。
    
    > [!NOTE]
    > 発行プロファイルでは、展開先のサーバー、サーバーへのログオンに必要な資格情報、展開するデータベース、およびその他の展開オプションを指定します。

    ** [新規...] ** を選択すると、[発行プロファイルの作成] ウィザードが表示されます。 このウィザードを使用して、Microsoft Azure などの Web サイトをホストするプロバイダーから発行プロファイルをインポートするか、新しいプロファイルを作成するかして、次の手順でサーバー、資格情報、その他の設定を追加することができます。
    
    発行プロファイルのインポートまたは新しい発行プロファイルの作成の詳細については、「[発行プロファイルの作成](https://msdn.microsoft.com/library/dd465337.aspx#creating_a_profile)」を参照してください。
    
3. [ **アドインを発行する**] ページで、 [ **Web プロジェクトの配置**] リンクを選択します。
    
    **[Web を発行する]** ダイアログ ボックスが表示されます。このウィザードの使用方法については、「[方法: Visual Studio でワンクリック発行を使用して Web アプリケーション プロジェクトを配置する](https://msdn.microsoft.com/library/dd465337.aspx)」を参照してください。
    

## <a name="to-package-your-add-in-using-visual-studio-2015"></a>Visual Studio 2015 を使用してアドインをパッケージ化するには

次に示す、Visual Studio 2015 を使用してアドインをパッケージ化する手順を実行します。

1. **[アドインを発行する]** ページで、**[アドインのパッケージ化]** リンクをクリックします。
    
    Office/SharePoint アドインの発行 ウィザードが表示されます。
    
2. **[Web サイトがホストされている場所]** ドロップダウン リストで、アドインのコンテンツ ファイルをホストする Web サイトの HTTPS URL を選択するか入力して、**[完了]** を選択します。 
    
    このウィザードを完了するには、HTTPS プレフィックスで始まる URL を指定する必要があります。Web サイトの HTTP エンドポイントを使用する場合は、パッケージの作成の完了後に、テキスト エディターで XML マニフェスト ファイルを開いて、Web サイトの HTTPS プレフィックスを HTTP プレフィックスに置換します。 

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] Azure の Web サイトは自動的に HTTPS エンドポイントを提供します。

    Visual Studio は、アドインの発行に必要なファイルを生成して、発行の出力フォルダーを開きます。 
    
AppSource へのアドインの提出を予定している場合は、**[検証チェックの実行]** リンクをクリックして、アドインの受け入れが阻害される問題点を識別します。アドインをストアに提出する前に、すべての問題に対処してください。

XML マニフェストを適切な場所にアップロードして[アドインを発行](../publish/publish.md)できるようになりました。XML マニフェストは、`app.publish` フォルダーの `OfficeAppManifests` にあります。たとえば、次のようになります。

 `%UserProfile%\Documents\Visual Studio 2015\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## <a name="see-also"></a>関連項目

- [Office アドインを発行する](../publish/publish.md)
- [AppSource と Office 内でソリューションを使用できるようにする](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store)
    
