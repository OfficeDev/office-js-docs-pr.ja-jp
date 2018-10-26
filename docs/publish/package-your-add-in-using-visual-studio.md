---
title: 発行のための準備として Visual Studio を使用してアドインをパッケージ化する| Microsoft Docs
description: この記事では、Visual Studio 2017 を使用して、Web プロジェクトを展開し、アドインをパッケージ化する方法について説明します。
ms.date: 01/25/2018
ms.openlocfilehash: 3515f88e41bc5f0af62a3b043beae5177f3291ac
ms.sourcegitcommit: c400a220783b03a739449e2d3ff00bbffe5ec7c1
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/20/2018
ms.locfileid: "25681764"
---
# <a name="package-your-add-in-using-visual-studio-to-prepare-for-publishing"></a>発行のための準備として Visual Studio を使用してアドインをパッケージ化する

Office アドイン パッケージには、アドインの発行に使用する XML [マニフェスト ファイル](../develop/add-in-manifests.md)が含まれています。プロジェクトの web アプリケーション ファイルを個別に発行する必要があります。この記事では、Visual Studio 2017 を使用して、Web プロジェクトを展開し、アドインをパッケージ化する方法について説明します。

## <a name="to-deploy-your-web-project-using-visual-studio-2017"></a>Visual Studio 2017 を使用して Web プロジェクトを展開するには

Visual Studio 2017 を使用して Web プロジェクトを展開する次の手順を完了します。

1. **[ソリューション エクスプローラー]** で、アドイン プロジェクトのショートカット メニューを開き、**[発行]** を選択します。
    
    [**アドインの発行**] ページが表示されます。
    
2. **[現在のプロファイル]** ドロップダウン リストで、プロファイルを選択するか、**[新規…]** を選択して新しいプロファイルを作成します。
    
    > [!NOTE]
    > 発行プロファイルでは、展開先のサーバー、サーバーへのログオンに必要な資格情報、展開するデータベース、およびその他の展開オプションを指定します。

    [**新規...**] を選択した場合、[**発行プロファイルの作成**] ページとともにウィザードが表示されます。このウィザードを使用して、Microsoft Azure などの Web サイトをホストするプロバイダーから発行プロファイルをインポートするか、新しいプロファイルを作成するかして、次の手順でサーバー、資格情報、その他の設定を追加することができます。
    
    発行プロファイルのインポートまたは新しい発行プロファイルの作成の詳細については、「[発行プロファイルの作成](https://msdn.microsoft.com/library/dd465337.aspx#creating_a_profile)」を参照してください。
    
3. [**アドインを発行する**] ページで、[**Web プロジェクトの展開**] リンクを選択します。
    
    **[Web を発行する]** ダイアログ ボックスが表示されます。このウィザードの使用方法については、「[方法: Visual Studio でワンクリック発行を使用して Web プロジェクトを展開する](https://msdn.microsoft.com/library/dd465337.aspx)」を参照してください。
    

## <a name="to-package-your-add-in-using-visual-studio-2017"></a>Visual Studio 2017 を使用してアドインをパッケージ化するには

Visual Studio 2017 を使用してアドインをパッケージ化する次の手順を完了します。

1. **[アドインの発行]** ページで、**[アドインのパッケージ]** ボタンを選択します。
    
    ウィザードが **[アドインのパッケージ]** ページと共に表示されます。
    
2. **[Web サイトがホストされている場所]** ボックスで、アドインのコンテンツ ファイルをホストする Web サイトの URL を入力して、**[完了]** を選択します。
    
    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] Azure の Web サイトは自動的に HTTPS エンドポイントを提供します。

    Visual Studio は、アドインの発行に必要なファイルを生成して、発行の出力フォルダーを開きます。
    
AppSource にアドインの提出を予定している場合は、**[検証チェックの実行]** ボタンをクリックして、アドインの受け入れが阻害される問題点を識別します。 アドインをストアに提出する前に、すべての問題に対処してください。

XML マニフェストを適切な場所にアップロードして[アドインを発行](../publish/publish.md)できるようになりました。XML マニフェストは、`app.publish`フォルダーの`OfficeAppManifests`にあります。たとえば、次のようになります。

 `%UserProfile%\Documents\Visual Studio 2017\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## <a name="see-also"></a>関連項目

- [Office アドインを発行する](../publish/publish.md)
- [AppSource と Office 内でソリューションを使用できるようにする](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store)
    
