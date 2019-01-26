---
title: 発行のための準備として Visual Studio を使用してアドインをパッケージ化する | Microsoft Docs
description: Visual Studio 2017 を使用して Web プロジェクトを展開しアドインをパッケージ化する方法です。
ms.date: 01/25/2018
localization_priority: Priority
ms.openlocfilehash: a135e8e72703c3de60290a9eb7b2e03c63449124
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/23/2019
ms.locfileid: "29386436"
---
# <a name="package-your-add-in-using-visual-studio-to-prepare-for-publishing"></a>発行のための準備として Visual Studio を使用してアドインをパッケージ化する

Office アドイン パッケージには、アドインの発行に使用する XML [マニフェスト ファイル](../develop/add-in-manifests.md)が含まれています。 プロジェクトの Web アプリケーション ファイルは個別に発行する必要があります。 この記事では、Visual Studio 2017 を使用して、Web プロジェクトを展開し、アドインをパッケージ化する方法について説明します。

## <a name="to-deploy-your-web-project-using-visual-studio-2017"></a>Visual Studio 2017 を使用して Web プロジェクトを展開するには

次に示す、Visual Studio 2017 を使用して Web プロジェクトを展開する手順を実行します。

1. **[ソリューション エクスプローラー]** で、アドイン プロジェクトのショートカット メニューを開き、**[発行]** を選択します。
    
    [**アドインの発行**] ページが表示されます。
    
2. **[現在のプロファイル]** ドロップダウン リストで、プロファイルを選択するか、**[新規…]** を選択して新しいプロファイルを作成します。
    
    > [!NOTE]
    > 発行プロファイルでは、展開先のサーバー、サーバーへのログオンに必要な資格情報、展開するデータベース、およびその他の展開オプションを指定します。

    **[新規...]** を選択すると、ウィザードが表示され、その **[発行プロファイルの作成]** ページが表示されます。 このウィザードを使用して、Microsoft Azure などの Web サイトをホストするプロバイダーから発行プロファイルをインポートするか、新しいプロファイルを作成するかして、次の手順でサーバー、資格情報、その他の設定を追加することができます。
    
    発行プロファイルのインポートまたは新しい発行プロファイルの作成の詳細については、「[発行プロファイルの作成](https://msdn.microsoft.com/library/dd465337.aspx#creating_a_profile)」を参照してください。
    
3. **[アドインを発行する]** ページで、**[Web プロジェクトの配置]** リンクを選択します。
    
    **[発行]** ダイアログ ボックスが表示されます。 このウィザードの使用法の詳細については、「[手順: Visual Studio でワンクリック発行を使用して Web プロジェクトを展開する](https://msdn.microsoft.com/library/dd465337.aspx)」を参照してください。
    

## <a name="to-package-your-add-in-using-visual-studio-2017"></a>Visual Studio 2017 を使用してアドインをパッケージ化するには

次に示す、Visual Studio 2017 を使用してアドインをパッケージ化する手順を実行します。

1. **[アドインの発行]** ページで、**[アドインのパッケージ]** を選択します。
    
    ウィザードが表示され、その **[アドインのパッケージ]** ページが表示されます。
    
2. **[Web サイトがホストされている場所]** ドロップダウン リストで、アドインのコンテンツ ファイルをホストする Web サイトの URL を選択するか入力して、**[完了]** を選択します。
    
    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] Azure の Web サイトは自動的に HTTPS エンドポイントを提供します。

    Visual Studio は、アドインの発行に必要なファイルを生成して、発行の出力フォルダーを開きます。
    
AppSource にアドインを提出する予定がある場合は、**[検証チェックを実行する]** をクリックして、アドインが受け入れられなくなる問題点を特定します。 アドインをストアに提出する前に、すべての問題を解決してください。

XML マニフェストを適切な場所にアップロードして[アドインを発行](../publish/publish.md)できるようになりました。XML マニフェストは、`app.publish` フォルダーの `OfficeAppManifests` にあります。たとえば、次のようになります。

 `%UserProfile%\Documents\Visual Studio 2017\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## <a name="see-also"></a>関連項目

- [Office アドインを発行する](../publish/publish.md)
- [AppSource と Office 内でソリューションを使用できるようにする](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store)
    
