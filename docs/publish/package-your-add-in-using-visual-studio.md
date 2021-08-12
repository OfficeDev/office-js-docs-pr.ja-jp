---
title: Visual Studio を使用してアドインを発行する
description: Visual Studio 2019 を使用して Web プロジェクトを展開し、アドインをパッケージ化する方法。
ms.date: 12/02/2019
localization_priority: Normal
ms.openlocfilehash: 7ae9d3289f04d9f3ab969112e9371de51489997e5dc319d0726b94fd5a5e92cd
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57086233"
---
# <a name="publish-your-add-in-using-visual-studio"></a>Visual Studio を使用してアドインを発行する

Office アドイン パッケージには、アドインの発行に使用する XML [マニフェスト ファイル](../develop/add-in-manifests.md)が含まれています。 プロジェクトの Web アプリケーション ファイルは個別に発行する必要があります。 この記事では、Visual Studio 2019 を使用して Web プロジェクトを展開し、アドインをパッケージ化する方法について説明します。

> [!NOTE]
> Yeoman ジェネレーターを使用して作成し、Visual Studio Code またはその他のエディターで開発した Office アドインの発行については、「[Visual Studio Code で開発したアドインの発行](publish-add-in-vs-code.md)」を参照してください。

## <a name="to-deploy-your-web-project-using-visual-studio-2019"></a>Visual Studio 2019 を使用して Web プロジェクトを展開するには

Visual Studio 2019 を使用して Web プロジェクトを展開するには、次の手順を実行します。

1. [**ビルド**] タブから、[**公開 [アドインの名前]**] を選択します。

2. [**発行先の選択**] ウィンドウで、優先されるターゲットに公開するオプションのいずれかを選択します。 各発行ターゲットでは、Azure Virtual Machine やフォルダーの場所など、開始するための詳細な情報を含める必要があります。 公開場所を指定し、必要な情報をすべて入力したら、[**公開**] を選択します

    > [!NOTE]
    > 発行先を選択すると、展開先のサーバー、サーバーへのサインインに必要な資格情報、展開するデータベース、その他の展開オプションが指定されます。

3. 各発行ターゲット オプションの展開手順の詳細については、「[First look at deployment in Visual Studio (Visual Studioでの展開の最初の画面)](/visualstudio/deployment/deploying-applications-services-and-components?view=vs-2019&preserve-view=true)」を参照してください。

## <a name="to-package-and-publish-your-add-in-using-iis-ftp-or-web-deploy-using-visual-studio-2019"></a>IIS、FTP、または Visual Studio 2019 を使用したWeb 配置を使用してアドインをパッケージ化して公開するには

Visual Studio 2019 を使用してアドインをパッケージ化するには、次の手順を実行します。

1. [**ビルド**] タブから、[**公開 [アドインの名前]**] を選択します。
2. [**発行先の選択**]ウィンドウで **IIS、FTPなど** を選択し、[**構成**] を選択します。 次に、[**発行**] を選択します。
3. プロセスをガイドするウィザードが表示されます。 公開方法が Web 配置などの優先される方法であることを確認します。
4. [**接続先 URL**] ボックスに、アドインのコンテンツ ファイルをホストする Web サイトの URL を入力し、[**次へ**] を選択します。 アドインを AppSource に提出する場合には、[**接続の検証**] ボタンを選択し、アドインの受け入れを妨げている問題を特定できます。 アドインをストアに提出する前に、すべての問題に対処する必要があります。
5. **ファイル発行オプション** を含む必要な設定を確認し、[**保存**] を選択します。

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] Azure の Web サイトは自動的に HTTPS エンドポイントを提供します。

XML マニフェストを適切な場所にアップロードして[アドインを発行](../publish/publish.md)できるようになりました。XML マニフェストは、`app.publish` フォルダーの `OfficeAppManifests` にあります。たとえば、次のようになります。

 `%UserProfile%\Documents\Visual Studio 2019\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`

## <a name="see-also"></a>関連項目

- [Office アドインを発行する](../publish/publish.md)
- [AppSource と Office 内でソリューションを使用できるようにする](/office/dev/store/submit-to-the-office-store)
