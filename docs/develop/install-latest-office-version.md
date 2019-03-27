---
title: Office の最新バージョンをインストールする
description: Office の最新ビルドを取得するためにオプトインする方法に関する情報。
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 345b7ad49bab672b9e9dd3a055bd8bfeed2962e3
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871578"
---
# <a name="install-the-latest-version-of-office"></a>Office の最新バージョンをインストールする

新しい開発者用機能 (現時点ではプレビュー版のものを含む) は、Office の最新ビルドの取得をオプトインしたサブスクライバーに最初に配信されます。

## <a name="opt-in-to-getting-the-latest-builds"></a>最新のビルドを取得するためにオプトインする

Office の最新ビルドを取得するためにオプトインするには、次の操作を行います。

- Office 365 Solo のサブスクライバーは、「[Office Insider になる](https://products.office.com/office-insider)」を参照してください。
- 一般法人向け Office 365 をご利用の場合は、「[一般法人向け Office 365 の先行リリース ビルドをインストールする](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead)」を参照してください。
- Mac で Office を実行している場合は、次の操作を行います。
    - Office for Mac プログラムを起動します。
    - [ヘルプ] メニューで [**更新プログラムのチェック**] を選択します。
    - [Microsoft AutoUpdate] ボックスで、チェック ボックスをオンにして Office Insider プログラムに参加します。

## <a name="get-the-latest-build"></a>最新ビルドを取得する

Office の最新ビルドを取得するには、次の操作を行います。

1. [Office 展開ツール](https://www.microsoft.com/download/details.aspx?id=49117)をダウンロードします。
2. ツールを実行します。これにより、Setup.exe および configuration.xml という 2 つのファイルが抽出されます。
3. configuration.xml を[先行リリース構成ファイル](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml)に置き換えます。
4. 次のコマンドを管理者として実行します: `setup.exe /configure configuration.xml`

    > [!NOTE]
    > このコマンドの実行には時間がかかることがあります (進行状況は表示されません)。

インストール処理の完了時点で、最新の Office アプリケーションがインストールされています。 最新のビルドであることを確認するには、任意の Office アプリケーションから **[ファイル]**、**[アカウント]** の順に移動します。 [Office 更新プログラム] に、[(Office Insiders)] ラベルが表示され、その下にバージョン番号が表示されます。

![Office Insiders のラベルと製品情報を示すスクリーンショット](../images/office-insiders.png)

## <a name="minimum-office-builds-for-office-javascript-api-requirement-sets"></a>Office JavaScript API の要件セットの最小 Office ビルド

API の要件セットの各プラットフォームの最小製品ビルドについては、次をご覧ください。

- [Word JavaScript API の要件セット](/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets)
- [Excel JavaScript API の要件セット](/office/dev/add-ins/reference/requirement-sets/excel-api-requirement-sets)
- [OneNote JavaScript API の要件セット](/office/dev/add-ins/reference/requirement-sets/onenote-api-requirement-sets)
- [ダイアログ API の要件セット](/office/dev/add-ins/reference/requirement-sets/dialog-api-requirement-sets)
- [Office 共通 API の要件セット](/office/dev/add-ins/reference/requirement-sets/office-add-in-requirement-sets)
