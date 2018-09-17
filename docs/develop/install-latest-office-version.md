---
title: Office の最新バージョンをインストールする
description: Office の最新のビルドを取得するを有効にする方法に関する情報です。
ms.date: 12/04/2017
ms.openlocfilehash: 14e26d9fa9f7ec3b2724cbf2e9787cde9dbe4094
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2018
ms.locfileid: "23943881"
---
# <a name="install-the-latest-version-of-office"></a>Office の最新バージョンをインストールする

新しい開発者用機能 (現時点ではプレビュー版のものを含む) は、Office の最新ビルドの取得をオプトインしたサブスクライバーに最初に配信されます。 

## <a name="opt-in-to-getting-the-latest-builds"></a>最新のビルドを取得するためにオプトインする

Office 2016 の最新ビルドを取得するためにオプトインするには: 

- Office 365 Home、Personal、または University のサブスクライバーは、「[Office Insider プログラム](https://products.office.com/office-insider)」を参照してください。
- 一般法人向け Office 365 をご利用の場合は、「[一般法人向け Office 365 の先行リリース ビルドをインストールする](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead)」を参照してください。
- Mac で Office 2016 を実行している場合:
    - Office 2016 for Mac プログラムを起動します。
    - [ヘルプ] メニューで [**更新プログラムのチェック**] を選びます。
    - [Microsoft AutoUpdate] ボックスで、チェック ボックスをオンにして Office Insider プログラムに参加します。 

## <a name="get-the-latest-build"></a>最新ビルドを取得する:

Office 2016 の最新ビルドを取得するには: 

1. [ Office 展開ツールのダウンロード](https://www.microsoft.com/download/details.aspx?id=49117) 。 
2. ツールを実行します。これにより、Setup.exe および configuration.xml という 2 つのファイルが抽出されます。
3. configuration.xml を[先行リリース構成ファイル](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml)に置き換えます。
4. 次のコマンドを管理者として実行します:  `setup.exe /configure configuration.xml` 

    > [!NOTE]
    > このコマンドの実行には時間がかかることがあります (進行状況は表示されません)。

インストール プロセスが完了すると、インストールされている最新の Office アプリケーションがあります。 最新のビルドがあることを確認するには、 **ファイル**に移動 > 任意の Office アプリケーションからの**アカウント** です。 [Office 更新プログラム] に、[(Office Insiders)] ラベルが表示され、その下にバージョン番号が表示されます。

![Office Insiders のラベルと製品情報を示すスクリーンショット](../images/office-insiders.png)

## <a name="minimum-office-builds-for-office-javascript-api-requirement-sets"></a>Office JavaScript API の要件セットの最小 Office ビルド

API の要件セットの各プラットフォームの最小製品ビルドについては、次をご覧ください。

- [Word JavaScript API の要件セット](https://docs.microsoft.com/javascript/office/requirement-sets/word-api-requirement-sets?view=office-js)
- [Excel JavaScript API の要件セット](https://docs.microsoft.com/javascript/office/requirement-sets/excel-api-requirement-sets?view=office-js)
- [OneNote JavaScript API の要件セット](https://docs.microsoft.com/javascript/office/requirement-sets/onenote-api-requirement-sets?view=office-js)
- [ダイアログ API の要件セット](https://docs.microsoft.com/javascript/office/requirement-sets/dialog-api-requirement-sets?view=office-js)
- [Office 共通 API の要件セット](https://docs.microsoft.com/javascript/office/requirement-sets/office-add-in-requirement-sets?view=office-js)
