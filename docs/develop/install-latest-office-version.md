---
title: Office の最新バージョンをインストールする
description: Office の最新ビルドを取得するためにオプトインする方法に関する情報。
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: d72012e3c2e642c74d8573c4d9bb3b29a8fc0274
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076021"
---
# <a name="install-the-latest-version-of-office"></a>Office の最新バージョンをインストールする

新しい開発者用機能 (現時点ではプレビュー版のものを含む) は、Office の最新ビルドの取得をオプトインしたサブスクライバーに最初に配信されます。

## <a name="opt-in-to-getting-the-latest-builds"></a>最新のビルドを取得するためにオプトインする

Office の最新ビルドを取得するためにオプトインするには、次の操作を行います。

- ユーザー、個人、またはMicrosoft 365 Familyのサブスクライバーの場合は[、「Be a Office Insider」を参照してください](https://insider.office.com)。
- 顧客の場合は、「Microsoft 365 Apps for businessの最初のリリース ビルドをインストールする[」をMicrosoft 365 Apps for businessしてください](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead)。
- Mac で Office を実行している場合は、次の操作を行います。
  - Office アプリケーションを起動します。
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

![製品情報と Insiders ラベルをOfficeスクリーンショット。](../images/office-insiders-label.png)

## <a name="minimum-office-builds-for-office-javascript-api-requirement-sets"></a>Office JavaScript API の要件セットの最小 Office ビルド

API の要件セットの各プラットフォームの最小製品ビルドについては、次をご覧ください。

- [Excel JavaScript API の要件セット](../reference/requirement-sets/excel-api-requirement-sets.md)
- [OneNote JavaScript API の要件セット](../reference/requirement-sets/onenote-api-requirement-sets.md)
- [Outlook JavaScript API の要件セット](../reference/requirement-sets/outlook-api-requirement-sets.md)
- [PowerPoint JavaScript API の要件セット](../reference/requirement-sets/powerpoint-api-requirement-sets.md)
- [Word JavaScript API の要件セット](../reference/requirement-sets/word-api-requirement-sets.md)
- [ダイアログ API の要件セット](../reference/requirement-sets/dialog-api-requirement-sets.md)
- [Office 共通 API の要件セット](../reference/requirement-sets/office-add-in-requirement-sets.md)
