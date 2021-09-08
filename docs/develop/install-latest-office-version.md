---
title: Office の最新バージョンをインストールする
description: Office の最新ビルドを取得するためにオプトインする方法に関する情報。
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: f5e1bf7afe89efd6f523a74b04d7b6a99f0195bf
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937744"
---
# <a name="install-the-latest-version-of-office"></a>Office の最新バージョンをインストールする

新しい開発者用機能 (現時点ではプレビュー版のものを含む) は、Office の最新ビルドの取得をオプトインしたサブスクライバーに最初に配信されます。

## <a name="opt-in-to-getting-the-latest-builds-of-office"></a>最新のビルドの取得をオプトインOffice

- ユーザー、個人、またはMicrosoft 365 Familyのサブスクライバーの場合は[、「Be a Office Insider」を参照してください](https://insider.office.com)。
- 顧客の場合は、「Microsoft 365 Apps for businessの最初のリリース ビルドをインストールする[」をMicrosoft 365 Apps for businessしてください](https://support.office.com/article/4dd8ba40-73c0-4468-b778-c7b744d03ead)。
- Mac で Office を実行している場合は、次の操作を行います。
  - Office アプリケーションを起動します。
  - [ヘルプ] メニューで [**更新プログラムのチェック**] を選択します。
  - [Microsoft AutoUpdate] ボックスで、チェック ボックスをオンにして Office Insider プログラムに参加します。

## <a name="get-the-latest-build-of-office"></a>最新のビルドを取得Office

1. [Office 展開ツール](https://www.microsoft.com/download/details.aspx?id=49117)をダウンロードします。
2. ツールを実行します。これにより、Setup.exe および configuration.xml という 2 つのファイルが抽出されます。
3. configuration.xml を[先行リリース構成ファイル](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml)に置き換えます。
4. 次のコマンドを管理者として実行します: `setup.exe /configure configuration.xml`

> [!NOTE]
> このコマンドの実行には時間がかかることがあります (進行状況は表示されません)。

インストール処理の完了時点で、最新の Office アプリケーションがインストールされています。 最新のビルドであることを確認するには、任意の Office アプリケーションから **[ファイル]**、**[アカウント]** の順に移動します。 [Office 更新プログラム] に、[(Office Insiders)] ラベルが表示され、その下にバージョン番号が表示されます。

![製品情報と Insiders ラベルをOfficeスクリーンショット。](../images/office-insiders-label.png)

## <a name="minimum-office-builds-for-office-javascript-api-requirement-sets"></a>Office JavaScript API の要件セットの最小 Office ビルド

- [Excel JavaScript API の要件セット](../reference/requirement-sets/excel-api-requirement-sets.md)
- [OneNote JavaScript API の要件セット](../reference/requirement-sets/onenote-api-requirement-sets.md)
- [Outlook JavaScript API の要件セット](../reference/requirement-sets/outlook-api-requirement-sets.md)
- [PowerPoint JavaScript API の要件セット](../reference/requirement-sets/powerpoint-api-requirement-sets.md)
- [Word JavaScript API の要件セット](../reference/requirement-sets/word-api-requirement-sets.md)
- [ダイアログ API の要件セット](../reference/requirement-sets/dialog-api-requirement-sets.md)
- [Office 共通 API の要件セット](../reference/requirement-sets/office-add-in-requirement-sets.md)
