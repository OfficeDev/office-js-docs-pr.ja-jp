---
title: Excel JavaScript API の要件セット
description: Excel ビルド用の Office アドイン要件セットの情報。
ms.date: 04/16/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 6da9e34359521157e809764907c3a6c3a62ae76c
ms.sourcegitcommit: 803587b324fc8038721709d7db5664025cf03c6b
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/17/2020
ms.locfileid: "43547230"
---
# <a name="excel-javascript-api-requirement-sets"></a>Excel JavaScript API の要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。

## <a name="requirement-set-availability"></a>要件セットの可用性

Excel アドインは、Windows での Office 2016 以降、Office on the web、Mac、および iPad など、複数のバージョンの Office で機能します。 次の表に、Excel の要件セット、各要件セットをサポートする Office ホスト アプリケーション、それらのアプリケーションのビルド バージョンまたはビルド番号を記載します。

> [!NOTE]
> 番号付きの要件セットまたは `ExcelApiOnline` で API を使用するには、CDN: https://appsforoffice.microsoft.com/lib/1/hosted/office.js で**実稼働**ライブラリを参照してください。
>
> プレビューの API の使用に関する詳細については、記事「[Excel JavaScript プレビュー API](excel-preview-apis.md)」を参照してください。

|  要件セット  |  Windows での Office<br>(Office 365 サブスクリプションに接続)  |  Office on iPad<br>(Office 365 サブスクリプションに接続)  |  Office on Mac<br>(Office 365 サブスクリプションに接続)  | Office on the web |
|:-----|-----|:-----|:-----|:-----|:-----|
| [プレビュー](excel-preview-apis.md)  | プレビュー API を試すには、最新版 Office を使用してください (場合によっては、[Office Insider プログラム](https://insider.office.com)に参加する必要があります) |
| [ExcelApiOnline](excel-api-online-requirement-set.md) | 該当なし | 該当なし | 該当なし | 最新 ([要件セットのページ](./excel-api-online-requirement-set.md)を参照) |
| [ExcelApi 1.10](excel-api-1-10-requirement-set.md) | バージョン 1907 (ビルド 11929.20306) 以降 | 2.30 以降 | 16.30 以降 | 2019 年 10 月 |
| [ExcelApi 1.9](excel-api-1-9-requirement-set.md)  | バージョン 1903 (ビルド 11425.20204) 以降 | 2.24 以降 | 16.24 以降 | 2019 年 5 月 |
| [ExcelApi 1.8](excel-api-1-8-requirement-set.md)  | バージョン 1808 (ビルド 10730.20102) 以降 | 2.17 以降 | 16.17 以降 | 2018 年 9 月 |
| [ExcelApi 1.7](excel-api-1-7-requirement-set.md)  | バージョン 1801 (ビルド 9001.2171) 以降   | 2.9 以降  | 16.9 以降  | 2018 年 4 月 |
| [ExcelApi 1.6](excel-api-1-6-requirement-set.md)  | バージョン 1704 (ビルド 8201.2001) 以降   | 2.2 以降  | 15.36 以降 | 2017 年 4 月 |
| [ExcelApi 1.5](excel-api-1-5-requirement-set.md)  | バージョン 1703 (ビルド 8067.2070) 以降   | 2.2 以降  | 15.36 以降 | 2017 年 3 月 |
| [ExcelApi 1.4](excel-api-1-4-requirement-set.md)  | バージョン 1701 (ビルド 7870.2024) 以降   | 2.2 以降  | 15.36 以降 | 2017 年 1 月 |
| [ExcelApi 1.3](excel-api-1-3-requirement-set.md)  | バージョン 1608 (ビルド 7369.2055) 以降   | 1.27 以降 | 15.27 以降 | 2016 年 9 月 |
| [ExcelApi 1.2](excel-api-1-2-requirement-set.md)  | バージョン 1601 (ビルド 6741.2088) 以降   | 1.21 以降 | 15.22 以降 | 2016 年 1 月 |
| [ExcelApi 1.1](excel-api-1-1-requirement-set.md)  | バージョン 1509 (ビルド 4266.1001) 以降   | 1.19 以降 | 15.20 以降 | 2016 年 1 月 |

> [!NOTE]
> 永続ライセンス版 Office でサポートされる要件セットは次のとおりです。
>
> - Office 2019 では ExcelApi 1.8 以前がサポートされています。
> - Office 2016 では ExcelApi 1.1 要求セットのみがサポートされています。

## <a name="office-versions-and-build-numbers"></a>Office のバージョンとビルド番号

Office のバージョンとビルド番号の詳細については、次を参照してください。

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel)
- [Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)
- [Office のホストと API の要件を指定する](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office アドインの XML マニフェスト](../../develop/add-in-manifests.md)
- [Office Online Server 概要](/officeonlineserver/office-online-server-overview)
