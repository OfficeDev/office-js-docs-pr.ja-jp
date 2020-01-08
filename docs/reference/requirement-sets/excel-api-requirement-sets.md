---
title: Excel JavaScript API の要件セット
description: Excel ビルド用の Office アドイン要件セットの情報
ms.date: 01/06/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 742c2e925ed7b1be8e1449cd9021c6cf03519d7b
ms.sourcegitcommit: abe8188684b55710261c69e206de83d3a6bd2ed3
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/08/2020
ms.locfileid: "40969525"
---
# <a name="excel-javascript-api-requirement-sets"></a>Excel JavaScript API の要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)」をご覧ください。

## <a name="requirement-set-availability"></a>要件セットの可用性

Excel アドインは、Windows での Office 2016 以降、Office on the web、Mac、および iPad など、複数のバージョンの Office で機能します。 次の表に、Excel の要件セット、各要件セットをサポートする Office ホスト アプリケーション、それらのアプリケーションのビルド バージョンまたはビルド番号を記載します。

> [!NOTE]
> 番号付きの要件セットまたは `ExcelApiOnline` で API を使用するには、CDN: https://appsforoffice.microsoft.com/lib/1/hosted/office.js で**実稼働**ライブラリを参照してください。
>
> プレビューの API の使用に関する詳細については、記事「[Excel JavaScript プレビュー API](./excel-preview-apis.md)」を参照してください。

|  要件セット  |  Windows での Office<br>(Office 365 サブスクリプションに接続)  |  Office on iPad<br>(Office 365 サブスクリプションに接続)  |  Office on Mac<br>(Office 365 サブスクリプションに接続)  | Office on the web |
|:-----|-----|:-----|:-----|:-----|:-----|
| [プレビュー](excel-preview-apis.md)  | プレビュー API を試すには、最新版 Office を使用してください (場合によっては、[Office Insider プログラム](https://products.office.com/office-insider)に参加する必要があります) |
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

- [Office 365 クライアントの更新プログラム チャネル リリースのバージョン番号およびビルド番号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [使用している Office のバージョンを確認する方法](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Office 365 クライアント アプリケーションのバージョン番号およびビルド番号を確認することができます。](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel)
- [Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Office のホストと API の要件を指定する](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office アドインの XML マニフェスト](/office/dev/add-ins/develop/add-in-manifests)
- [Office Online Server 概要](/officeonlineserver/office-online-server-overview)
