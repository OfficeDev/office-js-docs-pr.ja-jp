---
title: Excel JavaScript API の要件セット
description: Excel ビルド用の Office アドイン要件セットの情報。
ms.date: 05/05/2021
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 4ba1970f57bb08210878bc3e363598b37eea2265b773f7e533c48939edb6d9e3
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57095911"
---
# <a name="excel-javascript-api-requirement-sets"></a>Excel JavaScript API の要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office アプリケーションがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。

## <a name="requirement-set-availability"></a>要件セットの可用性

Excel アドインは、Windows 上の Office 2016 以降の Office や Micrsoft Offie on the web など複数のバージョンの Office で機能します。次の表は、Excel の要件セット、その要件セットをサポートする Office ホスト アプリケーション、それらのアプリケーションのビルド バージョンまたはビルド番号の一覧です。

> [!NOTE]
> 番号付きの要件セットまたは `ExcelApiOnline` で API を使用するには、CDN: https://appsforoffice.microsoft.com/lib/1/hosted/office.js で **実稼働** ライブラリを参照してください。
>
> プレビューの API の使用に関する詳細については、記事「[Excel JavaScript プレビュー API](excel-preview-apis.md)」を参照してください。

|  要件セット  |  Windows での Office<br>(Microsoft 365 サブスクリプションに接続)  |  Office on iPad<br>(Microsoft 365 サブスクリプションに接続)  |  Office on Mac<br>(Microsoft 365 サブスクリプションに接続)  | Office on the web |
|:-----|-----|:-----|:-----|:-----|:-----|
| [プレビュー](excel-preview-apis.md)  | プレビュー API を試すには、最新版 Office を使用してください (場合によっては、[Office Insider プログラム](https://insider.office.com)に参加する必要があります)。 |
| [ExcelApiOnline](excel-api-online-requirement-set.md) | 該当なし | 該当なし | 該当なし | 最新 ([要件セットのページ](excel-api-online-requirement-set.md)を参照) |
| [ExcelApi 1.12](excel-api-1-12-requirement-set.md) | バージョン 2008 (ビルド 13127.20408) 以降 | 16.40 以降 | 16.40 以降 | 2020 年 9 月 |
| [ExcelApi 1.11](excel-api-1-11-requirement-set.md) | バージョン 2002 (ビルド 12527.20470) 以降 | 16.35 以降 | 16.33 以降 | 2020 年 5 月 |
| [ExcelApi 1.10](excel-api-1-10-requirement-set.md) | バージョン 1907 (ビルド 11929.20306) 以降 | 16.0 以降 | 16.30 以降 | 2019 年 10 月 |
| [ExcelApi 1.9](excel-api-1-9-requirement-set.md)  | バージョン 1903 (ビルド 11425.20204) 以降 | 16.0 以降 | 16.24 以降 | 2019 年 5 月 |
| [ExcelApi 1.8](excel-api-1-8-requirement-set.md)  | バージョン 1808 (ビルド 10730.20102) 以降 | 16.0 以降 | 16.17 以降 | 2018 年 9 月 |
| [ExcelApi 1.7](excel-api-1-7-requirement-set.md)  | バージョン 1801 (ビルド 9001.2171) 以降   | 16.0 以降  | 16.9 以降  | 2018 年 4 月 |
| [ExcelApi 1.6](excel-api-1-6-requirement-set.md)  | バージョン 1704 (ビルド 8201.2001) 以降   | 15.0 以降  | 15.36 以降 | 2017 年 4 月 |
| [ExcelApi 1.5](excel-api-1-5-requirement-set.md)  | バージョン 1703 (ビルド 8067.2070) 以降   | 15.0 以降  | 15.36 以降 | 2017 年 3 月 |
| [ExcelApi 1.4](excel-api-1-4-requirement-set.md)  | バージョン 1701 (ビルド 7870.2024) 以降   | 15.0 以降  | 15.36 以降 | 2017 年 1 月 |
| [ExcelApi 1.3](excel-api-1-3-requirement-set.md)  | バージョン 1608 (ビルド 7369.2055) 以降   | 15.0 以降 | 15.27 以降 | 2016 年 9 月 |
| [ExcelApi 1.2](excel-api-1-2-requirement-set.md)  | バージョン 1601 (ビルド 6741.2088) 以降   | 15.0 以降 | 15.22 以降 | 2016 年 1 月 |
| [ExcelApi 1.1](excel-api-1-1-requirement-set.md)  | バージョン 1509 (ビルド 4266.1001) 以降   | 15.0 以降 | 15.20 以降 | 2016 年 1 月 |

> [!NOTE]
> サブスクリプション版以外の Office でサポートされる要件セットは次のとおりです。
>
> - Office 2019 では ExcelApi 1.8 以前がサポートされています。
> - Office 2016 では ExcelApi 1.1 要求セットのみがサポートされています。

## <a name="office-versions-and-build-numbers"></a>Office のバージョンとビルド番号

Office のバージョンとビルド番号の詳細については、次を参照してください。

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="how-to-use-excel-requirement-sets-at-runtime-and-in-the-manifest"></a>実行時およびマニフェストで Excel 要件セットを使用する方法

> [!NOTE]
> このセクションでは、[Office バージョンと要件セット](../../develop/office-versions-and-requirement-sets.md) の概要、および [Office アプリケーションと API 要件の指定](../../develop/specify-office-hosts-and-api-requirements.md) について理解していることを前提としています。

要件セットは、API メンバーの名前付きグループです。 Office アドインは、Office アプリケーションがアドインの必要とする API をサポートしているかどうかを判断するために、ランタイム チェックを実施したり、マニフェストで指定されている要件セットを使用したりすることができます。

### <a name="checking-for-requirement-set-support-at-runtime"></a>実行時に要件セットのサポートを確認する

次のコード サンプルは、アドインが実行されている Office アプリケーションが指定された API の要件セットをサポートしているかどうかを確認する方法を示しています。

```js
if (Office.context.requirements.isSetSupported('ExcelApi', '1.3')) {
  // Perform actions.
}
else {
  // Provide alternate flow/logic.
}
```

### <a name="defining-requirement-set-support-in-the-manifest"></a>マニフェストで要件セットのサポートを定義する

アドインのマニフェストで [Requirements 要素](../manifest/requirements.md)を使用して、アドインをアクティブにするために必要な最小要件セットや API メソッド (またはその両方) を指定できます。Office アプリケーションまたはプラットフォームが、マニフェストの `Requirements` 要素で指定されている要件セットまたは API メソッドをサポートしていない場合、アドインはそのアプリケーションまたはプラットフォームで実行されず、[**個人用アドイン**] に表示されるアドインの一覧に表示されません。アドインが完全に機能するためには特定の要件セットを必要とするが、その要件セットをサポートしていないプラットフォームのユーザーに対しても価値を提供できる場合は、マニフェストの要件セットのサポートを定義する代わりに、上記のように実行時に要件サポートを確認することをお勧めします。

次のコード サンプルは、アドインが ExcelApi 要件セットのバージョン 1.3 以上をサポートする Office クライアント アプリケーションのすべてで読み込まれる必要があることを指定する、アドインのマニフェストの `Requirements` 要素を示しています。

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

## <a name="see-also"></a>関連項目

- [Excel JavaScript API リファレンス ドキュメント](/javascript/api/excel)
- [Office アドインの XML マニフェスト](../../develop/add-in-manifests.md)
