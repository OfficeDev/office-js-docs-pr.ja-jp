---
title: Word JavaScript API の要件セット
description: Word ビルド用の Office アドイン要件セットの情報。
ms.date: 05/05/2021
ms.prod: word
localization_priority: Priority
ms.openlocfilehash: 7edead1b1683eca1fd00e92c12043974933864c0ff0efda202c9fcb78f45f249
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57098678"
---
# <a name="word-javascript-api-requirement-sets"></a>Word JavaScript API の要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office アプリケーションがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。

## <a name="requirement-set-availability"></a>要件セットの可用性

Word アドインは、Windows の Office 2016 以降、Office on the web、iPad、および Mac など、複数のバージョンの Office で機能します。次の表は、Word の要件セット、その要件セットをサポートする Office クライアント アプリケーション、およびそれらのアプリケーションのビルド番号またはバージョン番号の一覧です。

> [!NOTE]
> 番号付きの要件セットで API を使用するには、CDN: https://appsforoffice.microsoft.com/lib/1/hosted/office.js で **実稼働** ライブラリを参照してください。
>
> プレビューの API の使用に関する詳細については、記事「[Word JavaScript プレビュー API](word-preview-apis.md)」を参照してください。

|  要件セット  |   Windows での Office\*<br>(Microsoft 365 サブスクリプションに接続)  |  Office on iPad<br>(Microsoft 365 サブスクリプションに接続)  |  Office on Mac<br>(Microsoft 365 サブスクリプションに接続)  | Office on the web  |
|:-----|-----|:-----|:-----|:-----|
| [プレビュー](word-preview-apis.md) | プレビュー API を試すには、最新版 Office を使用してください (場合によっては、[Office Insider プログラム](https://insider.office.com)に参加する必要があります) |
| [WordApi 1.3](word-api-1-3-requirement-set.md) | バージョン 1612 (ビルド 7668.1000) 以降| 2017 年 3 月、2.22 以降 | 2017 年 3 月、15.32 以降| 2017 年 3 月 |
| [WordApi 1.2](word-api-1-2-requirement-set.md) | 2015年 12 月更新プログラム、バージョン 1601 (ビルド 6568.1000) 以降 | 2016 年 1 月、1.18 以降 | 2016 年 1 月、15.19 以降| 2016 年 9 月 |
| [WordApi 1.1](word-api-1-1-requirement-set.md) | バージョン 1509 (ビルド 4266.1001) 以降| 2016 年 1 月、1.18 以降 | 2016 年 1 月、15.19 以降| 2016 年 9 月 |

> [!NOTE]
> サブスクリプション版以外の Office でサポートされる要件セットは次のとおりです。
>
> - Office 2019 では WordApi 1.3 以前がサポートされています。
> - Office 2016 では WordApi 1.1 要求セットのみがサポートされています。

## <a name="office-versions-and-build-numbers"></a>Office のバージョンとビルド番号

Office のバージョンとビルド番号の詳細については、次を参照してください。

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="see-also"></a>関連項目

- [Word JavaScript API リファレンス ドキュメント](/javascript/api/word)
- [Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)
- [Office アプリケーションと API 要件を指定する](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office アドインの XML マニフェスト](../../develop/add-in-manifests.md)
