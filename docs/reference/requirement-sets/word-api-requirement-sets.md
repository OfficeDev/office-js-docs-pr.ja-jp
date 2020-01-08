---
title: Word JavaScript API の要件セット
description: Word ビルド用の Office アドイン要件セットの情報
ms.date: 01/06/2020
ms.prod: word
localization_priority: Priority
ms.openlocfilehash: c90daafe46d301b404ee902b38bb7417562adc44
ms.sourcegitcommit: abe8188684b55710261c69e206de83d3a6bd2ed3
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/08/2020
ms.locfileid: "40969532"
---
# <a name="word-javascript-api-requirement-sets"></a>Word JavaScript API の要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)」をご覧ください。

## <a name="requirement-set-availability"></a>要件セットの可用性

Word アドインは、Windows での Office 2016 以降、Office on the web、iPad、および Mac など、複数のバージョンの Office で機能します。 次の表は、Word の要件セット、その要件セットをサポートする Office ホスト アプリケーション、およびそれらのアプリケーションのビルド番号またはバージョン番号の一覧です。

> [!NOTE]
> 番号付きの要件セットで API を使用するには、CDN: https://appsforoffice.microsoft.com/lib/1/hosted/office.js で**実稼働**ライブラリを参照してください。
>
> プレビューの API の使用に関する詳細については、記事「[Excel JavaScript プレビュー API](word-preview-apis.md)」を参照してください。

|  要件セット  |   Windows での Office\*<br>(Office 365 サブスクリプションに接続)  |  Office on iPad<br>(Office 365 サブスクリプションに接続)  |  Office on Mac<br>(Office 365 サブスクリプションに接続)  | Office on the web  |
|:-----|-----|:-----|:-----|:-----|
| [プレビュー](word-preview-apis.md) | プレビュー API を試すには、最新版 Office を使用してください (場合によっては、[Office Insider プログラム](https://products.office.com/office-insider)に参加する必要があります) |
| [WordApi 1.3](word-api-1-3-requirement-set.md) | バージョン 1612 (ビルド 7668.1000) 以降| 2017 年 3 月、2.22 以降 | 2017 年 3 月、15.32 以降| 2017 年 3 月 |
| [WordApi 1.2](word-api-1-2-requirement-set.md) | 2015年 12 月更新プログラム、バージョン 1601 (ビルド 6568.1000) 以降 | 2016 年 1 月、1.18 以降 | 2016 年 1 月、15.19 以降| 2016 年 9 月 |
| [WordApi 1.1](word-api-1-1-requirement-set.md) | バージョン 1509 (ビルド 4266.1001) 以降| 2016 年 1 月、1.18 以降 | 2016 年 1 月、15.19 以降| 2016 年 9 月 |

> [!NOTE]
> 永続ライセンス版 Office でサポートされる要件セットは次のとおりです。
>
> - Office 2019 では WordApi 1.3 以前がサポートされています。
> - Office 2016 では WordApi 1.1 要求セットのみがサポートされています。

## <a name="office-versions-and-build-numbers"></a>Office のバージョンとビルド番号

Office のバージョンとビルド番号の詳細については、次を参照してください。

- [Office 365 クライアントの更新プログラム チャネル リリースのバージョン番号およびビルド番号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [使用している Office のバージョンを確認する方法](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Office 365 クライアント アプリケーションのバージョン番号およびビルド番号を確認することができます。](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)

## <a name="see-also"></a>関連項目

- [Word JavaScript API リファレンス ドキュメント](/javascript/api/word)
- [Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Office のホストと API の要件を指定する](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office アドインの XML マニフェスト](/office/dev/add-ins/develop/add-in-manifests)
