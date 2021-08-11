---
title: ダイアログ API の要件セット
description: ダイアログ API 要件セットの詳細について説明します。
ms.date: 07/19/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 3c5aea3fecc6b48a830e48cf7739e93ef16dab6bacee1338b94774911a06ef5d
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57098945"
---
# <a name="dialog-api-requirement-sets"></a>ダイアログ API の要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office アプリケーションがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。

Office アドインは Office の複数のバージョンで機能します。 次の表に、Dialog API 要件セット、その要件セットをサポートする Office クライアント アプリケーション、およびアプリケーションのビルドまたはバージョン番号をOfficeします。

|  要件セット  | Windows 版 Office 2013\*<br>(1 回限りの購入) | Office 2016 以降のWindows\*<br>(1 回限りの購入)   | Windows での Office<br>(サブスクリプション) |  Office on iPad<br>(サブスクリプション)  |  Office on Mac<br>(サブスクリプション)  | Office on the web  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| DialogApi 1.2  | 該当なし | 該当なし | サポートを見る<br>下のセクション | 2.37 以降 | 16.37 以降 | 2020 年 6 月 | 該当なし |
| DialogApi 1.1  | ビルド 15.0.4855.1000 以降 | ビルド 16.0.4390.1000 以降 | バージョン 1602 (ビルド 6741.0000) 以降 | 1.22 以降 | 15.20 以降 | 2017 年 1 月 | バージョン 1608 (ビルド 7601.6800) 以降|

>\*一度に購入したユーザーはOffice更新プログラムを受け入れてない可能性があります。 その場合、dialogApi をサポートするために更新された DLL がユーザーのコンピューターにインストールされていない場合でも、Office が UI でバージョンを報告するために使用する DLL は、ここに示されているバージョンよりも大きい場合があります。 必要なパッチがインストールされていることを確認するには、Office 更新リスト ([Office 2013 リストまたは Office 2016](/officeupdates/msp-files-office-2013)リスト) に移動し **、osfclient-x-none** を検索し、一覧に記載されている更新プログラムをインストールする必要があります。 [](/officeupdates/msp-files-office-2016)

## <a name="office-on-windows-subscription-support"></a>Office (Windows) のサポートに関する情報

DialogApi 1.2 要件セットは、コンシューマー チャネル バージョン 2005 (ビルド、12827.20268 以上) でサポートされています。 Windows Office では、この機能は 2020 年 6 月 9 日以降に利用可能な Semi-Annual チャネルおよび月次 Enterprise チャネル ビルドでもサポートされます。 各チャネルでサポートされる最小ビルドは次のとおりです。  

|チャネル | バージョン | ビルド|
|:-----|:-----|:-----|
|現在のチャネル | 2005 以上 | 12827.20160 以上|
|月次エンタープライズ チャネル | 2004 以上 | 12730.20430 以上|
|半期エンタープライズ チャネル | 2002 以上 | 12527.20720 以上|

## <a name="office-versions-and-build-numbers"></a>Office のバージョンとビルド番号

バージョン、ビルド番号、Office Online Server の詳細については以下を参照してください。

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server 概要](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。

## <a name="dialog-api-11-and-12"></a>ダイアログ API 1.1 および 1.2

ダイアログ API 1.1 は、API の最初のバージョンです。 要件セット 1.2 は、親ページから[Office.dialog.messageChild](/javascript/api/office/office.dialog#messageChild_message_)メソッドを使用してダイアログ ボックスにデータを送信するサポートを追加します。 これらの API の詳細については [、「Dialog API リファレンス」を](/javascript/api/office/office.ui) 参照してください。

## <a name="see-also"></a>関連項目

- [Office アドインで Office ダイアログ API を使用する](../../develop/dialog-api-in-office-add-ins.md)
- [Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)
- [Office アプリケーションと API 要件を指定する](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office アドインの XML マニフェスト](../../develop/add-in-manifests.md)
