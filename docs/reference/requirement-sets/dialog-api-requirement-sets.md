---
title: ダイアログ API の要件セット
description: ダイアログ API の要件セットの詳細について説明します。
ms.date: 09/14/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 79b6960387519ac3c8b41b0b31cf6f40b5e7e067
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771362"
---
# <a name="dialog-api-requirement-sets"></a>ダイアログ API の要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office アプリケーションがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。

Office アドインは Office の複数のバージョンで機能します。 次の表に、ダイアログ API の要件セット、その要件セットをサポートする Office クライアント アプリケーション、およびアプリケーションのビルド番号またはバージョン番号Officeします。

|  要件セット  | Windows 版 Office 2013\*<br>(1 回限りの購入) | Office 2016 以降 (Windows)\*<br>(1 回限りの購入)   | Windows での Office<br>(サブスクリプション) |  Office on iPad<br>(サブスクリプション)  |  Office on Mac<br>(サブスクリプション)  | Office on the web  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| DialogApi 1.2  | N/A | N/A | サポートを参照する<br>セクション | 2.67 以降 | 16.37 以降 | 2020 年 6 月 | N/A |
| DialogApi 1.1  | ビルド 15.0.4855.1000 以降 | ビルド 16.0.4390.1000 以降 | バージョン 1602 (ビルド 6741.0000) 以降 | 1.22 以降 | 15.20 以降 | 2017 年 1 月 | バージョン 1608 (ビルド 7601.6800) 以降|

>\* 1 回の購入のユーザー Office、一部の修正プログラムと更新プログラムを受け入れてない場合があります。 その場合、DialogApi をサポートするために必要な更新された DLL がユーザーのコンピューターにインストールされていない場合でも、ui でバージョンを報告するために Office が使用する DLL は、ここに示されているバージョンよりも大きい可能性があります。 必要な修正プログラムがインストールされていることを確認するには、ユーザーは Office 更新リスト ([Office 2013](/officeupdates/msp-files-office-2013) リストまたは [Office 2016](/officeupdates/msp-files-office-2016)リスト) に移動し **、osfclient-x-none** を検索して、一覧に示されている修正プログラムをインストールする必要があります。

## <a name="office-on-windows-subscription-support"></a>Office Windows (サブスクリプション) のサポート

DialogApi 1.2 要件セットは、コンシューマー チャネル バージョン 2005 (ビルド、12827.20268 以上) でサポートされています。 For Office on Windows, the feature is also supported in the Semi-Annual Channel and Monthly Enterprise Channel builds available June 9th, 2020 or later. 各チャネルでサポートされる最小ビルドは次のとおりです。  

|チャネル | バージョン | ビルド|
|:-----|:-----|:-----|
|最新チャネル | 2005 以上 | 12827.20160 以上|
|月次エンタープライズ チャネル | 2004 以上 | 12730.20430 以上|
|半期エンタープライズ チャネル | 2002 以上 | 12527.20720 以上|

## <a name="office-versions-and-build-numbers"></a>Office のバージョンとビルド番号

バージョン、ビルド番号、Office Online Server の詳細については以下を参照してください。

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server 概要](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。

## <a name="dialog-api-11-and-12"></a>ダイアログ API 1.1 および 1.2

ダイアログ API 1.1 は、API の最初のバージョンです。 要件セット 1.2 では、親ページから [Office.dialog.messageChild](/javascript/api/office/office.dialog#messageChild_message_) メソッドを使用してデータを送信するサポートがダイアログ ボックスに追加されます。 これらの API の詳細については、ダイアログ API のリファレンス [トピックを](/javascript/api/office/office.ui) 参照してください。

## <a name="see-also"></a>関連項目

- [Office アドインで Office ダイアログ API を使用する](../../develop/dialog-api-in-office-add-ins.md)
- [Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)
- [Office アプリケーションと API 要件を指定する](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office アドインの XML マニフェスト](../../develop/add-in-manifests.md)
