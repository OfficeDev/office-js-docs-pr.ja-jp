---
title: ダイアログ API の要件セット
description: ''
ms.date: 03/11/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: daa16235101a32adda8d462056a2db94a0da2d4f
ms.sourcegitcommit: 05b73cdec5f4db7f0b8d48a5a552ee296a0332ca
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/11/2020
ms.locfileid: "42600719"
---
# <a name="dialog-api-requirement-sets"></a>ダイアログ API の要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。

Office アドインは Office の複数のバージョンで機能します。次の表は、ダイアログ API の要件セット、その要件セットをサポートする Office ホスト アプリケーション、Office アプリケーションのビルド番号またはバージョン番号の一覧です。

|  要件セット  | Windows 版 Office 2013\*<br>(1 回限りの購入) | Office 2016 以降 (Windows)\*<br>(1 回限りの購入)   | Windows での Office<br>(Office 365 サブスクリプションに接続) |  Office on iPad<br>(Office 365 サブスクリプションに接続)  |  Office on Mac<br>(Office 365 サブスクリプションに接続)  | Office on the web  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| DialogApi 1.1  | ビルド 15.0.4855.1000 以降 | ビルド 16.0.4390.1000 以降 | バージョン 1602 (ビルド 6741.0000) 以降 | 1.22 以降 | 15.20 以降| 2017 年 1 月 | バージョン 1608 (ビルド 7601.6800) 以降|

>\*ワンタイム購入オフィスのユーザーは、すべての修正プログラムと更新を承諾していない場合があります。 その場合、Office が UI でそのバージョンを報告するために使用する DLL が、ユーザーのコンピューターにインストールされていない更新された Dll がインストールされていない場合でも、ここにリストされているバージョンよりも大きくなる可能性があります。 必要な修正プログラムがインストールされていることを確認するには、ユーザーは Office 更新プログラムの一覧 ([office 2013 リスト](/officeupdates/msp-files-office-2013)または[office 2016 の一覧](/officeupdates/msp-files-office-2016)) に移動し、 **osfclient**を検索して、一覧に記載されている修正プログラムをインストールする必要があります。

## <a name="office-versions-and-build-numbers"></a>Office のバージョンとビルド番号

バージョン、ビルド番号、Office Online Server の詳細については以下を参照してください。

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server 概要](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。

## <a name="dialog-api-11"></a>ダイアログ API 1.1

ダイアログ API 1.1 は、API の最初のバージョンです。 API の詳細については、「 [DIALOG api](/javascript/api/office/office.ui)リファレンス」を参照してください。

## <a name="see-also"></a>関連項目

- [Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)
- [Office のホストと API の要件を指定する](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office アドインの XML マニフェスト](../../develop/add-in-manifests.md)
