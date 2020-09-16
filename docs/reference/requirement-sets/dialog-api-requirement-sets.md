---
title: ダイアログ API の要件セット
description: ダイアログ API の要件セットの詳細については、「」を参照してください。
ms.date: 09/14/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: c30a463cc1a5043d7c86709978a47796f93c380e
ms.sourcegitcommit: ed2a98b6fb5b432fa99c6cefa5ce52965dc25759
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/16/2020
ms.locfileid: "47819715"
---
# <a name="dialog-api-requirement-sets"></a>ダイアログ API の要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office アプリケーションがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。

Office アドインは Office の複数のバージョンで機能します。 次の表に、ダイアログ API の要件セット、その要件セットをサポートする Office クライアントアプリケーション、Office アプリケーションのビルド番号またはバージョン番号を示します。

|  要件セット  | Windows 版 Office 2013\*<br>(1 回限りの購入) | Office 2016 以降 (Windows)\*<br>(1 回限りの購入)   | Windows での Office<br>認証 |  Office on iPad<br>認証  |  Office on Mac<br>認証  | Office on the web  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| Offapi 1.2  | 該当なし | 該当なし | サポートを参照<br>セクション (下) | 2.67 以降 | 16.37 以降 | 2020 年 6 月 | 該当せず |
| DialogApi 1.1  | ビルド 15.0.4855.1000 以降 | ビルド 16.0.4390.1000 以降 | バージョン 1602 (ビルド 6741.0000) 以降 | 1.22 以降 | 15.20 以降 | 2017 年 1 月 | バージョン 1608 (ビルド 7601.6800) 以降|

>\* ワンタイム購入オフィスのユーザーは、すべての修正プログラムと更新を承諾していない場合があります。 その場合、Office が UI でそのバージョンを報告するために使用する DLL が、ユーザーのコンピューターにインストールされていない更新された Dll がインストールされていない場合でも、ここにリストされているバージョンよりも大きくなる可能性があります。 必要な修正プログラムがインストールされていることを確認するには、ユーザーは Office 更新プログラムの一覧 ([office 2013 リスト](/officeupdates/msp-files-office-2013) または [office 2016 の一覧](/officeupdates/msp-files-office-2016)) に移動し、 **osfclient**を検索して、一覧に記載されている修正プログラムをインストールする必要があります。

## <a name="office-on-windows-subscription-support"></a>Office on Windows (サブスクリプション) のサポート

Errorapi 1.2 の要件セットは、コンシューマ Channel バージョン 2005 (build、12827.20268 以降) でサポートされています。 Windows 版 Office の場合、この機能は半期チャネルでもサポートされています。月間のエンタープライズチャネル構築は、2020年6月9日以降に利用可能になります。 各チャネルでサポートされている最小ビルドは次のとおりです。  

|チャネル | バージョン | ビルド|
|:-----|:-----|:-----|
|現在のチャネル | 2005以上 | 12827.20160 以上|
|月次エンタープライズ チャネル | 2004以上 | 12730.20430 以上|
|半期エンタープライズ チャネル | 2002以上 | 12527.20720 以上|

## <a name="office-versions-and-build-numbers"></a>Office のバージョンとビルド番号

バージョン、ビルド番号、Office Online Server の詳細については以下を参照してください。

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server 概要](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。

## <a name="dialog-api-11-and-12"></a>ダイアログ API 1.1 および1.2

ダイアログ API 1.1 は、API の最初のバージョンです。 要件セット1.2 は、親ページからメソッドを使用してダイアログボックスにデータを送信するためのサポートを追加し `Office.ui.messageChild` ます。 これらの Api の詳細については、「 [DIALOG api](/javascript/api/office/office.ui) リファレンス」を参照してください。

## <a name="see-also"></a>関連項目

- [Office アドインで Office ダイアログ API を使用する](../../develop/dialog-api-in-office-add-ins.md)
- [Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)
- [Office アプリケーションと API 要件を指定する](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office アドインの XML マニフェスト](../../develop/add-in-manifests.md)
