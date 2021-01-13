---
title: リボン API の要件セット
description: 動的リボン API Officeサポートするプラットフォームとビルドを指定します。
ms.date: 11/07/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 91c909755779d122fba8d77dc246784f6a0dd1a3
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839986"
---
# <a name="ribbon-api-requirement-sets"></a>リボン API の要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office アプリケーションがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。

リボン API セットは、カスタム アドイン コマンド (つまり、カスタム リボン ボタンとメニュー項目) が有効または無効になっている場合のプログラムによる制御をサポートします。

Office アドインは Office の複数のバージョンで機能します。 次の表に、リボン API の要件セット、その要件セットをサポートする Office クライアント アプリケーション、およびアプリケーションのビルド番号またはバージョン番号をOfficeします。

|  要件セット  | Windows 版 Office 2013<br>(1 回限りの購入) | Office 2016 以降 (Windows)<br>(1 回限りの購入)   | Windows での Office\*<br>(Microsoft 365 サブスクリプションに接続) |  Office on iPad<br>(Microsoft 365 サブスクリプションに接続)  |  Office on Mac\*<br>(Microsoft 365 サブスクリプションに接続)  | Office on the web\*  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| RibbonApi 1.1  | N/A | N/A | サポートを参照する<br>セクション | N/A | 16.38 | 2020 年 11 月 | N/A|

> **&#42;** リボン API は Excel でのみサポートされ、Microsoft 365 サブスクリプションが必要です。

## <a name="office-on-windows-subscription-support"></a>Office Windows (サブスクリプション) のサポート

要件セットは、コンシューマー チャネル バージョン 2006 (ビルド、13001.20498 以上) でサポートされています。 For Office on Windows the feature is also supported in the Semi-Annual Channel and Monthly Enterprise Channel builds available July 14th, 2020 or later. 各チャネルでサポートされる最小ビルドは次のとおりです。  

|チャネル | バージョン | ビルド|
|:-----|:-----|:-----|
|最新チャネル | 2006 以上 | 20266.20266 以上|
|月次エンタープライズ チャネル | 2005 以上 | 12827.20538 以上|
|月次エンタープライズ チャネル | 2004 | 12730.20602 以上|
|半期エンタープライズ チャネル | 2002 以上 | 12527.20880 以上|

## <a name="more-information"></a>詳細

バージョン、ビルド番号、Office Online Server の詳細については以下を参照してください。

- [Microsoft 365 クライアントの更新プログラム チャネル リリースのバージョン番号とビルド番号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [使用している Office のバージョンを確認する方法](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Microsoft 365 クライアント アプリケーションのバージョンとビルド番号を確認できる場所](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server 概要](/officeonlineserver/office-online-server-overview)

> [!NOTE]
> **RibbonApi 1.1** 要件セットはマニフェストでまだサポートされていないので、マニフェストのセクションで指定 `<Requirements>` することはできません。


## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。

## <a name="ribbon-api-11"></a>リボン API 1.1

リボン API 1.1 は、API の最初のバージョンです。 API の詳細については [、Office.ribbon ](/javascript/api/office/office.ribbon) リファレンス トピックを参照してください。

## <a name="see-also"></a>関連項目

- [Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)
- [Office アプリケーションと API 要件を指定する](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office アドインの XML マニフェスト](../../develop/add-in-manifests.md)