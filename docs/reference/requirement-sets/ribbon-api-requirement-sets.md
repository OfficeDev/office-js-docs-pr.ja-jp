---
title: リボン API の要件セット
description: 動的リボン Api をサポートしている Office プラットフォームとビルドを指定します。
ms.date: 07/07/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 6a0e6af3a74b0b0402710fd66bac6c915aa4c18a
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094282"
---
# <a name="ribbon-api-requirement-sets"></a>リボン API の要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)」をご覧ください。

リボン API セットは、カスタムアドインコマンド (つまり、カスタムリボンボタンとメニュー項目) を有効または無効にするときのプログラムによる制御をサポートします。

Office アドインは Office の複数のバージョンで機能します。 次の表に、リボン API 要件セット、その要件セットをサポートする Office ホストアプリケーション、Office アプリケーションのビルド番号またはバージョン番号の一覧を示します。

|  要件セット  | Windows 版 Office 2013<br>(1 回限りの購入) | Office 2016 以降 (Windows)<br>(1 回限りの購入)   | Windows での Office\*<br>(Microsoft 365 サブスクリプションに接続) |  iPad 上の Office<br>(Microsoft 365 サブスクリプションに接続)  |  Mac 上の Office\*<br>(Microsoft 365 サブスクリプションに接続)  | Office on the web\*  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| RibbonApi 1.1  | N/A | N/A | バージョン 2002 (ビルド 12527.20264) 以降 | 16.38 以降 | N/A | 2020 年 2 月 | N/A|

> **&#42;** プレビューフェーズでは、リボン API は Excel でのみサポートされており、Microsoft 365 サブスクリプションを必要とします。 Insider チャネルからの最新の月次バージョンとビルドを使ってください。 このバージョンを入手するには、Office Insider への参加が必要です。 詳細については、「[Office Insider になる](https://products.office.com/office-insider?tab=tab-1)」を参照してください。 ビルドが graduates の半期チャネルに対して実行されている場合、リボン API を含むプレビュー機能のサポートは、そのビルドに対して無効になっていることに注意してください。

バージョン、ビルド番号、Office Online Server の詳細については以下を参照してください。

- [Microsoft 365 クライアントの更新プログラムチャネルリリースのバージョン番号とビルド番号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [使用している Office のバージョンを確認する方法](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Microsoft 365 クライアントアプリケーションのバージョン番号とビルド番号を確認することができます。](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server 概要](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。

## <a name="ribbon-api-11"></a>リボン API 1.1

リボン API 1.1 は、API の最初のバージョンです。 API の詳細については、「 [Office. ribbon](/javascript/api/office/office.ribbon) reference」のトピックを参照してください。

## <a name="see-also"></a>関連項目

- [Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Office のホストと API の要件を指定する](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office アドインの XML マニフェスト](/office/dev/add-ins/develop/add-in-manifests)
