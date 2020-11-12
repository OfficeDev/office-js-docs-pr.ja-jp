---
title: リボン API の要件セット
description: 動的リボン Api をサポートしている Office プラットフォームとビルドを指定します。
ms.date: 11/07/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 878670367b253fa7700434681244b43b9cfa36a7
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996516"
---
# <a name="ribbon-api-requirement-sets"></a>リボン API の要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office アプリケーションがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)」をご覧ください。

リボン API セットは、カスタムアドインコマンド (つまり、カスタムリボンボタンとメニュー項目) を有効または無効にするときのプログラムによる制御をサポートします。

Office アドインは Office の複数のバージョンで機能します。 次の表に、リボン API 要件セット、その要件セットをサポートする Office クライアントアプリケーション、Office アプリケーションのビルド番号またはバージョン番号の一覧を示します。

|  要件セット  | Windows 版 Office 2013<br>(1 回限りの購入) | Office 2016 以降 (Windows)<br>(1 回限りの購入)   | Windows での Office\*<br>(Microsoft 365 サブスクリプションに接続) |  Office on iPad<br>(Microsoft 365 サブスクリプションに接続)  |  Office on Mac\*<br>(Microsoft 365 サブスクリプションに接続)  | Office on the web\*  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| RibbonApi 1.1  | N/A | N/A | サポートを参照<br>セクション (下) | N/A | 16.38 | 2020年11月 | N/A|

> **&#42;** リボン API は Excel でのみサポートされており、Microsoft 365 サブスクリプションを必要とします。

## <a name="office-on-windows-subscription-support"></a>Office on Windows (サブスクリプション) のサポート

要件セットは、コンシューマ Channel バージョン 2006 (ビルド、13001.20498 以降) でサポートされています。 Windows 版 Office の場合、この機能は Semi-Annual チャネルでもサポートされており、毎月のエンタープライズチャネルビルドは2020年7月14日以降に利用可能になります。 各チャネルでサポートされている最小ビルドは次のとおりです。  

|チャネル | バージョン | ビルド|
|:-----|:-----|:-----|
|最新チャネル | 2006以上 | 20266.20266 以上|
|月次エンタープライズ チャネル | 2005以上 | 12827.20538 以上|
|月次エンタープライズ チャネル | 2004 | 12730.20602 以上|
|半期エンタープライズ チャネル | 2002以上 | 12527.20880 以上|

## <a name="more-information"></a>詳細情報

バージョン、ビルド番号、Office Online Server の詳細については以下を参照してください。

- [Microsoft 365 クライアントの更新プログラム チャネル リリースのバージョン番号およびビルド番号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [使用している Office のバージョンを確認する方法](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Microsoft 365 クライアントアプリケーションのバージョン番号とビルド番号を確認することができます。](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server 概要](/officeonlineserver/office-online-server-overview)

> [!NOTE]
> **Ribbonapi 1.1** 要件セットはマニフェストでまだサポートされていないため、マニフェストのセクションで指定することはできません `<Requirements>` 。


## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。

## <a name="ribbon-api-11"></a>リボン API 1.1

リボン API 1.1 は、API の最初のバージョンです。 API の詳細については、「 [Office. ribbon ](/javascript/api/office/office.ribbon) reference」のトピックを参照してください。

## <a name="see-also"></a>関連項目

- [Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Office アプリケーションと API 要件を指定する](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office アドインの XML マニフェスト](/office/dev/add-ins/develop/add-in-manifests)
