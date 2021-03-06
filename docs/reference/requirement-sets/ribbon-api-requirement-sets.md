---
title: リボン API の要件セット
description: 動的リボン API Officeサポートするプラットフォームとビルドを指定します。
ms.date: 05/12/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: a608eff12fb21d7a4a6beb195749141bd473aa1c
ms.sourcegitcommit: 30f6c620380075e3459cac748ca0c656427b384d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/12/2021
ms.locfileid: "52330186"
---
# <a name="ribbon-api-requirement-sets"></a>リボン API の要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office アプリケーションがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。

リボン API セットは、カスタム アドイン コマンド (カスタム リボン ボタンとメニュー項目) が有効または無効になっている場合のプログラムによる制御をサポートします。

Office アドインは Office の複数のバージョンで機能します。 次の表に、リボン API 要件セット、その要件セットをサポートする Office クライアント アプリケーション、およびアプリケーションのビルド番号またはバージョン番号をOfficeします。

|  要件セット  | Windows 版 Office 2013<br>(1 回限りの購入) | Office 2016 以降のWindows<br>(1 回限りの購入)   | Windows での Office\*<br>(Microsoft 365 サブスクリプションに接続) |  Office on iPad<br>(Microsoft 365 サブスクリプションに接続)  |  Office on Mac\*<br>(Microsoft 365 サブスクリプションに接続)  | Office on the web\*  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| RibbonAPI 1.1  | 該当なし | 該当なし | サポートを見る<br>下のセクション | 該当なし | 16.38 | 2020 年 11 月 | 該当なし|
| RibbonApi 1.2  | 該当なし | 該当なし | 2102 (ビルド 13801.20294) | 該当なし | 近日公開 | 2021 年 5 月 | 該当なし|

> **&#42;** リボン API は、サブスクリプションでのみサポートExcel、サブスクリプションをMicrosoft 365します。

## <a name="support-for-version-11-on-office-on-windows-subscription"></a>バージョン 1.1 on Office (サブスクリプション) Windowsサポート

RibbonApi 要件セットの 1.1 バージョンは、コンシューマー チャネル バージョン 2006 (ビルド 13001.20498 以上) でサポートされています。 Office Windows、この機能は 2020 年 7 月 14 日以降に利用可能な Semi-Annual チャネルおよび月次 Enterprise チャネル ビルドでもサポートされます。 各チャネルでサポートされる最小ビルドは次のとおりです。  

|チャネル | バージョン | ビルド|
|:-----|:-----|:-----|
|最新チャネル | 2006 以上 | 20266.20266 以上|
|月次エンタープライズ チャネル | 2005 以上 | 12827.20538 以上|
|月次エンタープライズ チャネル | 2004 | 12730.20602 以上|
|半期エンタープライズ チャネル | 2002 以上 | 12527.20880 以上|

## <a name="more-information"></a>詳細情報

バージョン、ビルド番号、Office Online Server の詳細については以下を参照してください。

- [クライアントの更新チャネル リリースのバージョン番号とビルド番号Microsoft 365します。](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [使用している Office のバージョンを確認する方法](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [クライアント アプリケーションのバージョンとビルド番号をMicrosoft 365場所](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server 概要](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。

## <a name="ribbon-api-11"></a>リボン API 1.1

リボン API 1.1 は、API の最初のバージョンです。 API の詳細については[、「Office.ribbon リファレンス」を](/javascript/api/office/office.ribbon)参照してください。

## <a name="ribbon-api-12"></a>リボン API 1.2

リボン API 1.2 では、コンテキスト タブのサポートが追加されます。 詳細については、「[Office アドインでカスタム コンテキスト タブを作成する (プレビュー)](../../design/contextual-tabs.md)」を参照してください。

> [!NOTE]
> RibbonApi **1.2** 要件セットはマニフェストでまだサポートされていないので、マニフェストのセクションで指定 `<Requirements>` しないでください。

## <a name="see-also"></a>関連項目

- [Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)
- [Office アプリケーションと API 要件を指定する](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office アドインの XML マニフェスト](../../develop/add-in-manifests.md)
