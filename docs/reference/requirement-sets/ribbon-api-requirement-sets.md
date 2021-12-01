---
title: リボン API の要件セット
description: 動的リボン API Officeサポートするプラットフォームとビルドを指定します。
ms.date: 11/29/2021
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: 1801d95da8dd0b2b707e1237498db71ca81474b5
ms.sourcegitcommit: 5daf91eb3be99c88b250348186189f4dc1270956
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/01/2021
ms.locfileid: "61242048"
---
# <a name="ribbon-api-requirement-sets"></a>リボン API の要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office アプリケーションがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。

リボン API セットは、カスタム アドイン コマンド (カスタム リボン ボタンとメニューアイテム) の有効化と無効化、およびリボンにコンテキスト タブが表示される場合のプログラムによる制御をサポートします。

> [!NOTE]
> RibbonApi 要件セットは、作業ウィンドウ アドインでのみサポートされます。

Office アドインは Office の複数のバージョンで機能します。 次の表に、リボン API 要件セット、その要件セットをサポートする Office クライアント アプリケーション、およびアプリケーションのビルド番号またはバージョン番号をOfficeします。

|  要件セット  | Office 2021 以降のWindows<br>(1 回限りの購入) | Windows での Office\*<br>(Microsoft 365 サブスクリプションに接続) |  Office on iPad<br>(Microsoft 365 サブスクリプションに接続)  |  Office on Mac\*<br>(Microsoft 365 サブスクリプションに接続)  | Office on the web\*  |  Office Online Server  |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| RibbonApi 1.2  | ビルド 16.0.14326.20454 以降 | 2102 (ビルド 13801.20294) | N/A | サポートされていません | 2021 年 5 月 | N/A|
| RibbonApi 1.1  | ビルド 16.0.14326.20454 以降 | サポートを見る<br>下のセクション | N/A | 16.38 | 2020 年 11 月 | 該当なし|

> **&#42;** リボン API は、リボン API でのみサポートExcel。

## <a name="support-for-version-11-on-office-on-windows-subscription"></a>バージョン 1.1 on Office (サブスクリプション) Windowsサポート

RibbonApi 要件セットの 1.1 バージョンは、コンシューマー チャネル バージョン 2006 (ビルド 13001.20498 以上) でサポートされています。 Office Windows、この機能は 2020 年 7 月 14 日以降に利用可能な Semi-Annual チャネルおよび月次 Enterprise チャネル ビルドでもサポートされます。 各チャネルでサポートされる最小ビルドは次のとおりです。  

|チャネル | バージョン | ビルド|
|:-----|:-----|:-----|
|最新チャネル | 2006 以上 | 20266.20266 以上|
|月次エンタープライズ チャネル | 2005 以上 | 12827.20538 以上|
|月次エンタープライズ チャネル | 2004 | 12730.20602 以上|
|半期エンタープライズ チャネル | 2002 以上 | 12527.20880 以上|

## <a name="more-information"></a>詳細

バージョン、ビルド番号、Office Online Server の詳細については以下を参照してください。

- [クライアントの更新チャネル リリースのバージョン番号とビルド番号Microsoft 365します。](/officeupdates/update-history-microsoft365-apps-by-date)
- [使用している Office のバージョンを確認する方法](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19)
- [クライアント アプリケーションのバージョンとビルド番号をMicrosoft 365場所](/officeupdates/update-history-microsoft365-apps-by-date)
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
