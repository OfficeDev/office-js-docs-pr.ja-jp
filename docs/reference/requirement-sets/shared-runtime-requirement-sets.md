---
title: 共有ランタイム要件セット
description: SharedRuntime Api をサポートするプラットフォームと Office ホストを指定します。
ms.date: 02/11/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: dbb9d908154da074eaff6901c778adea168504a9
ms.sourcegitcommit: 7464eac3b54a6a6b65e27549a3ad603af6ee1011
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/27/2020
ms.locfileid: "42315883"
---
# <a name="shared-runtime-requirement-sets"></a>共有ランタイム要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)」をご覧ください。

JavaScript コードを実行する Office アドインの部分 (作業ウィンドウ、アドインコマンドから起動される関数ファイル、Excel カスタム関数) は、1つの JavaScript ランタイムを共有できます。 これにより、すべてのパーツが一連のグローバル変数を共有し、読み込まれたライブラリセットを共有して、永続的なストレージを介してメッセージを渡さずに相互に通信できるようになります。

次の表に、SharedRuntime 1.1 の要件セット、その要件セットをサポートする Office ホストアプリケーション、Office アプリケーションのビルド番号またはバージョン番号を示します。

|  要件セット  |  Windows での Office 2013 (またはそれ以降のバージョン)<br>(1 回限りの購入) | Windows での Office<br>(Office 365 サブスクリプションに接続)   |  Office on iPad<br>(Office 365 サブスクリプションに接続)  |  Office on Mac<br>(Office 365 サブスクリプションに接続)  | Office on the web  | Office Online Server |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| SharedRuntime 1.1  | 該当なし | バージョン 2002 (ビルド 12527.20092) 以降 | 該当なし | 16.35 以降 | 2020 年 2 月 | 該当なし |

バージョン、ビルド番号、Office Online Server の詳細については以下を参照してください。

- [Office 365 クライアントの更新プログラム チャネル リリースのバージョン番号およびビルド番号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [使用している Office のバージョンを確認する方法](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Office 365 クライアント アプリケーションのバージョン番号およびビルド番号を確認することができます。](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server 概要](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。

## <a name="see-also"></a>関連項目

- [Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Office のホストと API の要件を指定する](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office アドインの XML マニフェスト](/office/dev/add-ins/develop/add-in-manifests)
