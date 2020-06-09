---
title: アドイン コマンドの要件セット
description: Office アドインコマンドの要件セットの概要
ms.date: 03/11/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 5a979b5ca57cf1ddc8ebf021b72ca5fb8755a167
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612123"
---
# <a name="add-in-commands-requirement-sets"></a>アドイン コマンドの要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。

アドイン コマンドは、Office UI を拡張し、アドインでアクションを開始する UI 要素です。アドイン コマンドを使用すると、リボン上のボタンやアイテムをコンテキスト メニューに追加できます。詳細については、「[Excel、Word および PowerPoint のアドイン コマンド](../../design/add-in-commands.md)」と「[Outlook のアドイン コマンド](../../outlook/add-in-commands-for-outlook.md)」を参照してください。

アドイン コマンドの最初のリリースには、対応する要件セットがありません (つまり、AddinCommands 1.0 要件セットはありません)。次の表に、初期リリースのバージョンをサポートする Office ホスト アプリケーション、およびそれらのアプリケーションのビルド バージョンまたはビルド番号を示します。  

| リリース   |  Windows 版 Office 2013<br>(1 回限りの購入) | Windows 版 Office 2016<br>(1 回限りの購入) | Windows 版 Office 2019<br>(1 回限りの購入) | Windows での Office<br>(Office 365 サブスクリプションに接続済み)   |  Office on iPad<br>(Office 365 サブスクリプションに接続済み)  |  Office on Mac<br>(Office 365 サブスクリプションに接続)  | Office on the web  |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| アドイン コマンド (初期リリース、要件設定なし) | N/A | 16.0.4678.1000 *Outlook でのみサポートされています* | バージョン 1809 (ビルド 10827.20150) 以降 |バージョン 1603 (ビルド 6769.0000) 以降 | 該当なし | 15.33 以降| 2016 年 1 月 |

アドイン コマンド 1.1 の要件セットでは、「[ドキュメントで作業ウィンドウを自動的に開く](../../develop/automatically-open-a-task-pane-with-a-document.md)」機能が導入されています。

次の表に、アドイン コマンド 1.1 の要件セット、その要件セットをサポートする Office ホスト アプリケーション、Office アプリケーションのビルド番号またはバージョン番号を示します。

|  要件セット  |  Windows 版 Office 2013<br>(1 回限りの購入) | Windows 版 Office 2016<br>(1 回限りの購入) | Windows 版 Office 2019<br>(1 回限りの購入) | Windows での Office<br>(Office 365 サブスクリプションに接続済み)   |  Office on iPad<br>(Office 365 サブスクリプションに接続済み)  |  Office on Mac<br>(Office 365 サブスクリプションに接続)  | Office on the web  |  
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| AddInCommands 1.1  | N/A | 16.0.4678.1000 *Outlook でのみサポートされています*  | バージョン 1809 (ビルド 10827.20150) 以降 | バージョン 1705 (ビルド 8121.1000) 以降 | N/A | 15.34 以降\*| 2017 年 5 月 |

>\* [Office.context.requirements.isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) メソッドはバージョン 16.9 &ndash; 16.14 (バージョン 16.9、16.14 も含む) で `false` を返しますが、これは間違っており、要件セットはこれらのバージョンでサポートされて*います*。

## <a name="office-versions-and-build-numbers"></a>Office のバージョンとビルド番号

バージョン、ビルド番号、Office Online Server の詳細については以下を参照してください。

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server 概要](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。

## <a name="see-also"></a>関連項目

- [Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)
- [Office のホストと API の要件を指定する](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office アドインの XML マニフェスト](../../develop/add-in-manifests.md)
