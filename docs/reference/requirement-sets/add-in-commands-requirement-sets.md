---
title: アドイン コマンドの要件セット
description: Office アドインコマンドの要件セットの概要
ms.date: 03/11/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: d1904d092988a445be3e481123eecbad39097764
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094415"
---
# <a name="add-in-commands-requirement-sets"></a>アドイン コマンドの要件セット

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).

Add-in commands are UI elements that extend the Office UI and start actions in your add-in. You can use add-in commands to add a button on the ribbon or an item to a context menu. For more information, see [Add-in commands for Excel, Word, and PowerPoint](../../design/add-in-commands.md) and [Add-in commands for Outlook](../../outlook/add-in-commands-for-outlook.md).

The initial release of add-in commands doesn't have a corresponding requirement set (that is, there isn't an AddinCommands 1.0 requirement set). The following table lists the Office host applications that support the initial release version, and the build versions or number for those applications.  

| リリース   |  Windows 版 Office 2013<br>(1 回限りの購入) | Windows 版 Office 2016<br>(1 回限りの購入) | Windows 版 Office 2019<br>(1 回限りの購入) | Windows での Office<br>(Microsoft 365 サブスクリプションに接続)   |  iPad 上の Office<br>(Microsoft 365 サブスクリプションに接続)  |  Mac 上の Office<br>(Microsoft 365 サブスクリプションに接続)  | Office on the web  |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| アドイン コマンド (初期リリース、要件設定なし) | N/A | 16.0.4678.1000 *Outlook でのみサポートされています* | バージョン 1809 (ビルド 10827.20150) 以降 |バージョン 1603 (ビルド 6769.0000) 以降 | 該当なし | 15.33 以降| 2016 年 1 月 |

アドイン コマンド 1.1 の要件セットでは、「[ドキュメントで作業ウィンドウを自動的に開く](../../develop/automatically-open-a-task-pane-with-a-document.md)」機能が導入されています。

次の表に、アドイン コマンド 1.1 の要件セット、その要件セットをサポートする Office ホスト アプリケーション、Office アプリケーションのビルド番号またはバージョン番号を示します。

|  要件セット  |  Windows 版 Office 2013<br>(1 回限りの購入) | Windows 版 Office 2016<br>(1 回限りの購入) | Windows 版 Office 2019<br>(1 回限りの購入) | Windows での Office<br>(Microsoft 365 サブスクリプションに接続)   |  iPad 上の Office<br>(Microsoft 365 サブスクリプションに接続)  |  Mac 上の Office<br>(Microsoft 365 サブスクリプションに接続)  | Office on the web  |  
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
