---
title: アドイン コマンドの要件セット
description: ''
ms.date: 06/20/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 0d2385f32c314fccb60335d721dd8db3bb30393a
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950972"
---
# <a name="add-in-commands-requirement-sets"></a>アドイン コマンドの要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)」をご覧ください。

アドイン コマンドは、Office UI を拡張し、アドインでアクションを開始する UI 要素です。アドイン コマンドを使用すると、リボン上のボタンやアイテムをコンテキスト メニューに追加できます。詳細については、「[Excel、Word および PowerPoint のアドイン コマンド](/office/dev/add-ins/design/add-in-commands)」と「[Outlook のアドイン コマンド](/outlook/add-ins/add-in-commands-for-outlook)」を参照してください。

アドイン コマンドの最初のリリースには、対応する要件セットがありません (つまり、AddinCommands 1.0 要件セットはありません)。次の表に、初期リリースのバージョンをサポートする Office ホスト アプリケーション、およびそれらのアプリケーションのビルド バージョンまたはビルド番号を示します。  

| リリース   |  Windows 版 Office 2013<br>(1 回限りの購入) | Windows 版 Office 2016<br>(1 回限りの購入) | Windows 版 Office 2019<br>(1 回限りの購入) | Windows での Office<br>(Office 365 サブスクリプションに接続)   |  Office on iPad<br>(Office 365 サブスクリプションに接続)  |  Office on Mac<br>(Office 365 サブスクリプションに接続)  | Office on the web  |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| アドイン コマンド (初期リリース、要件設定なし) | N/A | 16.0.4678.1000 *Outlook でのみサポートされています* | バージョン 1809 (ビルド 10827.20150) 以降 |バージョン 1603 (ビルド 6769.0000) 以降 | 該当なし | 15.33 以降| 2016 年 1 月 |

アドイン コマンド 1.1 の要件セットでは、「[ドキュメントで作業ウィンドウを自動的に開く](/office/dev/add-ins/develop/automatically-open-a-task-pane-with-a-document)」機能が導入されています。

次の表に、アドイン コマンド 1.1 の要件セット、その要件セットをサポートする Office ホスト アプリケーション、Office アプリケーションのビルド番号またはバージョン番号を示します。

|  要件セット  |  Windows 版 Office 2013<br>(1 回限りの購入) | Windows 版 Office 2016<br>(1 回限りの購入) | Windows 版 Office 2019<br>(1 回限りの購入) | Windows での Office<br>(Office 365 サブスクリプションに接続)   |  Office on iPad<br>(Office 365 サブスクリプションに接続)  |  Office on Mac<br>(Office 365 サブスクリプションに接続)  | Office on the web  |  
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| AddInCommands 1.1  | N/A | 16.0.4678.1000 *Outlook でのみサポートされています*  | バージョン 1809 (ビルド 10827.20150) 以降 | バージョン 1705 (ビルド 8121.1000) 以降 | N/A | 15.34 以降\*| 2017 年 5 月 |

>\* [Office.context.requirements.isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) メソッドはバージョン 16.9 &ndash; 16.14 (バージョン 16.9、16.14 も含む) で `false` を返しますが、これは間違っており、要件セットはこれらのバージョンでサポートされて*います*。

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
