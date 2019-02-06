---
title: アドイン コマンドの要件セット
description: ''
ms.date: 11/21/2018
ms.prod: non-product-specific
localization_priority: Priority
ms.openlocfilehash: 952cff4877fcdd0fbdf9b380f9ae34e83e271a46
ms.sourcegitcommit: 33dcf099c6b3d249811580d67ee9b790c0fdccfb
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/05/2019
ms.locfileid: "29742353"
---
# <a name="add-in-commands-requirement-sets"></a>アドイン コマンドの要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)」をご覧ください。

アドイン コマンドは、Office UI を拡張し、アドインでアクションを開始する UI 要素です。アドイン コマンドを使用すると、リボン上のボタンやアイテムをコンテキスト メニューに追加できます。詳細については、「[Excel、Word および PowerPoint のアドイン コマンド](https://docs.microsoft.com/office/dev/add-ins/design/add-in-commands)」と「[Outlook のアドイン コマンド](https://docs.microsoft.com/outlook/add-ins/add-in-commands-for-outlook)」を参照してください。

アドイン コマンドの最初のリリースには、対応する要件セットがありません (つまり、AddinCommands 1.0 要件セットはありません)。次の表に、初期リリースのバージョンをサポートする Office ホスト アプリケーション、およびそれらのアプリケーションのビルド バージョンまたはビルド番号を示します。  

| リリース   |  Office 2013 for Windows | Office 2016 for Windows (非サブスクリプション) | Office 365 for Windows   |  Office 365 for iPad  |  Office 365 for Mac  | Office Online  |  
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| アドイン コマンド (初期リリース、要件設定なし) | N/A | 16.0.4678.1000 *Outlook でのみサポートされています* |バージョン 1603 (ビルド 6769.0000) 以降 | 該当なし | 15.33 以降| 2016 年 1 月 |

アドイン コマンド 1.1 の要件セットでは、「[ドキュメントで作業ウィンドウを自動的に開く](https://docs.microsoft.com/office/dev/add-ins/develop/automatically-open-a-task-pane-with-a-document)」機能が導入されています。

次の表に、アドイン コマンド 1.1 の要件セット、その要件セットをサポートする Office ホスト アプリケーション、Office アプリケーションのビルド番号またはバージョン番号を示します。 

|  要件セット  |  Office 2013 for Windows | Office 2016 for Windows (非サブスクリプション) | Office 365 for Windows   |  Office 365 for iPad  |  Office 365 for Mac  | Office Online  |  
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| AddInCommands 1.1  | N/A | 16.0.4678.1000 *Outlook でのみサポートされています*  | バージョン 1705 (ビルド 8121.1000) 以降 | N/A | 15.34 以降\*| 2017 年 5 月 |

>\* [Office.context.requirements.isSetSupported](https://docs.microsoft.com/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) メソッドはバージョン 16.9 &ndash; 16.14 (バージョン 16.9、16.14 も含む) で `false` を返しますが、これは間違っており、要件セットはこれらのバージョンでサポートされて*います*。

バージョン、ビルド番号、Office Online Server の詳細については以下を参照してください。

- [Office 365 クライアントの更新プログラム チャネル リリースのバージョン番号およびビルド番号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [使用している Office のバージョンを確認する方法](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Office 365 クライアント アプリケーションのバージョン番号およびビルド番号を確認することができます。](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server 概要](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。

## <a name="see-also"></a>関連項目

- [Office のバージョンと要件セット](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Office のホストと API の要件を指定する](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office アドインの XML マニフェスト](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
