---
title: アドイン コマンドの要件セット
description: アドイン コマンドOfficeセットの概要。
ms.date: 03/12/2022
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: 799511ad85e8e04422cc52e38ffc2a4278410e4e
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745535"
---
# <a name="add-in-commands-requirement-sets"></a>アドイン コマンドの要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office アプリケーションがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。

アドイン コマンドは、Office UI を拡張し、アドインでアクションを開始する UI 要素です。アドイン コマンドを使用すると、リボン上のボタンやアイテムをコンテキスト メニューに追加できます。詳細については、「[Excel、Word および PowerPoint のアドイン コマンド](../../design/add-in-commands.md)」と「[Outlook のアドイン コマンド](../../outlook/add-in-commands-for-outlook.md)」を参照してください。

> [!NOTE]
> Outlookアドインはアドイン コマンドをサポートしていますが、Outlook でアドイン コマンドを有効にする API とマニフェスト要素はメールボックス [1.3](../objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md) 要件セットに含まれています。 AddinCommands 要件セットは、この要件にOutlook。

アドイン コマンドの最初のリリースには、対応する要件セットが含まれています (つまり、AddinCommands 1.0 要件セットは存在しない)。 次の表に、Officeバージョンをサポートするクライアント アプリケーションと、それらのアプリケーションのビルド バージョンまたは番号を示します。  

| リリース   |  Windows 版 Office 2013<br>(1 回限りの購入) | Windows 版 Office 2016<br>(1 回限りの購入) | Windows 版 Office 2019<br>(1 回限りの購入) | Windows 版 Office 2021<br>(1 回限りの購入) | Windows での Office<br>(サブスクリプション)   |  Office on iPad<br>(サブスクリプション)  |  Office on Mac<br>(両方のサブスクリプション<br> Mac 2019 以降Office 1 回の購入)   | Office on the web  |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| アドイン コマンド (初期リリース、要件設定なし) | N/A | 該当なし | バージョン 1809 (ビルド 10827.20150) 以降| 16.0.14326.20454 以降 |バージョン 1603 (ビルド 6769.0000) 以降 | 該当なし | 15.33 以降| 2016 年 1 月 |

アドイン コマンド **1.1 要件** セットでは、ドキュメントを含む作業ウィンドウを自動開 [く機能が導入されています](../../develop/automatically-open-a-task-pane-with-a-document.md)。

アドイン コマンド **1.3** 要件セットでは、マニフェスト マークアップが導入され、アドインは Office リボン上のカスタム タブの配置をカスタマイズし、組み込みの Office リボン コントロールをカスタム コントロール グループに挿入できます。

次の表に、アドイン コマンド要件セット、その要件セットをサポートする Office クライアント アプリケーション、および Office アプリケーションのビルドまたはバージョン番号を示します。

|  要件セット  |  Windows 版 Office 2013<br>(1 回限りの購入) | Windows 版 Office 2016<br>(1 回限りの購入) | Windows 版 Office 2019<br>(1 回限りの購入) |  Windows 版 Office 2021<br>(1 回限りの購入) | Windows での Office<br>(サブスクリプション)   |  Office on iPad<br>(サブスクリプション)  |  Office on Mac<br>(両方のサブスクリプション<br> Mac 2019 以降Office 1 回の購入)   | Office on the web  |  
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| AddinCommands 1.3  | 該当なし | 該当なし | 該当なし | 該当なし | バージョン 2204 (ビルド 14827.10000) 以降 | 該当なし | 16.57.105.0 以降 | 2020 年 11 月 |
| AddInCommands 1.1  | N/A | 該当なし  | バージョン 1809 (ビルド 10827.20150) 以降&dagger; | 16.0.14326.20454 以降&dagger; | バージョン 1705 (ビルド 8121.1000) 以降&dagger; | 該当なし | 15.34 以降&dagger;\*| 2017 年 5 月 |

\* [Office.context.requirements.isSetSupported](/javascript/api/office/office.requirementsetsupport#office-office-requirementsetsupport-issetsupported-member(1)) メソッドはバージョン 16.9 &ndash; 16.14 (バージョン 16.9、16.14 も含む) で `false` を返しますが、これは間違っており、要件セットはこれらのバージョンでサポートされて *います*。

&dagger;OneNoteは、一部のユーザーでのみOffice on the web。

## <a name="office-versions-and-build-numbers"></a>Office のバージョンとビルド番号

バージョン、ビルド番号、Office Online Server の詳細については以下を参照してください。

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server 概要](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。

## <a name="see-also"></a>関連項目

- [Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)
- [Office アプリケーションと API 要件を指定する](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office アドインの XML マニフェスト](../../develop/add-in-manifests.md)
