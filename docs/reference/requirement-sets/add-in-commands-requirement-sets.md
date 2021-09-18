---
title: アドイン コマンドの要件セット
description: アドイン コマンドOfficeセットの概要。
ms.date: 09/08/2021
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: 73bedf79ff9698ed14ed0e17976a3c9e1602cc7e
ms.sourcegitcommit: 3fe9e06a52c57532e7968dc007726f448069f48d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/18/2021
ms.locfileid: "59443532"
---
# <a name="add-in-commands-requirement-sets"></a>アドイン コマンドの要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office アプリケーションがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。

アドイン コマンドは、Office UI を拡張し、アドインでアクションを開始する UI 要素です。アドイン コマンドを使用すると、リボン上のボタンやアイテムをコンテキスト メニューに追加できます。詳細については、「[Excel、Word および PowerPoint のアドイン コマンド](../../design/add-in-commands.md)」と「[Outlook のアドイン コマンド](../../outlook/add-in-commands-for-outlook.md)」を参照してください。

アドイン コマンドの最初のリリースには、対応する要件セットが含まれています (つまり、AddinCommands 1.0 要件セットは存在しない)。 次の表に、Officeバージョンをサポートするクライアント アプリケーションと、それらのアプリケーションのビルド バージョンまたは番号を示します。  

| リリース   |  Windows 版 Office 2013<br>(1 回限りの購入) | Windows 版 Office 2016<br>(1 回限りの購入) | Windows 版 Office 2019<br>(1 回限りの購入) | Windows での Office<br>(Microsoft 365 サブスクリプションに接続)   |  Office on iPad<br>(Microsoft 365 サブスクリプションに接続)  |  Office on Mac<br>(Microsoft 365 サブスクリプションに接続)  | Office on the web  |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| アドイン コマンド (初期リリース、要件設定なし) | N/A | 16.0.4678.1000 *Outlook でのみサポートされています* | バージョン 1809 (ビルド 10827.20150) 以降 |バージョン 1603 (ビルド 6769.0000) 以降 | 該当なし | 15.33 以降| 2016 年 1 月 |

アドイン コマンド **1.1 要件** セットでは、ドキュメントを含む作業ウィンドウを自動開く [機能が導入されています](../../develop/automatically-open-a-task-pane-with-a-document.md)。

アドイン コマンド **1.3** 要件セットでは、マニフェスト マークアップが導入され、アドインは Office リボン上のカスタム タブの配置をカスタマイズし、組み込みの Office リボン コントロールをカスタム コントロール グループに挿入できます。

次の表に、アドイン コマンドの要件セット、その要件セットをサポートする Office クライアント アプリケーション、および Office アプリケーションのビルドまたはバージョン番号を示します。

|  要件セット  |  Windows 版 Office 2013<br>(1 回限りの購入) | Windows 版 Office 2016<br>(1 回限りの購入) | Windows 版 Office 2019<br>(1 回限りの購入) | Windows での Office<br>(Microsoft 365 サブスクリプションに接続)   |  Office on iPad<br>(Microsoft 365 サブスクリプションに接続)  |  Office on Mac<br>(Microsoft 365 サブスクリプションに接続)  | Office on the web  |  
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| AddinCommands 1.3  | 該当なし | 該当なし  | 該当なし | サポートされていません | N/A | サポートされていません | 2020 年 11 月 |
| AddInCommands 1.1  | N/A | 16.0.4678.1000 *Outlook でのみサポートされています*  | バージョン 1809 (ビルド 10827.20150) 以降 | バージョン 1705 (ビルド 8121.1000) 以降 | N/A | 15.34 以降\*| 2017 年 5 月 |

>\* [Office.context.requirements.isSetSupported](/javascript/api/office/office.requirementsetsupport#isSetSupported_name__minVersion_) メソッドはバージョン 16.9 &ndash; 16.14 (バージョン 16.9、16.14 も含む) で `false` を返しますが、これは間違っており、要件セットはこれらのバージョンでサポートされて *います*。

> [!IMPORTANT]
> AddinCommands 1.3 はプレビュー中で、このページ *でのみPowerPoint on the web。* テスト環境と開発環境でのみマークアップを試することをお勧めします。 実稼働環境やビジネスクリティカルなドキュメント内でプレビュー マークアップを使用しない。

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
