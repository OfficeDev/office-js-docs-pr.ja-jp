---
title: アドイン コマンドの要件セット
description: Office アドインコマンドの要件セットの概要。
ms.date: 11/01/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 08fcb5df0e614e9b9f3ec9479fc958cc79adf320
ms.sourcegitcommit: 3189c4bd62dbe5950b19f28ac2c1314b6d304dca
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/17/2020
ms.locfileid: "49087960"
---
# <a name="add-in-commands-requirement-sets"></a>アドイン コマンドの要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office アプリケーションがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。

アドイン コマンドは、Office UI を拡張し、アドインでアクションを開始する UI 要素です。アドイン コマンドを使用すると、リボン上のボタンやアイテムをコンテキスト メニューに追加できます。詳細については、「[Excel、Word および PowerPoint のアドイン コマンド](../../design/add-in-commands.md)」と「[Outlook のアドイン コマンド](../../outlook/add-in-commands-for-outlook.md)」を参照してください。

アドインコマンドの最初のリリースには、対応する要件セットがありません (つまり、AddinCommands 1.0 の要件セットはありません)。 次の表に、最初のリリースバージョンをサポートする Office クライアントアプリケーションと、それらのアプリケーションのビルドバージョンまたはバージョン番号を示します。  

| リリース   |  Windows 版 Office 2013<br>(1 回限りの購入) | Windows 版 Office 2016<br>(1 回限りの購入) | Windows 版 Office 2019<br>(1 回限りの購入) | Windows での Office<br>(Microsoft 365 サブスクリプションに接続)   |  Office on iPad<br>(Microsoft 365 サブスクリプションに接続)  |  Office on Mac<br>(Microsoft 365 サブスクリプションに接続)  | Office on the web  |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| アドイン コマンド (初期リリース、要件設定なし) | N/A | 16.0.4678.1000 *Outlook でのみサポートされています* | バージョン 1809 (ビルド 10827.20150) 以降 |バージョン 1603 (ビルド 6769.0000) 以降 | 該当なし | 15.33 以降| 2016 年 1 月 |

**1.1** のアドインコマンドの要件セットには、[ドキュメントを使用して作業ウィンドウを autoopen](../../develop/automatically-open-a-task-pane-with-a-document.md)にする機能が導入されています。

「アドインコマンド **1.3** の要件セット」では、アドインを使用して office リボンのカスタムタブの配置をカスタマイズしたり、組み込みの office リボンコントロールをカスタムコントロールグループに挿入したりするためのマニフェストマークアップについて説明します。

次の表に、アドインコマンドの要件セット、その要件セットをサポートする Office クライアントアプリケーション、Office アプリケーションのビルド番号またはバージョン番号を示します。

|  要件セット  |  Windows 版 Office 2013<br>(1 回限りの購入) | Windows 版 Office 2016<br>(1 回限りの購入) | Windows 版 Office 2019<br>(1 回限りの購入) | Windows での Office<br>(Microsoft 365 サブスクリプションに接続)   |  Office on iPad<br>(Microsoft 365 サブスクリプションに接続)  |  Office on Mac<br>(Microsoft 365 サブスクリプションに接続)  | Office on the web  |  
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| AddinCommands 1.3  | N/A | N/A  | N/A | 近日公開 | N/A | 近日公開 | 2020 年 11 月 |
| AddInCommands 1.1  | N/A | 16.0.4678.1000 *Outlook でのみサポートされています*  | バージョン 1809 (ビルド 10827.20150) 以降 | バージョン 1705 (ビルド 8121.1000) 以降 | N/A | 15.34 以降\*| 2017 年 5 月 |

>\* [Office.context.requirements.isSetSupported](/javascript/api/office/office.requirementsetsupport#issetsupported-name--minversion-) メソッドはバージョン 16.9 &ndash; 16.14 (バージョン 16.9、16.14 も含む) で `false` を返しますが、これは間違っており、要件セットはこれらのバージョンでサポートされて *います*。

> [!IMPORTANT]
> AddinCommands 1.3 はプレビュー段階であり、 *web 上の PowerPoint でのみ使用でき* ます。 テストおよび開発環境でマークアップを試すことをお勧めします。 運用環境または業務上重要なドキュメント内では、プレビューマークアップを使用しないでください。

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
