---
title: キーボード ショートカットの要件セット
description: キーボード ショートカットの要件は、Officeの情報を設定します。
ms.date: 02/15/2022
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: bf7cd3cb8e0a6054f3e279e148e4b47c480e28fb
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745905"
---
# <a name="keyboard-shortcuts-requirement-sets"></a>キーボード ショートカットの要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office アプリケーションがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。

Office アドインは Office の複数のバージョンで機能します。 次の表に、キーボード ショートカット要件セット、その要件セットをサポートする Office クライアント アプリケーション、および Office アプリケーションのビルドまたはバージョン番号を示します。

|  要件セット  | Windows での Office 2013 以降<br>(1 回限りの購入) | Windows での Office<br>(Microsoft 365 サブスクリプションに接続) |  Office on iPad<br>(Microsoft 365 サブスクリプションに接続)  |  Office on Mac<br>(両方のサブスクリプション<br> Mac 2019 以降Office 1 回の購入)   | Office on the web  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| KeyboardShortcuts 1.1  | 該当なし | バージョン: 2111 (ビルド 14701.10000) | 該当なし | 16.55 | 2021 年 9 月 |

> [!NOTE]
> **KeyboardShortcuts 1.1** 要件セットは、ユーザーのExcel。

## <a name="office-versions-and-build-numbers"></a>Office のバージョンとビルド番号

バージョン、ビルド番号、Office Online Server の詳細については以下を参照してください。

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server 概要](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。

## <a name="keyboardshortcuts-11"></a>KeyboardShortcuts 1.1

この要件セットの API の詳細については、「[Office.actions」を参照してください](/javascript/api/office/office.actions)。

## <a name="see-also"></a>関連項目

- [Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)
- [Office アプリケーションと API 要件を指定する](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office アドインの XML マニフェスト](../../develop/add-in-manifests.md)
