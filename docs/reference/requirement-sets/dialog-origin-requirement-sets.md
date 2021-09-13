---
title: ダイアログ配信元の要件セット
description: 詳細については、「Dialog Origin requirement sets」を参照してください。
ms.date: 07/22/2021
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: db97b19c0a23fa7dbd1b93e03ccd7a7b76317d7a
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154345"
---
# <a name="dialog-origin-requirement-sets"></a>ダイアログ配信元の要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office アプリケーションがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。

Office アドインは Office の複数のバージョンで機能します。 次の表に、Dialog Origin 要件セット、その要件セットをサポートする Office クライアント アプリケーション、およびアプリケーションのビルドまたはバージョン番号をOfficeします。

|  要件セット  | Windows 版 Office 2013<br>(1 回限りの購入) | Windows 版 Office 2016<br>(1 回限りの購入) | Office 2019 以降のWindows<br>(1 回限りの購入) | Windows での Office<br>(サブスクリプション) |  Office on iPad<br>(サブスクリプション)  |  Office on Mac<br>(サブスクリプション)  | Office on the web  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| DialogOrigin 1.1  | ビルド<br>15.0.5371.1000<br>以降 | ビルド<br>16.0.5200.1000<br>以降 | ビルド<br>TBD<br>以降 | TBD | 2.52 以降 | 16.52 以降 | 2021 年 7 月 | バージョン 2108<br>(ビルド 10377.1000)<br>以降 |

## <a name="office-versions-and-build-numbers"></a>Office のバージョンとビルド番号

バージョン、ビルド番号、Office Online Server の詳細については以下を参照してください。

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server 概要](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。

## <a name="dialog-origin-11"></a>Dialog Origin 1.1

Dialog Origin 1.1 は API の最初のバージョンです。 ダイアログとその親ページ間のクロスドメイン メッセージングのサポートを提供します。 これらの API の詳細については[、「Office.ui」を](/javascript/api/office/office.ui)参照してください。

## <a name="see-also"></a>関連項目

- [Office アドインで Office ダイアログ API を使用する](../../develop/dialog-api-in-office-add-ins.md)
- [Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)
- [Office アプリケーションと API 要件を指定する](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office アドインの XML マニフェスト](../../develop/add-in-manifests.md)
