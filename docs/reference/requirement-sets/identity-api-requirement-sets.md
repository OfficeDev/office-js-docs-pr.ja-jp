---
title: Identity API の要件セット
description: Id API の要件 Office アドインの情報を設定します。
ms.date: 07/30/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: c2c6ea449cef08248a9ba79051b7c0c5f9baa600
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293542"
---
# <a name="identity-api-requirement-sets"></a>Identity API の要件セット

要件セットは、API メンバーの名前付きグループです。 Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイムチェックを使用して、Office アプリケーションがアドインに必要な Api をサポートしているかどうかを判断します。 詳細については、「 [Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」を参照してください。

Office アドインは Office の複数のバージョンで機能します。 次の表に、Id API の要件セット、その要件セットをサポートする Office クライアントアプリケーション、Office アプリケーションのビルド番号またはバージョン番号を示します。

|  要件セット  | Windows での Office 2013 以降<br>(1 回限りの購入) | Windows での Office<br>(Microsoft 365 サブスクリプションに接続) |  Office on iPad<br>(Microsoft 365 サブスクリプションに接続)  |  Office on Mac<br>(Microsoft 365 サブスクリプションに接続)  | Office on the web  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| Identity Api 1.3  | N/A | 2008 (ビルド 13127.20000) 以降 | 近日対応予定 | 16.40 以降 | 8月、2020 * |

> \* 最初は、web 上の Office で要件セットがサポートされているのは、SharePoint Online および OneDrive.com から開かれたドキュメントのみです。 他のドキュメントのサポートは、2020の後の方に web 上の Office に送られます。

## <a name="office-versions-and-build-numbers"></a>Office のバージョンとビルド番号

バージョン、ビルド番号、Office Online Server の詳細については以下を参照してください。

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server 概要](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。

## <a name="identityapi-preview"></a>Identity Api プレビュー

この API の詳細については、「 [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-) での約束を使用するバージョン」または [getAccessTokenAsync](/javascript/api/office/office.auth#getaccesstokenasync-options--callback-)でコールバックを使用するバージョンのいずれかを参照してください。

## <a name="see-also"></a>関連項目

- [Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)
- [Office アプリケーションと API の要件を指定する](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office アドインの XML マニフェスト](../../develop/add-in-manifests.md)
