---
title: Identity API の要件セット
description: ID API 要件は、アドインOffice情報を設定します。
ms.date: 09/08/2021
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: e3af8767666d3015894c0b7bcdecd758b1a1547c
ms.sourcegitcommit: 3fe9e06a52c57532e7968dc007726f448069f48d
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/19/2021
ms.locfileid: "59448755"
---
# <a name="identity-api-requirement-sets"></a>ID API の要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office アプリケーションがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。

Office アドインは Office の複数のバージョンで機能します。 次の表に、Identity API 要件セット、その要件セットをサポートする Office クライアント アプリケーション、およびアプリケーションのビルドまたはバージョン番号をOfficeします。

|  要件セット  | Office 2021 以降のWindows<br>(1 回限りの購入) | Windows での Office<br>(Microsoft 365 サブスクリプションに接続) |  Office on iPad<br>(Microsoft 365 サブスクリプションに接続)  |  Office on Mac<br>(Microsoft 365 サブスクリプションに接続)  | Office on the web  |
|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1.3  | 2008 (ビルド 13127.20000) 以降 | 2008 (ビルド 13127.20000) 以降 | 非サポート | 16.40 以降 | Microsoft Office SharePoint OnlineとOneDrive\* |

\*現在、要件セットは、Office on the webおよびドキュメントから開いているドキュメントMicrosoft Office SharePoint OnlineサポートOneDrive。

> [!NOTE]
> Outlook: アドイン コードで Identity API セット 1.3 を要求するには、呼び出しでサポートされていないか確認します `isSetSupported('IdentityAPI', '1.3')` 。 アドインのマニフェストOutlook宣言はサポートされていません。 `undefined` ではないことを確認することで、API がサポートされているかどうかを判断することもできます。 詳細については、「[後続の要件セットからの API の使用](outlook-api-requirement-sets.md#using-apis-from-later-requirement-sets)」を参照してください。

## <a name="office-versions-and-build-numbers"></a>Office のバージョンとビルド番号

バージョン、ビルド番号、Office Online Server の詳細については以下を参照してください。

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server 概要](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。

## <a name="identityapi-preview"></a>IdentityAPI プレビュー

この API の詳細については [、getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-) で Promises を使用するバージョン、または [getAccessTokenAsync](/javascript/api/office/office.auth#getAccessTokenAsync_options__callback_)でコールバックを使用するバージョンを参照してください。

## <a name="see-also"></a>関連項目

- [Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)
- [Office アプリケーションと API 要件を指定する](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office アドインの XML マニフェスト](../../develop/add-in-manifests.md)
