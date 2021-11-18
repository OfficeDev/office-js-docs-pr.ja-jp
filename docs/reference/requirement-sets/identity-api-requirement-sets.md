---
title: Identity API の要件セット
description: ID API 要件は、アドインOffice情報を設定します。
ms.date: 11/16/2021
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: d953e3ca2d135b96ab8b3219d9fe0f52fbda9d99
ms.sourcegitcommit: 6e6c4803fdc0a3cc2c1bcd275288485a987551ff
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/18/2021
ms.locfileid: "61066717"
---
# <a name="identity-api-requirement-sets"></a>ID API の要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office アプリケーションがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。

Office アドインは Office の複数のバージョンで機能します。 次の表に、Identity API 要件セット、その要件セットをサポートする Office クライアント アプリケーション、およびアプリケーションのビルドまたはバージョン番号をOfficeします。

|  要件セット  | Office 2021 以降のWindows<br>(1 回限りの購入) | Windows での Office<br>(Microsoft 365 サブスクリプションに接続) |  Office on iPad<br>(Microsoft 365 サブスクリプションに接続)  |  Office on Mac<br>(Microsoft 365 サブスクリプションに接続)  | Office on the web  |
|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1.3  | ビルド 16.0.14326.20454 以降 | バージョン 2008 (ビルド 13127.20000) 以降 | サポート対象外 | 16.40 以降 | Microsoft Office SharePoint OnlineとOneDrive\* |

\*現在、要件セットは、Office on the webおよびドキュメントから開いているドキュメントMicrosoft Office SharePoint OnlineサポートOneDrive。

## <a name="outlook-and-identity-api-requirement-sets"></a>Outlook ID API 要件セット

アドイン コードで Identity API セット 1.3 をOutlookするには、呼び出しでサポートされていないか確認します `isSetSupported('IdentityAPI', '1.3')` 。 アドインのマニフェストOutlook宣言はサポートされていません。 `undefined` ではないことを確認することで、API がサポートされているかどうかを判断することもできます。 詳細については、「[後続の要件セットからの API の使用](outlook-api-requirement-sets.md#using-apis-from-later-requirement-sets)」を参照してください。

> [!NOTE]
> イベント ベースのライセンス認証を使用する Outlook アドインでは[、OfficeRuntime.Auth](/javascript/api/office-runtime/officeruntime.auth)インターフェイスは Office バージョン 2108 (ビルド 14326.20258) 以降の Windows Office でサポートされます。 この[Office。Auth インターフェイスは](/javascript/api/office/office.auth)バージョン 2109 (ビルド 14425.10000) 以降でサポートされています。 バージョンに応じて詳細を確認するには[、Office 2021](/officeupdates/update-history-office-2021)または[Microsoft 365](/officeupdates/update-history-office365-proplus-by-date)の更新履歴ページと、Office クライアントのバージョンと更新チャネルを見つける方法を[参照してください](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19)。

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
