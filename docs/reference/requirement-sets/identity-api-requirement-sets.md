---
title: Identity API の要件セット
description: Id API の要件 Office アドインの情報を設定します。
ms.date: 04/16/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 4552626d692b08bab65f866ab406988f5e88945a
ms.sourcegitcommit: 803587b324fc8038721709d7db5664025cf03c6b
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/17/2020
ms.locfileid: "43547242"
---
# <a name="identity-api-requirement-sets"></a>Identity API の要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)」をご覧ください。

Office アドインは Office の複数のバージョンで機能します。 次の表は、Identity API の要件セット、その要件セットをサポートする Office ホスト アプリケーション、Office アプリケーションのビルド番号またはバージョン番号の一覧です。

|  要件セット  | Windows での Office 2013 以降<br>(1 回限りの購入) | Windows での Office<br>(Office 365 サブスクリプションに接続済み) |  Office on iPad<br>(Office 365 サブスクリプションに接続済み)  |  Office on Mac<br>(Office 365 サブスクリプションに接続)  | Office on the web  | SharePoint Online | OneDrive.com |Outlook.com および Exchange Online|
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| Identity Api プレビュー  | N/A | プレビュー<b>*</b> | 近日対応予定 | プレビュー<b>*</b> | Preview<b>* &#8224;</b> | Preview<b>* &#8224;</b>| 近日公開 | 近日公開 |

> **&#42;** プレビューフェーズでは、Id API に Office 365 (サブスクリプション版の Office) が必要です。 Insider チャネルからの最新の月次バージョンとビルドを使ってください。 このバージョンを入手するには、Office Insider への参加が必要です。 詳細については、「[Office Insider になる](https://insider.office.com)」を参照してください。 ビルドが半期チャネルの運用に移行すると、そのビルドで SSO を含むプレビュー機能のサポートはオフになりますので、ご注意ください。
>
> **&#8224;** これらのプラットフォームで SSO Api を使用するアドインは、ユーザーのテナント管理者がアドインへの同意を付与されている場合にのみ機能します。 ユーザーが自分の Azure AD プロファイルに対しても同意を付与することはできません。

## <a name="office-versions-and-build-numbers"></a>Office のバージョンとビルド番号

バージョン、ビルド番号、Office Online Server の詳細については以下を参照してください。

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [Office Online Server 概要](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。

## <a name="identityapi-preview"></a>Identity Api プレビュー

この API の詳細については、「 [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-)での約束を使用するバージョン」または[getAccessTokenAsync](/javascript/api/office/office.auth#getaccesstokenasync-options--callback-)でコールバックを使用するバージョンのいずれかを参照してください。

## <a name="see-also"></a>関連項目

- [Office のバージョンと要件セット](../../develop/office-versions-and-requirement-sets.md)
- [Office のホストと API の要件を指定する](../../develop/specify-office-hosts-and-api-requirements.md)
- [Office アドインの XML マニフェスト](../../develop/add-in-manifests.md)
