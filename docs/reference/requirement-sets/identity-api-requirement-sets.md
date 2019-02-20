---
title: Identity API の要件セット
description: ''
ms.date: 02/18/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: e0a04f084deb0b366600b4f2597658d1f54ab7f5
ms.sourcegitcommit: 03773fef3d2a380028ba0804739d2241d4b320e5
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/20/2019
ms.locfileid: "30091240"
---
# <a name="identity-api-requirement-sets"></a>Identity API の要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)」をご覧ください。

Office アドインは Office の複数のバージョンで機能します。 次の表は、Identity API の要件セット、その要件セットをサポートする Office ホスト アプリケーション、Office アプリケーションのビルド番号またはバージョン番号の一覧です。

|  要件セット  | Windows 版 Office 2013/2016/2019 | Office 365 for Windows   |  Office 365 for iPad  |  Office 365 for Mac  | Office Online  | SharePoint Online | OneDrive.com |Outlook.com および Exchange Online|
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1.1  | N/A **& # 42;** | プレビュー **&#42;** | 間もなく提供開始 | プレビュー **&#42;**| プレビュー | プレビュー| 間もなく提供開始 | 間もなく提供開始 |

> **& # 42;** プレビューフェーズでは、id API に office 365 (office のサブスクリプション版 ("クリックして実行" とも呼ばれます) が必要です。 Insider チャネルからの最新の月次バージョンとビルドを使ってください。 このバージョンを入手するには、Office Insider への参加が必要です。 詳細については、「[Office Insider になる](https://products.office.com/office-insider?tab=tab-1)」を参照してください。 ビルドが半期チャネルの運用に移行すると、そのビルドで SSO を含むプレビュー機能のサポートはオフになりますので、ご注意ください。

バージョン、ビルド番号、Office Online Server の詳細については以下を参照してください。

- [Office 365 クライアントの更新プログラム チャネル リリースのバージョン番号およびビルド番号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [使用している Office のバージョンを確認する方法](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Office 365 クライアント アプリケーションのバージョン番号およびビルド番号を確認することができます。](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server 概要](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。

## <a name="identityapi-11"></a>IdentityAPI 1.1 

シングル サインオン IdentityAPI 1.1 は API の最初のバージョンです。 この API の詳細については、「[Office アドインのシングル サインオンを有効化する (プレビュー)](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins)」の「[SSO API リファレンス](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)」のセクションを参照してください。

## <a name="see-also"></a>関連項目

- [Office のバージョンと要件セット](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Office のホストと API の要件を指定する](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office アドインの XML マニフェスト](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
