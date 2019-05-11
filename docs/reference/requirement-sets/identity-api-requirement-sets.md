---
title: Identity API の要件セット
description: ''
ms.date: 05/08/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: c531fd6f7c77a02402bdd70b4f2cb091787a6a52
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952188"
---
# <a name="identity-api-requirement-sets"></a>Identity API の要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)」をご覧ください。

Office アドインは Office の複数のバージョンで機能します。 次の表は、Identity API の要件セット、その要件セットをサポートする Office ホスト アプリケーション、Office アプリケーションのビルド番号またはバージョン番号の一覧です。

|  要件セット  | Office 2013 以降 (Windows)<br>(1 回限りの購入) | Windows 版 Office<br>(Office 365 に接続) |  Office for iPad<br>(Office 365 に接続)  |  Office for Mac<br>(Office 365 に接続)  | Office Online  | SharePoint Online | OneDrive.com |Outlook.com および Exchange Online|
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1.1  | N/A | プレビュー<b>*</b> | 近日対応予定 | プレビュー<b>*</b> | プレビュー<b>*</b> | プレビュー<b>*</b>| 近日公開 | 近日公開 |

> **& # 42;** プレビューフェーズでは、Id API に Office 365 (サブスクリプション版の Office) が必要です。 Insider チャネルからの最新の月次バージョンとビルドを使ってください。 このバージョンを入手するには、Office Insider への参加が必要です。 詳細については、「[Office Insider になる](https://products.office.com/office-insider?tab=tab-1)」を参照してください。 ビルドが半期チャネルの運用に移行すると、そのビルドで SSO を含むプレビュー機能のサポートはオフになりますので、ご注意ください。

バージョン、ビルド番号、および Office Online Server の詳細については以下を参照してください。

- [Office 365 クライアントの更新プログラム チャネル リリースのバージョン番号およびビルド番号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [使用している Office のバージョンを確認する方法](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Office 365 クライアント アプリケーションのバージョン番号およびビルド番号を確認することができます。](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server 概要](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

共通 API の要件セットの詳細については、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」をご覧ください。

## <a name="identityapi-11"></a>IdentityAPI 1.1

シングル サインオン IdentityAPI 1.1 は API の最初のバージョンです。 この API の詳細については、「[Office アドインのシングル サインオンを有効化する (プレビュー)](/office/dev/add-ins/develop/sso-in-office-add-ins)」の「[SSO API リファレンス](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)」のセクションを参照してください。

## <a name="see-also"></a>関連項目

- [Office のバージョンと要件セット](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Office のホストと API の要件を指定する](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office アドインの XML マニフェスト](/office/dev/add-ins/develop/add-in-manifests)
