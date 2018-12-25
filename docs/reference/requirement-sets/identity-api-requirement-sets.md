---
title: Identity API の要件セット
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 43a220cfada5883f292edd13cc753dc6c70e3504
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433923"
---
# <a name="identity-api-requirement-sets"></a>Identity API の要件セット

要件セットは、API メンバーの名前付きグループです。Office アドインは、マニフェストで指定されている要件セットを使用するか、ランタイム チェックを使用して、Office ホストがアドインに必要な API をサポートしているかどうかを判別します。詳しくは、「[Office のバージョンと要件セット](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)」をご覧ください。

Office アドインは Office の複数のバージョンで機能します。 次の表は、Identity API の要件セット、その要件セットをサポートする Office ホスト アプリケーション、Office アプリケーションのビルド番号またはバージョン番号の一覧です。

|  要件セット  | Office 2013 for Windows | Office 365 for Windows   |  Office 365 for iPad  |  Office 365 for Mac  | Office Online  | SharePoint Online | OneDrive.com |Outlook.com および Exchange Online|
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1.1  | 該当なし | プレビュー **&#42;** | 間もなく提供開始 | プレビュー **&#42;**| プレビュー | プレビュー| 間もなく提供開始 | 間もなく提供開始 |

> **&#42;** プレビュー段階では、Identity API は Windows 2016 および Mac で、ファースト オプションを使用する Insider プログラムのユーザーに対してのみサポートされます。 Insider プログラムに参加するには、「[Office Insider に登録する](https://products.office.com/office-insider?tab=tab-1)」を参照してください。 ファースト トラックに切り替えるには、「[Insider ファースト](https://answers.microsoft.com/ja-JP/msoffice/forum/msoffice_officeinsider-mso_win10-msoinsider_reg/its-here-office-insider-fast-for-office-2016-on/dbe8e7bb-9523-44a4-948b-9436fedfd961)」を参照してください。

バージョン、ビルド番号、Office Online Server の詳細については以下を参照してください。

- [Office 365 クライアントの更新プログラム チャネル リリースのバージョン番号およびビルド番号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [使用している Office のバージョンを確認する方法](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Office 365 クライアント アプリケーションのバージョン番号およびビルド番号を確認することができます。](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server 概要](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

共通 API の要件セットの詳細は、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」を参照してください。

## <a name="identityapi-11"></a>IdentityAPI 1.1 

シングル サインオン IdentityAPI 1.1 は API の最初のバージョンです。 この API の詳細については、「[Office アドインのシングル サインオンを有効化する (プレビュー)](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins)」の「[SSO API リファレンス](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)」のセクションを参照してください。

## <a name="see-also"></a>関連項目

- [Office のバージョンと要件セット](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Office のホストと API の要件を指定する](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office アドインの XML マニフェスト](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
