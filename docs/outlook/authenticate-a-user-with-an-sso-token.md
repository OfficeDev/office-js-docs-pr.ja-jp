---
title: シングル サインオン トークンを使用してユーザーを認証する
description: サービスに SSO を実装するために Outlook アドインが提供するシングル サインオン トークンを使用することについて説明します。
ms.date: 04/28/2020
localization_priority: Normal
ms.openlocfilehash: 6d144e9ae4dcaf03705deb75f58c2f67a9c03106
ms.sourcegitcommit: 7d5407d3900d2ad1feae79a4bc038afe50568be0
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/30/2020
ms.locfileid: "46530465"
---
# <a name="authenticate-a-user-with-a-single-sign-on-token-in-an-outlook-add-in-preview"></a>Outlook アドインでシングル サインオン トークンを使用してユーザーを認証する (プレビュー)

シングル サインオン (SSO) は、アドインがユーザーを認証する (またオプションでアクセス トークンを認証および取得して [Microsoft Graph API](/graph/overview) を呼び出す) ための、シームレスな方法を提供します。

この方法を使用すると、アドインはサーバーのバックエンド API にスコープされたアクセス トークンを取得できます。 アドインはこれを `Authorization` ヘッダーのベアラー トークンとして使用して、API へのコールバックを認証します。 オプションとして、サーバー側のコードも持つことができます。

- On-Behalf-Of フローを完了して、Microsoft Graph API にスコープ設定されたアクセス トークンを取得する
- トークン内の ID 情報を使用して、独自のバックエンド サービスに対するユーザーの識別と認証を確立する

Office アドインの SSO の概要については、[「Office アドインのシングル サインオンを有効化する」](../develop/sso-in-office-add-ins.md) および[「Office アドインの Microsoft Graph への承認」](../develop/authorize-to-microsoft-graph.md)を参照してください。


## <a name="enable-modern-authentication-in-your-microsoft-365-tenancy"></a>Microsoft 365 テナントで先進認証を有効にする

Outlook アドインで SSO を使用するには、Microsoft 365 テナントの先進認証を有効にする必要があります。 この方法の詳細については、「[Exchange Online: テナントの先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)」を参照してください。

## <a name="register-your-add-in"></a>アドインを登録する

SSO を使用するには、Outlook アドインに Azure Active Directory (AAD) v2.0 を登録したサーバー側ウェブ API が必要です。 詳細については、「[Azure AD v2.0 のエンドポイントで SSO を使用する Office アドインを登録する](../develop/register-sso-add-in-aad-v2.md)」をご覧ください。

### <a name="provide-consent-when-sideloading-an-add-in"></a>アドインのサイドロード時に同意する

SSO を使用するアドインが AppSource から取得される場合、Microsoft Graph のスコープが含まれている場合は、同意を得るためのバックアップの認証方法が必要です。 アドインを開発している場合は、事前に同意を得る必要があります。 詳細については、「[アドインに管理者の同意を付与する](../develop/grant-admin-consent-to-an-add-in.md)」を参照してください。

## <a name="update-the-add-in-manifest"></a>アドイン マニフェストを更新する

アドインで SSO を使用できるようにするための次の手順では、`VersionOverridesV1_1` の [VersionOverrides](../reference/manifest/versionoverrides.md) 要素の最後に `WebApplicationInfo` 要素を追加します。 詳細については、「[アドインを構成する](../develop/sso-in-office-add-ins.md#configure-the-add-in)」を参照してください。

## <a name="get-the-sso-token"></a>SSO トークンを取得する

アドインがクライアント側スクリプトを含む SSO トークンを取得します。 詳細については、[「クライアント側のコードを追加する」](../develop/sso-in-office-add-ins.md#add-client-side-code)を参照してください。

## <a name="use-the-sso-token-at-the-back-end"></a>バックエンドで SSO トークンを使用する

ほとんどの場合、アドインがサーバー側に渡してそこで使用しない場合は、アクセス トークンを取得してもあまり意味はありません。 サーバー側で可能となること、また必要な対応に関する詳細は、[「サーバー側のコードを追加する」](../develop/sso-in-office-add-ins.md#add-server-side-code)を参照してください。

> [!IMPORTANT]
> ID として SSO トークンを *Outlook* アドインで使用するときには、代替の ID として [Exchange の ID トークンも使用](authenticate-a-user-with-an-identity-token.md)することをお勧めします。 アドインのユーザーは、複数のクライアントを使用することがあり、一部のクライアントは SSO トークンの提示をサポートしていないことがあります。 代わりに Exchange の ID トークン使用すると、そうしたユーザーに資格情報の入力を求めるダイアログを複数回表示しないようにできます。 詳細については、「[シナリオ: Outlook アドインでサービスにシングル サインオンを実装する](implement-sso-in-outlook-add-in.md)」を参照してください。

## <a name="see-also"></a>関連項目

- Microsoft Graph API へのアクセスに SSO トークンを使用するサンプル Outlook アドインについては、[「AttachmentsDemo サンプル アドイン」](https://github.com/OfficeDev/outlook-add-in-attachments-demo)を参照してください。
- [SSO API リファレンス](../develop/sso-in-office-add-ins.md#sso-api-reference)
- [IdentityAPI 要件セット](../reference/requirement-sets/identity-api-requirement-sets.md)
