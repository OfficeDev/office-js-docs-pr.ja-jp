---
title: シングル サインオン トークンを使用してユーザーを認証する
description: サービスに SSO を実装するために Outlook アドインが提供するシングル サインオン トークンを使用することについて説明します。
ms.date: 09/03/2021
ms.localizationpriority: medium
ms.openlocfilehash: 41eddbcc1db05ca618506ce4810bf2bb795e59f7
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154225"
---
# <a name="authenticate-a-user-with-a-single-sign-on-token-in-an-outlook-add-in"></a>アドイン内のシングル サインオン トークンを使用してユーザー Outlook認証する

シングル サインオン (SSO) は、アドインがユーザーを認証する (またオプションでアクセス トークンを認証および取得して [Microsoft Graph API](/graph/overview) を呼び出す) ための、シームレスな方法を提供します。

この方法を使用すると、アドインはサーバーのバックエンド API にスコープされたアクセス トークンを取得できます。 アドインはこれを `Authorization` ヘッダーのベアラー トークンとして使用して、API へのコールバックを認証します。 必要に応じて、サーバー側のコードを使用することもできます。

- On-Behalf-Of フローを完了して、Microsoft Graph API にスコープ設定されたアクセス トークンを取得する
- トークン内の ID 情報を使用して、独自のバックエンド サービスに対するユーザーの識別と認証を確立する

Office アドインの SSO の概要については、[「Office アドインのシングル サインオンを有効化する」](../develop/sso-in-office-add-ins.md) および[「Office アドインの Microsoft Graph への承認」](../develop/authorize-to-microsoft-graph.md)を参照してください。

## <a name="enable-modern-authentication-in-your-microsoft-365-tenancy"></a>テナントでモダン認証を有効Microsoft 365する

SSO をアドインと一緒Outlookするには、テナントのモダン認証を有効Microsoft 365があります。 この方法の詳細については、「[Exchange Online: テナントの先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)」を参照してください。

## <a name="register-your-add-in"></a>アドインを登録する

SSO を使用するには、Outlook アドインに Azure Active Directory (AAD) v2.0 を登録したサーバー側ウェブ API が必要です。 詳細については、「[Azure AD v2.0 のエンドポイントで SSO を使用する Office アドインを登録する](../develop/register-sso-add-in-aad-v2.md)」をご覧ください。

### <a name="provide-consent-when-sideloading-an-add-in"></a>アドインのサイドロード時に同意する

アドインを開発する場合は、事前に同意する必要があります。 詳細については、「管理者の [同意をアドインに付与する」を参照してください](../develop/grant-admin-consent-to-an-add-in.md)。

## <a name="update-the-add-in-manifest"></a>アドイン マニフェストを更新する

アドインで SSO を使用できるようにするための次の手順では、`VersionOverridesV1_1` の [VersionOverrides](../reference/manifest/versionoverrides.md) 要素の最後に `WebApplicationInfo` 要素を追加します。 詳細については、「[アドインを構成する](../develop/sso-in-office-add-ins.md#configure-the-add-in)」を参照してください。

## <a name="get-the-sso-token"></a>SSO トークンを取得する

アドインがクライアント側スクリプトを含む SSO トークンを取得します。 詳細については、[「クライアント側のコードを追加する」](../develop/sso-in-office-add-ins.md#add-client-side-code)を参照してください。

## <a name="use-the-sso-token-at-the-back-end"></a>バックエンドで SSO トークンを使用する

ほとんどの場合、アドインがサーバー側に渡してそこで使用しない場合は、アクセス トークンを取得してもあまり意味はありません。 サーバー側で可能となること、また必要な対応に関する詳細は、[「サーバー側のコードを追加する」](../develop/sso-in-office-add-ins.md#add-server-side-code)を参照してください。

> [!IMPORTANT]
> ID として SSO トークンを *Outlook* アドインで使用するときには、代替の ID として [Exchange の ID トークンも使用](authenticate-a-user-with-an-identity-token.md)することをお勧めします。 アドインのユーザーは、複数のクライアントを使用することがあり、一部のクライアントは SSO トークンの提示をサポートしていないことがあります。 代わりに Exchange の ID トークン使用すると、そうしたユーザーに資格情報の入力を求めるダイアログを複数回表示しないようにできます。 詳細については、「[シナリオ: Outlook アドインでサービスにシングル サインオンを実装する](implement-sso-in-outlook-add-in.md)」を参照してください。

## <a name="see-also"></a>関連項目

- SSO トークンをOutlookして Microsoft Graph API にアクセスするアドインのサンプルについては、「Outlook SSO」を[参照してください](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/auth/Outlook-Add-in-SSO)。
- [SSO API リファレンス](../develop/sso-in-office-add-ins.md#sso-api-reference)
- [IdentityAPI 要件セット](../reference/requirement-sets/identity-api-requirement-sets.md)
