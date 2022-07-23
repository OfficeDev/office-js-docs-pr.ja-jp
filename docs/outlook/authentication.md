---
title: Outlook アドインの認証オプション
description: Outlook アドインは、特定のシナリオに応じて、さまざまな認証メソッドを提供します。
ms.date: 09/03/2021
ms.localizationpriority: high
ms.openlocfilehash: 17ab09a1f0cdbf7668fa80080e587dd3d800f2c6
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958363"
---
# <a name="authentication-options-in-outlook-add-ins"></a>Outlook アドインの認証オプション

Outlook アドインは、アドインをホストするサーバー、内部ネットワーク、クラウド内の別の場所などに関わらず、インターネット上のあらゆる場所から情報にアクセスできます。 その情報が保護されている場合、アドインにはユーザーを認証する方法が必要になります。 Outlook アドインは、特定のシナリオに応じて、さまざまな認証メソッドを提供します。

## <a name="single-sign-on-access-token"></a>シングル サインオン アクセス トークン

シングル サインオン アクセス トークンは、アドインがアクセス トークンを認証および取得して [Microsoft Graph API](/graph/overview) を呼び出すための、シームレスな方法を提供します。 ユーザーが資格情報を入力する必要がないため、この機能は摩擦を低減します。

> [!NOTE]
> 現在、シングル サインオン API は Word、Excel、Outlook、PowerPoint でサポートされています。シングル サインオン API の現在のサポート状態に関する詳細は、「[IdentityAPI の要件セット](/javascript/api/requirement-sets/common/identity-api-requirement-sets)」を参照してください。Outlook アドインで作業している場合は、Microsoft 365 テナントで先進認証を有効にしてください。この方法の詳細については、「[Exchange Online: テナントで先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)」を参照してください。

アドインが次の場合は、SSO アクセス トークンの使用を検討してください。

- 主に Microsoft 365 ユーザーが使用する
- 次のものにアクセスする必要がある。
  - Microsoft Graph の一部として公開されている Microsoft サービス
  - ユーザーが制御する Microsoft 以外のサービス

SSO 認証方法は、Azure Active Directory が提供する [OAuth2 On-Behalf-Of フロー](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of)を使用します。 それには、アドインを[アプリケーション登録ポータル](https://apps.dev.microsoft.com/)に登録し、必要な Microsoft Graph スコープをマニフェストで指定する必要があります。

この方法を使用すると、アドインはサーバーのバックエンド API にスコープされたアクセス トークンを取得できます。 アドインはこれを `Authorization` ヘッダーのベアラー トークンとして使用して、API へのコールバックを認証します。 その時点で、サーバーでは次の操作を実行できます。

- On-Behalf-Of フローを完了して、Microsoft Graph API にスコープ設定されたアクセス トークンを取得する
- トークン内の ID 情報を使用して、独自のバックエンド サービスに対するユーザーの識別と認証を確立する

より詳しい概要については、[SSO 認証方式の概要全文](../develop/sso-in-office-add-ins.md)を参照してください。

Outlook アドインでの SSO トークンの使用の詳細については、「[Outlook アドインでシングルサインオン トークンを使用してユーザーを認証する](authenticate-a-user-with-an-sso-token.md)」を参照してください。

SSO トークンを使用するアドインのサンプルについては、「[Outlook アドイン SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO)」を参照してください。

## <a name="exchange-user-identity-token"></a>Exchange のユーザー ID トークン

Exchange のユーザー ID トークンは、アドインがユーザーの ID を確立する方法を提供します。 ユーザーの ID を確認することで、バックエンド システムにワンタイム認証を実行し、将来の要求に対する認証としてユーザー ID トークンを使用することができます。 次の場合、Exchange のユーザー ID トークンを使用します。

- アドインが主に Exchange のオンプレミス ユーザーによって使用される場合。
- アドインが、ユーザーが制御する Microsoft 以外のサービスにアクセスする必要がある場合。
- SSO をサポートしていないバージョンの Office でアドインが実行されている場合の代替認証として。

アドインでは、[getUserIdentityTokenAsync](/javascript/api/outlook/office.mailbox#outlook-office-mailbox-getuseridentitytokenasync-member(1)) を呼び出して Exchange のユーザー ID トークンを取得できます。 それらのトークンの使用の詳細については、「[Exchange の ID トークンを使用してユーザーを認証する](authenticate-a-user-with-an-identity-token.md)」を参照してください。

## <a name="access-tokens-obtained-via-oauth2-flows"></a>OAuth2 フローで取得されたアクセス トークン

アドインは、認証のために OAuth2 をサポートする Microsoft やその他のサービスにアクセスすることもできます。 アドインが次の場合は、OAuth2 トークンの使用を検討してください。

- 制御外のサービスにアクセスする必要があります。

このメソッドを使用すると、アドインは [displayDialogAsync](/javascript/api/office/office.ui#office-office-ui-displaydialogasync-member(1)) メソッドを使用して OAuth2 フローを初期化することで、ユーザーにサービスへのサインインを求めます。

## <a name="callback-tokens"></a>コールバック トークン

コールバック トークンは、[Exchange Web サービス (EWS)](/exchange/client-developer/exchange-web-services/explore-the-ews-managed-api-ews-and-web-services-in-exchange) または [Outlook REST API](/previous-versions/office/office-365-api/api/version-2.0/use-outlook-rest-api) を使用して、サーバーのバックエンドからユーザーのメールボックスへのアクセスを提供します。アドインが次に当てはまる場合は、コールバック トークンの使用を検討してください。

- サーバーのバックエンドからユーザーのメールボックスにアクセスする必要がある。

アドインは、[getCallbackTokenAsync メソッド](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)の 1 つを使用して、コールバック トークンを取得します。 アクセスのレベルは、アドイン マニフェストで指定されたアクセス許可によって制御されます。
