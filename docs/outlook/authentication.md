---
title: Outlook アドインの認証オプション
description: Outlook アドインは、特定のシナリオに応じて、さまざまな認証メソッドを提供します。
ms.date: 10/17/2022
ms.localizationpriority: high
ms.openlocfilehash: d8ae8971c4095e5314885514226cd8f52728fb07
ms.sourcegitcommit: eca6c16d0bb74bed2d35a21723dd98c6b41ef507
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/18/2022
ms.locfileid: "68607528"
---
# <a name="authentication-options-in-outlook-add-ins"></a>Outlook アドインの認証オプション

Outlook アドインは、アドインをホストするサーバー、内部ネットワーク、クラウド内の別の場所などに関わらず、インターネット上のあらゆる場所から情報にアクセスできます。 その情報が保護されている場合、アドインにはユーザーを認証する方法が必要になります。 Outlook アドインは、特定のシナリオに応じて、さまざまな認証メソッドを提供します。

## <a name="single-sign-on-access-token"></a>シングル サインオン アクセス トークン

シングル サインオン アクセス トークンは、アドインがアクセス トークンを認証および取得して [Microsoft Graph API](/graph/overview) を呼び出すための、シームレスな方法を提供します。 ユーザーが資格情報を入力する必要がないため、この機能は摩擦を低減します。

> [!NOTE]
> The Single Sign-on API is currently supported for Word, Excel, Outlook, and PowerPoint. For more information about where the Single Sign-on API is currently supported, see [IdentityAPI requirement sets](/javascript/api/requirement-sets/common/identity-api-requirement-sets).
> If you are working with an Outlook add-in, be sure to enable Modern Authentication for the Microsoft 365 tenancy. For information about how to do this, see [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

アドインが次の場合は、SSO アクセス トークンの使用を検討してください。

- 主に Microsoft 365 ユーザーが使用する
- 次のものにアクセスする必要がある。
  - Microsoft Graph の一部として公開されている Microsoft サービス
  - ユーザーが制御する Microsoft 以外のサービス

SSO 認証方法は、Azure Active Directory が提供する [OAuth2 On-Behalf-Of フロー](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of)を使用します。 それには、アドインを[アプリケーション登録ポータル](https://apps.dev.microsoft.com/)に登録し、必要な Microsoft Graph スコープをマニフェストで指定する必要があります。

> [!NOTE]
> アドインが Office アドインの [Teams マニフェスト (プレビュー)](../develop/json-manifest-overview.md) を使用している場合、マニフェストの構成がいくつかありますが、Microsoft Graph スコープは指定されていません。 Teams マニフェストを使用する SSO 対応アドインはサイドロードできますが、現時点では他の方法では展開できません。

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

Callback tokens provide access to the user's mailbox from your server back-end, either using [Exchange Web Services (EWS)](/exchange/client-developer/exchange-web-services/explore-the-ews-managed-api-ews-and-web-services-in-exchange), or the [Outlook REST API](/previous-versions/office/office-365-api/api/version-2.0/use-outlook-rest-api). Consider using callback tokens if your add-in:

- サーバーのバックエンドからユーザーのメールボックスにアクセスする必要がある。

アドインは、[getCallbackTokenAsync メソッド](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)の 1 つを使用して、コールバック トークンを取得します。 アクセスのレベルは、アドイン マニフェストで指定されたアクセス許可によって制御されます。
