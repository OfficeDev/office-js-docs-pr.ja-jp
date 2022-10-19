---
title: シングル サインオン トークンを使用してユーザーを認証する
description: サービスに SSO を実装するために Outlook アドインが提供するシングル サインオン トークンを使用することについて説明します。
ms.date: 10/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: 23b7936cc0ba4453a2a10cbfe0731941a913c118
ms.sourcegitcommit: eca6c16d0bb74bed2d35a21723dd98c6b41ef507
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/18/2022
ms.locfileid: "68607444"
---
# <a name="authenticate-a-user-with-a-single-sign-on-token-in-an-outlook-add-in"></a>Outlook アドインでシングル サインオン トークンを使用してユーザーを認証する

シングル サインオン (SSO) は、アドインがユーザーを認証する (またオプションでアクセス トークンを認証および取得して [Microsoft Graph API](/graph/overview) を呼び出す) ための、シームレスな方法を提供します。

この方法を使用すると、アドインはサーバーのバックエンド API にスコープされたアクセス トークンを取得できます。 アドインはこれを `Authorization` ヘッダーのベアラー トークンとして使用して、API へのコールバックを認証します。 必要に応じて、サーバー側のコードを作成することもできます。

- On-Behalf-Of フローを完了して、Microsoft Graph API にスコープ設定されたアクセス トークンを取得する
- トークン内の ID 情報を使用して、独自のバックエンド サービスに対するユーザーの識別と認証を確立する

Office アドインの SSO の概要については、[「Office アドインのシングル サインオンを有効化する」](../develop/sso-in-office-add-ins.md) および[「Office アドインの Microsoft Graph への承認」](../develop/authorize-to-microsoft-graph.md)を参照してください。

## <a name="enable-modern-authentication-in-your-microsoft-365-tenancy"></a>Microsoft 365 テナントで先進認証を有効にする

Outlook アドインで SSO を使用するには、Microsoft 365 テナントの先進認証を有効にする必要があります。 この方法の詳細については、「[Exchange Online: テナントの先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)」を参照してください。

## <a name="register-your-add-in"></a>アドインを登録する

SSO を使用するには、Outlook アドインに Azure Active Directory (AAD) v2.0 を登録したサーバー側ウェブ API が必要です。 詳細については、「[Azure AD v2.0 のエンドポイントで SSO を使用する Office アドインを登録する](../develop/register-sso-add-in-aad-v2.md)」をご覧ください。

### <a name="provide-consent-when-sideloading-an-add-in"></a>アドインのサイドロード時に同意する

アドインを開発する場合は、事前に同意する必要があります。 詳細については、「 [管理者にアドインへの同意を付与する](../develop/grant-admin-consent-to-an-add-in.md)」を参照してください。

## <a name="update-the-add-in-manifest"></a>アドイン マニフェストを更新する

アドインで SSO を有効にする次の手順は、アドインのMicrosoft ID プラットフォーム登録からマニフェストにいくつかの情報を追加することです。 マークアップは、マニフェストの種類によって異なります。

- **XML マニフェスト**: [VersionOverrides](/javascript/api/manifest/versionoverrides) 要素の末尾に`VersionOverridesV1_1`要素を追加`WebApplicationInfo`します。 次に、必要な子要素を追加します。 マークアップの詳細については、「 [アドインの構成](../develop/sso-in-office-add-ins.md#configure-the-add-in)」を参照してください。
- **Teams マニフェスト (プレビュー)**: マニフェストのルート `{ ... }` オブジェクトに "webApplicationInfo" プロパティを追加します。 このオブジェクトに、アドインを登録したときにAzure portalで生成されたアドインの Web アプリのアプリケーション ID に設定された子 "id" プロパティを指定します。 (この記事の「アドイン [を登録する](#register-your-add-in) 」セクションを参照してください)。また、アドインを登録したときに設定したのと同じ **アプリケーション ID URI** に設定された子 "resource" プロパティも指定します。 この URI にはフォームが必要です `api://<fully-qualified-domain-name>/<application-id>`。 次に例を示します。

   ```json
   "webApplicationInfo": {
        "id": "a661fed9-f33d-4e95-b6cf-624a34a2f51d",
        "resource": "api://addin.contoso.com/a661fed9-f33d-4e95-b6cf-624a34a2f51d"
    },
   ```

  > [!NOTE]
  > Teams マニフェストを使用する SSO 対応アドインはサイドロードできますが、現時点では他の方法では展開できません。

## <a name="get-the-sso-token"></a>SSO トークンを取得する

アドインがクライアント側スクリプトを含む SSO トークンを取得します。 詳細については、[「クライアント側のコードを追加する」](../develop/sso-in-office-add-ins.md#add-client-side-code)を参照してください。

## <a name="use-the-sso-token-at-the-back-end"></a>バックエンドで SSO トークンを使用する

ほとんどの場合、アドインがサーバー側に渡してそこで使用しない場合は、アクセス トークンを取得してもあまり意味はありません。 サーバー側で可能となること、また必要な対応に関する詳細は、[「サーバー側のコードを追加する」](../develop/sso-in-office-add-ins.md#pass-the-access-token-to-server-side-code)を参照してください。

> [!IMPORTANT]
> ID として SSO トークンを *Outlook* アドインで使用するときには、代替の ID として [Exchange の ID トークンも使用](authenticate-a-user-with-an-identity-token.md)することをお勧めします。 アドインのユーザーは、複数のクライアントを使用することがあり、一部のクライアントは SSO トークンの提示をサポートしていないことがあります。 代わりに Exchange の ID トークン使用すると、そうしたユーザーに資格情報の入力を求めるダイアログを複数回表示しないようにできます。 詳細については、「[シナリオ: Outlook アドインでサービスにシングル サインオンを実装する](implement-sso-in-outlook-add-in.md)」を参照してください。

## <a name="sso-for-event-based-activation"></a>イベント ベースのアクティブ化の SSO

アドインでイベント ベースのアクティブ化が使用されている場合は、追加の手順を実行する必要があります。 詳細については、「 [イベント ベースのアクティブ化を使用する Outlook アドインでシングル サインオン (SSO) を有効にする](use-sso-in-event-based-activation.md)」を参照してください。

## <a name="see-also"></a>関連項目

- [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#office-runtime-officeruntime-auth-getaccesstoken-member(1))
- SSO トークンを使用して Microsoft Graph APIにアクセスするサンプル Outlook アドインについては、「[Outlook アドインの SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO)」を参照してください。
- [SSO API リファレンス](/javascript/api/office/office.auth#office-office-auth-getaccesstoken-member(1))
- [IdentityAPI 要件セット](/javascript/api/requirement-sets/common/identity-api-requirement-sets)
- [イベント ベースのアクティブ化を使用する Outlook アドインでシングル サインオン (SSO) を有効にする](use-sso-in-event-based-activation.md)
