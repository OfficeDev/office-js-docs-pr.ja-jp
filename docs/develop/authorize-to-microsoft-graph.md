---
title: SSO を使用した Microsoft Graph への承認
description: Microsoft アドインからデータをOfficeシングル サインオン (SSO) を使用する方法について説明Graph。
ms.date: 01/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 538648e96233bd0c2b497ef588d10c4f708e8522
ms.sourcegitcommit: 57e15f0787c0460482e671d5e9407a801c17a215
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/02/2022
ms.locfileid: "62320268"
---
# <a name="authorize-to-microsoft-graph-with-sso"></a>SSO を使用した Microsoft Graph への承認

ユーザーは個人用の Microsoft アカウントまたは Microsoft 365 Education または職場アカウントのいずれかを使用して、Office (オンライン、モバイル、およびデスクトップ プラットフォーム) にサインインします。 Office アドインの [Microsoft Graph](https://developer.microsoft.com/graph/docs) へのアクセスの承認には、ユーザーの Office サインオン資格証明を使用するのが最良の方法です。 これにより、2 回目はサインインする必要なく Microsoft Graph データにアクセスできます。

## <a name="add-in-architecture-for-sso-and-microsoft-graph"></a>SSO と Microsoft Graph のアドイン アーキテクチャ

Web アプリケーションのページと JavaScript をホスティングすることに加え、アドインでは、同一の[完全修飾ドメイン名](/windows/desktop/DNS/f-gly#_dns_fully_qualified_domain_name_fqdn__gly)で、Microsoft Graph へのアクセス トークンを取得して、要求を送信するための 1 つ以上の Web API をホストする必要もあります。

アドイン マニフェストには、アドインに必要な Microsoft Graph へのアクセス許可など、重要な Azure アプリ登録情報を Office に提供する **WebApplicationInfo** 要素が含まれています。

### <a name="how-it-works-at-runtime"></a>実行時の動作のしくみ

次の図は、サインインして Microsoft アカウントにアクセスするために必要な手順をGraph。 プロセス全体で OAuth 2.0 および JWT アクセス トークンが使用されます。

:::image type="content" source="../images/sso-access-to-microsoft-graph.svg" alt-text="SSO プロセスを示す図。" border="false":::

1. アドインのクライアント側コードは、API [getAccessToken Office.js呼び出します](/javascript/api/office-runtime/officeruntime.auth#getAccessToken_options_)。 これにより、アドインOfficeアクセス トークンを取得する必要があります。

    ユーザーがサインインしていない場合、Officeホストは、Microsoft ID プラットフォームにサインインして同意するための UI を提供します。

2. このOfficeホストは、アクセス トークンを要求します。Microsoft ID プラットフォーム。
3. このMicrosoft ID プラットフォームは、アクセス トークン *A を* ホストにOfficeします。 アクセス トークン *アドイン* の独自のサーバー側 API へのアクセスのみを提供します。 Microsoft サーバーへのアクセスは提供Graph。
4. このOfficeは、アドインのクライアント側コードにアクセス トークン *A* を返します。 これで、クライアント側コードはサーバー側 API に対して認証された呼び出しを行えます。
5. クライアント側のコードは、認証を必要とするサーバー側の Web API に対して HTTP 要求を行います。 これには、アクセス トークン *A が承認* 証明として含まれます。 サーバー側のコードは、アクセス トークン *A を検証します*。
6. サーバー側のコードでは、OAuth 2.0 On-Behalf-Of フロー (OBO) を使用して、Microsoft Graph へのアクセス許可を持つ新しいアクセス トークンを要求します。
7. このMicrosoft ID プラットフォームは、Microsoft Graph へのアクセス許可を持つ新しいアクセス トークン *B* を返します (アドインがアクセス許可を要求する場合は更新 *トークンoffline_access* します。 サーバーはオプションでアクセス トークン B をキャッシュ *できます*。
8. サーバー側のコードは、Microsoft Graph API に要求を行い、Microsoft Graph へのアクセス許可を持つアクセス トークン *B* を含Graph。
9. Microsoft Graphサーバー側のコードにデータを返します。
10. サーバー側のコードは、データをクライアント側のコードに戻します。

以降の要求では、サーバー側コードに対して認証された呼び出しを行う場合、クライアント コードは常にアクセス トークン *A* を渡します。 サーバー側のコードは、トークン B をキャッシュして、将来の API 呼び出しで再び要求する必要が生じないので、トークン *B* をキャッシュできます。

## <a name="develop-an-sso-add-in-that-accesses-microsoft-graph"></a>Microsoft Graph にアクセスする SSO を開発する

SSO を使用する他のアプリケーションと同様に、Microsoft Graphにアクセスするアドインを開発します。 詳しい説明については、「[シングル サインオンを有効にする」を参照Office参照してください](../develop/sso-in-office-add-ins.md)。違いは、アドインがサーバー側 Web API を持っている必要がある点です。

使用する言語とフレームワークによっては、記述する必要のあるサーバー側コードが簡単になるライブラリが使用できることがあります。 コードでは、次に示す操作を実行する必要があります。

* アクセス トークン A が *クライアント側* コードから渡される度に検証します。 詳細については、[アクセス トークンの検証](sso-in-office-add-ins.md#pass-the-access-token-to-server-side-code)に関するページを参照してください。
* アクセス トークン、ユーザーに関するメタデータ、アドインの資格情報 (ID とシークレット) を含む Microsoft ID プラットフォーム への呼び出しで OAuth 2.0 On-Behalf-Of フロー (OBO) を開始します。 OBO フローの詳細については、「Microsoft ID プラットフォーム [OAuth 2.0 On-Behalf-Of フロー」を参照してください](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow)。
* 必要に応じて、フローが完了したら、返されたアクセス トークン *B* を Microsoft サーバーに対するアクセス許可でキャッシュGraph。 アドインで Microsoft Graph を複数回呼び出す場合に、これが行われます。 詳細については、「Microsoft 認証ライブラリ (MSAL) を使用したトークンの取得と [キャッシュ」を参照してください。](/azure/active-directory/develop/msal-acquire-cache-tokens)
* (キャッシュされている可能性がある) アクセス トークン B を Microsoft Graph に渡して、Microsoft データを取得する 1 つ以上の *Web* API メソッドをGraph。

詳細なチュートリアルとシナリオについては、次を参照してください。

* [シングル サインオンを使用する Node.js Office アドインを作成する](create-sso-office-add-ins-nodejs.md)
* [シングル サインオンを使用する ASP.NET Office アドインを作成する](create-sso-office-add-ins-aspnet.md)
* [シナリオ: Outlook アドインでサービスにシングル サインオンを実装する](../outlook/implement-sso-in-outlook-add-in.md)

## <a name="distributing-sso-enabled-add-ins-in-microsoft-appsource"></a>Microsoft AppSource での SSO 対応アドインの配布

管理者Microsoft 365 [AppSource](https://appsource.microsoft.com) からアドインを取得すると、管理者は統合アプリを通じてアドインを再配布し、[](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps)Microsoft Graph スコープにアクセスするためのアドインに対する管理者の同意を付与できます。 ただし、エンド ユーザーが AppSource から直接アドインを取得することもできます。その場合、ユーザーはアドインに同意する必要があります。 これにより、ソリューションを提供した潜在的なパフォーマンスの問題が発生する可能性があります。

`allowConsentPrompt` `getAccessToken``OfficeRuntime.auth.getAccessToken( { allowConsentPrompt: true } );`Office の呼び出しでオプションを渡した場合、Microsoft ID プラットフォーム が Office に同意がまだアドインに付与されていないと報告した場合は、ユーザーに同意を求めるメッセージを表示できます。 ただし、セキュリティ上の理由からOfficeは、Microsoft `profile` のスコープに同意するようユーザーにGraphできます。 *Office Microsoft の他のスコープへの同意* を求Graphを求め、それもできません`User.Read`。 つまり、ユーザーがプロンプトに同意を与える場合は、Officeトークンを返します。 ただし、追加の Microsoft Graph スコープを持つ新しいアクセス トークンに対してアクセス トークンを交換しようとすると、エラー AADSTS65001 が失敗します。つまり、同意 (Microsoft Graph スコープへの) が付与されていないという意味です。

> [!NOTE]
> 管理者がエンドユーザーの同意 `{ allowConsentPrompt: true }` を `profile` 無効にしている場合でも、同意の要求はスコープに対しても失敗する可能性があります。 詳細については、「Configure [how end-users consent to applications](/azure/active-directory/manage-apps/configure-user-consent) using Azure Active Directory。

コードでは、別の認証システムに戻ってこのエラーを処理できます。このエラーは、ユーザーに Microsoft のスコープに対する同意を求Graphがあります。 コード例については、「シングル [サインオンを](create-sso-office-add-ins-nodejs.md)使用する Node.js Office アドインを作成する」および「シングル サインオンを使用する [ASP.NET Office](create-sso-office-add-ins-aspnet.md) アドインを作成する」とリンク先のサンプルを参照してください。 プロセス全体では、複数のラウンド トリップが必要Microsoft ID プラットフォーム。 このパフォーマンスの低下を回避するには、`forMSGraphAccess`オプションを呼び出し`getAccessToken`に含める 。 `OfficeRuntime.auth.getAccessToken( { forMSGraphAccess: true } )` これは、アドインOfficeスコープに Microsoft が必要Graphします。 Officeは、microsoft Microsoft ID プラットフォームスコープへの同意が既にアドインに付与済みGraph確認を求めるメッセージが表示されます。 アクセス トークンがある場合は、アクセス トークンが返されます。 呼び出しが実行できない場合は `getAccessToken` 、エラー 13012 が返されます。 コードは、トークンとトークンの交換を行わずに、すぐに別の認証システムに戻ってこのエラーを処理Microsoft ID プラットフォーム。

ベスト プラクティスとして、`forMSGraphAccess``getAccessToken`アドインが AppSource で配布され、Microsoft のスコープが必要になるGraphします。

## <a name="details-on-sso-with-an-outlook-add-in"></a>カスタム アドインを使用Outlook SSO の詳細

SSO を使用する Outlook アドインを開発し、テスト用にサイドロードした場合、管理者の同意が与えらた場合でも、Office は常にエラー 13012 `forMSGraphAccess` `getAccessToken` を返します。 このため、アドインを開発するときに`forMSGraphAccess`オプションをコメントアウトOutlook必要があります。 実稼働環境に展開する場合は、必ずオプションのアンコメントを解除してください。 偽の 13012 は、このページでサイドローディングを行Outlook。

アドインOutlook、テナントのモダン認証を必ず有効Microsoft 365してください。 この方法の詳細については、「[Exchange Online: テナントの先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)」を参照してください。

## <a name="see-also"></a>関連項目

* [OAuth2 トークン Exchange](https://tools.ietf.org/html/draft-ietf-oauth-token-exchange-02)
* [Microsoft ID プラットフォーム OAuth 2.0 On-Behalf-Of フロー](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow)
* [IdentityAPI 要件セット](../reference/requirement-sets/identity-api-requirement-sets.md)
