---
title: SSO を使用した Microsoft Graph への承認
description: Office アドインのユーザーがシングル サインオン (SSO) を使用して Microsoft Graph からデータをフェッチする方法について説明します。
ms.date: 06/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 4ecb945dcd99400065fde3e80e8b60d67266e0b1
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/06/2022
ms.locfileid: "66659648"
---
# <a name="authorize-to-microsoft-graph-with-sso"></a>SSO を使用した Microsoft Graph への承認

ユーザーは、個人用の Microsoft アカウントまたはMicrosoft 365 Educationまたは職場のアカウントを使用して Office にサインインします。 Office アドインの [Microsoft Graph](https://developer.microsoft.com/graph/docs) へのアクセスの承認には、ユーザーの Office サインオン資格証明を使用するのが最良の方法です。 これにより、2 回目はサインインする必要なく Microsoft Graph データにアクセスできます。

## <a name="add-in-architecture-for-sso-and-microsoft-graph"></a>SSO と Microsoft Graph のアドイン アーキテクチャ

Web アプリケーションのページと JavaScript をホスティングすることに加え、アドインでは、同一の[完全修飾ドメイン名](/windows/desktop/DNS/f-gly#_dns_fully_qualified_domain_name_fqdn__gly)で、Microsoft Graph へのアクセス トークンを取得して、要求を送信するための 1 つ以上の Web API をホストする必要もあります。

アドイン マニフェストには、アドインに必要な **\<WebApplicationInfo\>** Microsoft Graph へのアクセス許可など、重要な Azure アプリ登録情報を Office に提供する要素が含まれています。

### <a name="how-it-works-at-runtime"></a>実行時の動作のしくみ

次の図は、Microsoft Graph にサインインしてアクセスするために必要な手順を示しています。 プロセス全体で OAuth 2.0 と JWT アクセス トークンが使用されます。

:::image type="content" source="../images/sso-access-to-microsoft-graph.svg" alt-text="SSO プロセスを示す図。" border="false":::

1. アドインのクライアント側コードは、Office.js API [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#office-runtime-officeruntime-auth-getaccesstoken-member(1)) を呼び出します。 これにより、アドインのアクセス トークンを取得するように Office ホストに指示されます。

    ユーザーがサインインしていない場合、Office ホストとMicrosoft ID プラットフォームを組み合わせて、ユーザーがサインインして同意するための UI が提供されます。

2. Office ホストは、Microsoft ID プラットフォームからアクセス トークンを要求します。
3. Microsoft ID プラットフォームは、アクセス トークン *A* を Office ホストに返します。 アクセス トークン *A* は、アドイン独自のサーバー側 API へのアクセスのみを提供します。 Microsoft Graph へのアクセスは提供されません。
4. Office ホストは、アドインのクライアント側コードにアクセス トークン *A* を返します。 これで、クライアント側コードでサーバー側 API に対して認証された呼び出しを行うことができます。
5. クライアント側のコードは、認証を必要とするサーバー側の Web API に HTTP 要求を行います。 承認証明としてアクセス トークン *A* が含まれます。 サーバー側コードは、アクセス トークン *A* を検証します。
6. サーバー側のコードでは、OAuth 2.0 On-Behalf-Of フロー (OBO) を使用して、Microsoft Graph へのアクセス許可を持つ新しいアクセス トークンを要求します。
7. Microsoft ID プラットフォームは、Microsoft Graph へのアクセス許可を持つ新しいアクセス トークン *B* を返します (アドインがアクセス許可を要求した場合は更新トークン *offline_access*)。 必要に応じて、サーバーはアクセス トークン *B* をキャッシュできます。
8. サーバー側のコードは、Microsoft Graph APIに要求を行い、Microsoft Graph へのアクセス許可を持つアクセス トークン *B* を含みます。
9. Microsoft Graph は、サーバー側のコードにデータを返します。
10. サーバー側コードは、データをクライアント側のコードに戻します。

後続の要求では、サーバー側コードに対して認証された呼び出しを行うとき、クライアント コードは常にアクセス トークン *A* を渡します。 サーバー側のコードはトークン *B* をキャッシュできるため、今後の API 呼び出しで再度要求する必要はありません。

## <a name="develop-an-sso-add-in-that-accesses-microsoft-graph"></a>Microsoft Graph にアクセスする SSO を開発する

SSO を使用する他のアプリケーションと同様に、Microsoft Graph にアクセスするアドインを開発します。 詳細な説明については、「 [Office アドインのシングル サインオンを有効にする](../develop/sso-in-office-add-ins.md)」を参照してください。違いは、アドインにサーバー側の Web API が必須である点です。

使用する言語とフレームワークによっては、記述する必要のあるサーバー側コードが簡単になるライブラリが使用できることがあります。 コードでは、次に示す操作を実行する必要があります。

* アクセス トークン *A* がクライアント側のコードから渡されるたびに検証します。 詳細については、[アクセス トークンの検証](sso-in-office-add-ins.md#pass-the-access-token-to-server-side-code)に関するページを参照してください。
* アクセス トークン、ユーザーに関するメタデータ、アドインの資格情報 (ID とシークレット) を含むMicrosoft ID プラットフォームを呼び出して、OAuth 2.0 On-Behalf-Of フロー (OBO) を開始します。 OBO フローの詳細については、「[Microsoft ID プラットフォームと OAuth 2.0 On-Behalf-Of フロー」を参照](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow)してください。
* 必要に応じて、フローが完了したら、返されたアクセス トークン *B* を Microsoft Graph へのアクセス許可でキャッシュします。 アドインで Microsoft Graph を複数回呼び出す場合に、これが行われます。 詳細については、「[Microsoft Authentication Library (MSAL) を使用したトークンの取得とキャッシュ」を](/azure/active-directory/develop/msal-acquire-cache-tokens)参照してください。
* (キャッシュされている可能性がある) アクセス トークン *B* を Microsoft Graph に渡すことによって、Microsoft Graph データを取得する 1 つ以上の Web API メソッドを作成します。

詳細なチュートリアルとシナリオについては、次を参照してください。

* [シングル サインオンを使用する Node.js Office アドインを作成する](create-sso-office-add-ins-nodejs.md)
* [シングル サインオンを使用する ASP.NET Office アドインを作成する](create-sso-office-add-ins-aspnet.md)
* [シナリオ: Outlook アドインでサービスにシングル サインオンを実装する](../outlook/implement-sso-in-outlook-add-in.md)

## <a name="distributing-sso-enabled-add-ins-in-microsoft-appsource"></a>Microsoft AppSource での SSO 対応アドインの配布

Microsoft 365 管理者が [AppSource](https://appsource.microsoft.com) からアドインを取得すると、管理者は [統合アプリ](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps) を通じてアドインを再配布し、Microsoft Graph スコープにアクセスするためのアドインへの管理者の同意を付与できます。 ただし、エンド ユーザーが AppSource から直接アドインを取得することも可能です。その場合、ユーザーはアドインに同意を与える必要があります。 これにより、ソリューションを提供した潜在的なパフォーマンスの問題が発生する可能性があります。

コードが [次のような] `allowConsentPrompt` の呼び出し`getAccessToken`でオプションを渡した場合、Microsoft ID プラットフォームがアドインに同意がまだ付与されていないと Office に報告した場合、Office はユーザーに同意を求`OfficeRuntime.auth.getAccessToken( { allowConsentPrompt: true } );`めることができます。 ただし、セキュリティ上の理由から、Office はユーザーに Microsoft Graph `profile` スコープへの同意のみを求めることができます。 *Office では、他の Microsoft Graph スコープへの同意を求めることができません*。. `User.Read`. つまり、ユーザーがプロンプトに同意を与えると、Office はアクセス トークンを返します。 ただし、追加の Microsoft Graph スコープを使用して新しいアクセス トークンにアクセス トークンを交換しようとすると、エラー AADSTS65001 が失敗します。つまり、同意が (Microsoft Graph スコープに対して) 付与されていないことを意味します。

> [!NOTE]
> 管理者がエンド ユーザーの同意を `{ allowConsentPrompt: true }` オフにした場合でも、 `profile` スコープの同意要求は失敗する可能性があります。 詳細については、「 [Azure Active Directory を使用してエンド ユーザーがアプリケーションに同意する方法を構成する」を](/azure/active-directory/manage-apps/configure-user-consent)参照してください。

コードでは、Microsoft Graph スコープへの同意をユーザーに求める別の認証システムにフォールバックすることで、このエラーを処理できます。また、処理する必要があります。 コード例については、「 [シングル サインオンを使用する Node.js Office アドインを作成する」と「シングル サインオン](create-sso-office-add-ins-nodejs.md) を [使用する ASP.NET Office アドインを作成する](create-sso-office-add-ins-aspnet.md) 」およびリンク先のサンプルを参照してください。 プロセス全体で、Microsoft ID プラットフォームへの複数のラウンド トリップが必要です。 このパフォーマンスの低下を回避するには、; の呼び出し`getAccessToken`にオプションを含めます`forMSGraphAccess`。たとえば、 `OfficeRuntime.auth.getAccessToken( { forMSGraphAccess: true } )`. これは、アドインに Microsoft Graph スコープが必要であることを Office に通知します。 Office はMicrosoft ID プラットフォームに対して、Microsoft Graph スコープへの同意がアドインに既に付与されていることを確認するよう求めます。 アクセス トークンがある場合は、アクセス トークンが返されます。 まだ呼び出 `getAccessToken` していない場合は、エラー 13012 が返されます。 コードでは、Microsoft ID プラットフォームとトークンを交換しようとしても、すぐに別の認証システムにフォールバックすることで、このエラーを処理できます。

ベスト プラクティスとして、アドインが AppSource で`getAccessToken`配布され、Microsoft Graph スコープが必要な場合は常に渡`forMSGraphAccess`してください。

## <a name="details-on-sso-with-an-outlook-add-in"></a>Outlook アドインを使用した SSO の詳細

SSO を使用する Outlook アドインを開発し、テスト用にサイドロードすると、管理者の同意が与えられた場合でも、Office は *常に* エラー 13012 `forMSGraphAccess` を `getAccessToken` 返します。 このため、Outlook アドインを **開発するときは**、この`forMSGraphAccess`オプションをコメントアウトする必要があります。 運用環境にデプロイするときは、必ずオプションのコメントを解除してください。 偽の 13012 は、Outlook でサイドローディングしている場合にのみ発生します。

Outlook アドインの場合は、Microsoft 365 テナントの先進認証を有効にしてください。 この方法の詳細については、「[Exchange Online: テナントの先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)」を参照してください。

## <a name="see-also"></a>関連項目

* [OAuth2 トークン交換](https://tools.ietf.org/html/draft-ietf-oauth-token-exchange-02)
* [Microsoft ID プラットフォームと OAuth 2.0 On-Behalf-Of フロー](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow)
* [IdentityAPI 要件セット](/javascript/api/requirement-sets/common/identity-api-requirement-sets)
