---
title: SSO を使用した Microsoft Graph への承認
description: Microsoft アドインからデータをOfficeシングル サインオン (SSO) を使用する方法について説明Graph。
ms.date: 07/27/2021
ms.localizationpriority: medium
ms.openlocfilehash: 1416a30088197b83e7b1095b325615c3ac6a5dc2
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59149716"
---
# <a name="authorize-to-microsoft-graph-with-sso"></a>SSO を使用した Microsoft Graph への承認

ユーザーは個人用の Microsoft アカウントまたは Microsoft 365 Education または職場アカウントのいずれかを使用して、Office (オンライン、モバイル、およびデスクトップ プラットフォーム) にサインインします。 Office アドインの [Microsoft Graph](https://developer.microsoft.com/graph/docs) へのアクセスの承認には、ユーザーの Office サインオン資格証明を使用するのが最良の方法です。 これにより、2 回目はサインインする必要なく Microsoft Graph データにアクセスできます。

> [!NOTE]
> 現在、シングル サインオン API は Word、Excel、Outlook, および PowerPoint でサポートされています。 シングル サインオン API の現在のサポート状態に関する詳細は、「[IdentityAPI の要件セット](../reference/requirement-sets/identity-api-requirement-sets.md)」を参照してください。
> Outlook アドインで作業している場合は、Microsoft 365 テナントの先進認証が有効になっていることを確認してください。 この方法の詳細については、「[Exchange Online: テナントの先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)」を参照してください。

## <a name="add-in-architecture-for-sso-and-microsoft-graph"></a>SSO と Microsoft Graph のアドイン アーキテクチャ

Web アプリケーションのページと JavaScript をホスティングすることに加え、アドインでは、同一の[完全修飾ドメイン名](/windows/desktop/DNS/f-gly#_dns_fully_qualified_domain_name_fqdn__gly)で、Microsoft Graph へのアクセス トークンを取得して、要求を送信するための 1 つ以上の Web API をホストする必要もあります。

アドイン マニフェストには、アドインを Azure Active Directory (Azure AD) v2.0 エンドポイントに登録する方法と、アドインが必要とする Microsoft Graph へのアクセス許可を指定するマークアップが含まれています。

### <a name="how-it-works-at-runtime"></a>実行時の動作のしくみ

次の図は、Microsoft Graph へのサインインおよびアクセスの処理がどのように行われるのかを示しています。

![SSO プロセスを示す図。](../images/sso-access-to-microsoft-graph.png)

1. アドインで、JavaScript は新しい Office.js API [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getAccessToken_options_) を呼び出します。 これにより、Office クライアント アプリケーションはアドインへのアクセス トークンを取得するように指示されます。 (これは処理の後半で 2 つ目のトークンと置換されるので、これ以降では **ブートストラップ アクセス トークン** と呼びます。 デコードされたブートストラップ アクセス トークンの例については、[アクセス トークンの例](sso-in-office-add-ins.md#example-access-token)に関するページを参照してください)。
2. ユーザーがサインインしていない場合、Office クライアント アプリケーションはユーザーにサインインを求めるポップアップ ウィンドウを開きます。
3. 現在のユーザーが初めてアドインを使用する場合は、そのユーザーに同意を求めるダイアログを表示します。
4. クライアント Officeアプリケーションは、現在のユーザーのazure AD v2.0 エンドポイントからブートストラップ アクセス トークンを要求します。
5. Azure ADクライアント アプリケーションにブートストラップ トークンOffice送信します。
6. クライアント Officeは、呼び出しによって返される結果オブジェクトの一部として、ブートストラップ アクセス トークンをアドインに送信 `getAccessToken` します。
7. アドインの JavaScript は、アドインと同じ完全修飾ドメイン名でホストされている Web API に HTTP 要求を送信して、その要求に承認の証拠として **ブートストラップ アクセス トークン** を含めます。
8. 受け取った **ブートストラップ アクセス トークン** はサーバー側のコードで検証されます。
9. サーバー側コードは、ブートストラップ アクセス トークンと引き換えに、"代理" フロー [(OAuth2 Token Exchange](https://tools.ietf.org/html/draft-ietf-oauth-token-exchange-02)およびデーモンまたはサーバー アプリケーションから[Web API Azure](/azure/active-directory/develop/active-directory-authentication-scenarios)シナリオに定義) を使用して、Microsoft Graph のアクセス トークンを取得します。
10. Azure AD は、アドインに Microsoft Graph へのアクセス トークンを返します (アドインが *offline_access* アクセス許可を要求した場合は、更新トークンも返します)。
11. サーバー側のコードが Microsoft Graph へのアクセス トークンをキャッシュします。
12. サーバー側のコードが Microsoft Graph に要求を送信します。そのコードには、Microsoft Graph へのアクセス トークンを含めます。
13. Microsoft Graph は、アドインの UI に渡されるデータをアドインに返します。
14. Microsoft Graph へのアクセス トークンの期限が切れた場合、サーバー側のコードがその更新トークンを使用して、Microsoft Graph への新しいアクセス トークンを取得します。

## <a name="develop-an-sso-add-in-that-accesses-microsoft-graph"></a>Microsoft Graph にアクセスする SSO を開発する

Microsoft Graph にアクセスするアドインは、SSO を使用する他のすべてのアドインと同様に開発できます。 完全な説明については、「[Office アドインのシングル サインオンを有効化する](../develop/sso-in-office-add-ins.md)」を参照してください。違いは、アドインにサーバー側の Web API があることが必須であることと、この記事でアクセス トークンと呼ばれているものは “ブートストラップ アクセス トークン” と呼ばれているという点です。

使用する言語とフレームワークによっては、記述する必要のあるサーバー側コードが簡単になるライブラリが使用できることがあります。 コードでは、次に示す操作を実行する必要があります。

* ブートストラップ アクセス トークン、ユーザーに関するメタデータ、アドインの資格情報 (ID とシークレット) を含む Azure AD v2.0 エンドポイントへの呼び出しで、"代理" フローを開始します。
* Microsoft Graph に (おそらくキャッシュされる) アクセス トークンを渡し、Microsoft Graph データを取得する 1 つ以上の Web API メソッドを作成します。
* 必要に応じて、フローを開始する前に、以前に作成したトークン ハンドラーから受け取ったアドイン ブートストラップ アクセス トークンを検証します。 詳細については、[アクセス トークンの検証](sso-in-office-add-ins.md#validate-the-access-token)に関するページを参照してください。 
* 必要に応じて、フローが完了した後に、Microsoft Graph のアクセス トークンをキャッシュします。 アドインで Microsoft Graph を複数回呼び出す場合に、これが行われます。 このフローの詳細については、「[Azure Active Directory v2.0 と OAuth 2.0 の On-Behalf-Of フロー](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of)」を参照してください。

> [!NOTE]
> “代理” フローで取得した Microsoft Graph のデコードされたアクセス トークンの例については、「[Azure Active Directory v2.0 と OAuth 2.0 の On-Behalf-Of フロー](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of)」を参照してください。

詳細なチュートリアルとシナリオについては、次を参照してください。

* [シングル サインオンを使用する Node.js Office アドインを作成する](create-sso-office-add-ins-nodejs.md)
* [シングル サインオンを使用する ASP.NET Office アドインを作成する](create-sso-office-add-ins-aspnet.md)
* [シナリオ: Outlook アドインでサービスにシングル サインオンを実装する](../outlook/implement-sso-in-outlook-add-in.md)

## <a name="distributing-sso-enabled-add-ins-in-microsoft-appsource"></a>Microsoft AppSource での SSO 対応アドインの配布

管理者Microsoft 365 [AppSource](https://appsource.microsoft.com)からアドインを取得すると、管理者は統合アプリを通じてアドインを再配布し[](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps)、Microsoft Graph スコープにアクセスするためのアドインに対する管理者の同意を付与できます。 ただし、エンド ユーザーが AppSource から直接アドインを取得することもできます。その場合、ユーザーはアドインに同意する必要があります。 これにより、ソリューションを提供した潜在的なパフォーマンスの問題が発生する可能性があります。

コードが 、(など) の呼び出しでオプションを渡した場合 `allowConsentPrompt` `getAccessToken` `OfficeRuntime.auth.getAccessToken( { allowConsentPrompt: true } );` 、Azure AD がアドインにまだ同意が与えされていないことを Office に報告した場合、Office はユーザーに同意を求めるメッセージを表示できます。 ただし、セキュリティ上の理由から、Officeユーザーに対して Azure のスコープへの同意のみを求AD `profile` できます。 *Office、Microsoft* のスコープへの同意を求Graphを求め、それもできません `User.Read` 。 つまり、ユーザーがプロンプトに同意を与える場合、Officeブートストラップ トークンが返されます。 ただし、ブートストラップ トークンを Microsoft Graph へのアクセス トークンと交換しようとすると、エラー AADSTS65001 が失敗します。つまり、同意 (Microsoft Graph スコープへの) が付与されていないという意味です。

コードでは、別の認証システムに戻ってこのエラーを処理できます。このエラーを処理すると、Microsoft のスコープに対する同意を求Graphがあります。 (コード例については、「シングル サインオンを使用する[Node.js Office](create-sso-office-add-ins-nodejs.md)アドインを作成する」および「シングル サインオンを使用する[ASP.NET Office](create-sso-office-add-ins-aspnet.md)アドインを作成する」とリンク先のサンプルを参照してください。ただし、全体のプロセスでは、Azure への複数のラウンド トリップが必要AD。 このパフォーマンスの低下は、オプションを呼び出しに含めて回避 `forMSGraphAccess` できます。たとえば `getAccessToken` 、 `OfficeRuntime.auth.getAccessToken( { forMSGraphAccess: true } )` 。  これは、アドインOfficeスコープに Microsoft が必要Graphします。 Office Azure ADに対して、Microsoft Graphスコープへの同意が既にアドインに付与されているのを確認するように求めるメッセージが表示されます。 ブートストラップ トークンがある場合は、ブートストラップ トークンが返されます。 呼び出しが行えない場合は、 `getAccessToken` エラー 13012 が返されます。 コードは、Azure サーバーとトークンを交換する運命の試みを行わずに、すぐに別の認証システムに戻ってこのエラーを処理AD。

ベスト プラクティスとして、アドインが AppSource で配布され、Microsoft のスコープが必要になるGraph `forMSGraphAccess` `getAccessToken` します。

> [!TIP]
> SSO を使用する Outlook アドインを開発し、テスト用にサイドロードすると、管理者の同意が与えらた場合でも、Office は常にエラー 13012 を返します。 `forMSGraphAccess` `getAccessToken` このため、アドインを開発するときにオプションをコメント `forMSGraphAccess` アウトOutlook必要があります。 実稼働環境に展開する場合は、必ずオプションのアンコメントを解除してください。 偽の 13012 は、このページでサイドローディングを行Outlook。
