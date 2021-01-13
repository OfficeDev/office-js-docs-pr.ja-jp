---
title: SSO を使用した Microsoft Graph への承認
description: アドインのユーザー Officeシングル サインオン (SSO) を使用して Microsoft Graph からデータを取得する方法について説明します。
ms.date: 07/30/2020
localization_priority: Normal
ms.openlocfilehash: d6d06b9d7ff42b72495f513ed2c6b8f6f36df1d0
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839972"
---
# <a name="authorize-to-microsoft-graph-with-sso"></a>SSO を使用した Microsoft Graph への承認

ユーザーは個人用の Microsoft アカウントまたは Microsoft 365 Education または職場アカウントのいずれかを使用して、Office (オンライン、モバイル、およびデスクトップ プラットフォーム) にサインインします。 Office アドインの [Microsoft Graph](https://developer.microsoft.com/graph/docs) へのアクセスの承認には、ユーザーの Office サインオン資格証明を使用するのが最良の方法です。 これにより、2 回目はサインインする必要なく Microsoft Graph データにアクセスできます。

> [!NOTE]
> 現在、シングル サインオン API は Word、Excel、Outlook, および PowerPoint でサポートされています。 シングル サインオン API の現在のサポート状態に関する詳細は、「[IdentityAPI の要件セット](../reference/requirement-sets/identity-api-requirement-sets.md)」を参照してください。
> Outlook アドインで作業している場合は、Office 365 テナントの先進認証が有効になっていることを確認してください。 この方法の詳細については、「[Exchange Online: テナントの先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)」を参照してください。

## <a name="add-in-architecture-for-sso-and-microsoft-graph"></a>SSO と Microsoft Graph のアドイン アーキテクチャ

Web アプリケーションのページと JavaScript をホスティングすることに加え、アドインでは、同一の[完全修飾ドメイン名](/windows/desktop/DNS/f-gly#_dns_fully_qualified_domain_name_fqdn__gly)で、Microsoft Graph へのアクセス トークンを取得して、要求を送信するための 1 つ以上の Web API をホストする必要もあります。

アドイン マニフェストには、アドインを Azure Active Directory (Azure AD) v2.0 エンドポイントに登録する方法と、アドインが必要とする Microsoft Graph へのアクセス許可を指定するマークアップが含まれています。

### <a name="how-it-works-at-runtime"></a>実行時の動作のしくみ

次の図は、Microsoft Graph へのサインインおよびアクセスの処理がどのように行われるのかを示しています。

![SSO プロセスを示す図](../images/sso-access-to-microsoft-graph.png)

1. アドインで、JavaScript は新しい Office.js API [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-) を呼び出します。 これにより、Office クライアント アプリケーションはアドインへのアクセス トークンを取得するように指示されます。 (これは処理の後半で 2 つ目のトークンと置換されるので、これ以降では **ブートストラップ アクセス トークン** と呼びます。 デコードされたブートストラップ アクセス トークンの例については、[アクセス トークンの例](sso-in-office-add-ins.md#example-access-token)に関するページを参照してください)。
2. ユーザーがサインインしていない場合、Office クライアント アプリケーションはユーザーにサインインを求めるポップアップ ウィンドウを開きます。
3. 現在のユーザーが初めてアドインを使用する場合は、そのユーザーに同意を求めるダイアログを表示します。
4. クライアント Officeは、現在のユーザーのAzure AD v2.0 エンドポイントからブートストラップ アクセス トークンを要求します。
5. Azure ADクライアント アプリケーションにブートストラップ トークンOffice送信します。
6. クライアント Officeは、呼び出しによって返される結果オブジェクトの一部としてブートストラップ アクセス トークンをアドインに送信 `getAccessToken` します。
7. アドインの JavaScript は、アドインと同じ完全修飾ドメイン名でホストされている Web API に HTTP 要求を送信して、その要求に承認の証拠として **ブートストラップ アクセス トークン** を含めます。
8. 受け取った **ブートストラップ アクセス トークン** はサーバー側のコードで検証されます。
9. サーバー側のコードは [、(OAuth2 Token Exchange](https://tools.ietf.org/html/draft-ietf-oauth-token-exchange-02) およびデーモンまたはサーバー アプリケーションから [Web API への Azure](/azure/active-directory/develop/active-directory-authentication-scenarios)シナリオで定義されている) "代理" フローを使用して、ブートストラップ アクセス トークンと引き換えに Microsoft Graph のアクセス トークンを取得します。
10. Azure AD は、アドインに Microsoft Graph へのアクセス トークンを返します (アドインが *offline_access* アクセス許可を要求した場合は、更新トークンも返します)。
11. サーバー側のコードが Microsoft Graph へのアクセス トークンをキャッシュします。
12. サーバー側のコードが Microsoft Graph に要求を送信します。そのコードには、Microsoft Graph へのアクセス トークンを含めます。
13. Microsoft Graph はアドインにデータを返し、それをアドインの UI に渡します。
14. Microsoft Graph へのアクセス トークンの期限が切れた場合、サーバー側のコードがその更新トークンを使用して、Microsoft Graph への新しいアクセス トークンを取得します。

## <a name="develop-an-sso-add-in-that-accesses-microsoft-graph"></a>Microsoft Graph にアクセスする SSO を開発する

Microsoft Graph にアクセスするアドインは、SSO を使用する他のすべてのアドインと同様に開発できます。 完全な説明については、「[Office アドインのシングル サインオンを有効化する](../develop/sso-in-office-add-ins.md)」を参照してください。違いは、アドインにサーバー側の Web API があることが必須であることと、この記事でアクセス トークンと呼ばれているものは “ブートストラップ アクセス トークン” と呼ばれているという点です。

使用する言語とフレームワークによっては、記述する必要のあるサーバー側コードが簡単になるライブラリが使用できることがあります。 コードでは、次に示す操作を実行する必要があります。

* ブートストラップ アクセス トークン、ユーザーに関するメタデータ、アドインの資格情報 (ID とシークレット) を含む Azure AD v2.0 エンドポイントを呼び出して、"代理" フローを開始します。
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

Microsoft 365 管理者が[AppSource](https://appsource.microsoft.com)からアドインを取得すると、管理者は一元展開によって[](../publish/centralized-deployment.md)アドインを再配布し、Microsoft Graph スコープにアクセスするための管理者の同意をアドインに付与できます。 ただし、エンド ユーザーが AppSource から直接アドインを取得することもできます。その場合、ユーザーはアドインに同意する必要があります。 これにより、ソリューションを提供した場合の潜在的なパフォーマンスの問題が発生する可能性があります。

コードが "like" の呼び出しでオプションを渡した場合、Azure AD がアドインにまだ同意が付与されていないことを Office に報告した場合、Office はユーザーに同意を求めるメッセージを表示できます。 `allowConsentPrompt` `getAccessToken` `OfficeRuntime.auth.getAccessToken( { allowConsentPrompt: true } );` ただし、セキュリティ上の理由からOffice、Azure のスコープに同意するようユーザーにAD `profile` できます。 *Office Microsoft Graph のスコープに同意* を求め、それも `User.Read` できません。 つまり、ユーザーがプロンプトに同意すると、Officeが返されます。 ただし、Microsoft Graph へのアクセス トークンのブートストラップ トークンを交換しようとすると、エラー AADSTS65001 で失敗します。これは、(Microsoft Graph スコープに対する) 同意が付与されていないという意味です。

コードでは、認証の代替システムに戻ってこのエラーを処理できます。また、このエラーを処理する必要があります。これにより、Microsoft Graph スコープへの同意を求めるメッセージがユーザーに表示されます。 (コード例については、「シングル サインオンを使用する [Node.js Office](create-sso-office-add-ins-nodejs.md) アドインを作成する」および「シングル サインオンとリンク先のサンプルを使用する [ASP.NET Office](create-sso-office-add-ins-aspnet.md) アドインを作成する」を参照してください)。ただし、プロセス全体で Azure への複数回のラウンド トリップがAD。 このパフォーマンスの低下を回避するには、呼び出しに `forMSGraphAccess` `getAccessToken` オプションを含める必要があります `OfficeRuntime.auth.getAccessToken( { forMSGraphAccess: true } )` 。  これは、アドインOffice Microsoft Graph スコープが必要な場合に、そのスコープを示します。 Officeは、Microsoft Graph ADに対する同意が既にアドインに付与されているのを確認するように Azure ADに求めるメッセージを表示します。 設定されている場合は、ブートストラップ トークンが返されます。 呼び出しが失敗した場合は、 `getAccessToken` エラー 13012 が返されます。 コードでこのエラーを処理するには、Azure AD とトークンをやり取りする不必要な試みを行わずに、認証の代替システムにすぐに戻AD。

ベスト プラクティスとして、アドインが AppSource で配布され、Microsoft Graph スコープが必要になる場合に常に `forMSGraphAccess` `getAccessToken` 渡します。

> [!TIP]
> SSO を使用する Outlook アドインを開発し、テスト用にサイドロードすると、管理者の同意が与えらた場合でも、Office は常にエラー 13012 を返します。 `forMSGraphAccess` `getAccessToken` このため、Outlook アドインを開発するときにオプション `forMSGraphAccess` をコメント アウトする必要があります。 実稼働環境に展開する場合は、必ずオプションの非コミットを行います。 偽の 13012 は、Outlook でサイドロードしている場合にのみ発生します。