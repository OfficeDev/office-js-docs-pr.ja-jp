---
title: SSO を使用した Microsoft Graph への承認
description: Office アドインのユーザーがシングルサインオン (SSO) を使用して Microsoft Graph からデータを取得する方法について説明します。
ms.date: 07/30/2020
localization_priority: Normal
ms.openlocfilehash: 68440a347e11d909f0ebd4d4d29892711646da5e
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292905"
---
# <a name="authorize-to-microsoft-graph-with-sso"></a>SSO を使用した Microsoft Graph への承認

ユーザーは個人用の Microsoft アカウントまたは Microsoft 365 Education または職場アカウントのいずれかを使用して、Office (オンライン、モバイル、およびデスクトップ プラットフォーム) にサインインします。 Office アドインの [Microsoft Graph](https://developer.microsoft.com/graph/docs) へのアクセスの承認には、ユーザーの Office サインオン資格証明を使用するのが最良の方法です。 これにより、2 回目はサインインする必要なく Microsoft Graph データにアクセスできます。

> [!NOTE]
> 現在、シングル サインオン API は Word、Excel、Outlook, および PowerPoint でサポートされています。 シングル サインオン API の現在のサポート状態に関する詳細は、「[IdentityAPI の要件セット](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets)」を参照してください。
> Outlook アドインで作業している場合は、Office 365 テナントの先進認証が有効になっていることを確認してください。 この方法の詳細については、「[Exchange Online: テナントの先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)」を参照してください。


## <a name="add-in-architecture-for-sso-and-microsoft-graph"></a>SSO と Microsoft Graph のアドイン アーキテクチャ

Web アプリケーションのページと JavaScript をホスティングすることに加え、アドインでは、同一の[完全修飾ドメイン名](/windows/desktop/DNS/f-gly#_dns_fully_qualified_domain_name_fqdn__gly)で、Microsoft Graph へのアクセス トークンを取得して、要求を送信するための 1 つ以上の Web API をホストする必要もあります。

アドイン マニフェストには、アドインを Azure Active Directory (Azure AD) v2.0 エンドポイントに登録する方法と、アドインが必要とする Microsoft Graph へのアクセス許可を指定するマークアップが含まれています。

### <a name="how-it-works-at-runtime"></a>実行時の動作のしくみ

次の図は、Microsoft Graph へのサインインおよびアクセスの処理がどのように行われるのかを示しています。

![SSO プロセスを示す図](../images/sso-access-to-microsoft-graph.png)

1. アドインで、JavaScript は新しい Office.js API [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-) を呼び出します。 これにより、Office クライアントアプリケーションにアドインへのアクセストークンを取得するように指示します。 (これは処理の後半で 2 つ目のトークンと置換されるので、これ以降では**ブートストラップ アクセス トークン**と呼びます。 デコードされたブートストラップ アクセス トークンの例については、[アクセス トークンの例](sso-in-office-add-ins.md#example-access-token)に関するページを参照してください)。
2. ユーザーがサインインしていない場合、Office クライアントアプリケーションはユーザーにサインインを求めるポップアップウィンドウを開きます。
3. 現在のユーザーが初めてアドインを使用する場合は、そのユーザーに同意を求めるダイアログを表示します。
4. Office クライアントアプリケーションは、現在のユーザーの Azure AD v2.0 エンドポイントから **ブートストラップアクセストークン** を要求します。
5. Azure AD は、Office クライアントアプリケーションにブートストラップトークンを送信します。
6. Office クライアントアプリケーションは、呼び出しによって返される result オブジェクトの一部として、アドインに **ブートストラップアクセストークン** を送信し `getAccessToken` ます。
7. アドインの JavaScript は、アドインと同じ完全修飾ドメイン名でホストされている Web API に HTTP 要求を送信して、その要求に承認の証拠として**ブートストラップ アクセス トークン**を含めます。
8. 受け取った**ブートストラップ アクセス トークン**はサーバー側のコードで検証されます。
9. サーバー側のコードでは、"代理として定義される" フロー ( [OAuth2 Token Exchange](https://tools.ietf.org/html/draft-ietf-oauth-token-exchange-02) 、 [デーモンまたはサーバーアプリケーションから web API Azure シナリオ](/azure/active-directory/develop/active-directory-authentication-scenarios)) を使用して、ブートストラップアクセストークンの Exchange で Microsoft Graph のアクセストークンを取得します。
10. Azure AD は、アドインに Microsoft Graph へのアクセス トークンを返します (アドインが *offline_access* アクセス許可を要求した場合は、更新トークンも返します)。
11. サーバー側のコードが Microsoft Graph へのアクセス トークンをキャッシュします。
12. サーバー側のコードが Microsoft Graph に要求を送信します。そのコードには、Microsoft Graph へのアクセス トークンを含めます。
13. Microsoft Graph は、アドインの UI に渡すことができるデータをアドインに返します。
14. Microsoft Graph へのアクセス トークンの期限が切れた場合、サーバー側のコードがその更新トークンを使用して、Microsoft Graph への新しいアクセス トークンを取得します。

## <a name="develop-an-sso-add-in-that-accesses-microsoft-graph"></a>Microsoft Graph にアクセスする SSO を開発する

Microsoft Graph にアクセスするアドインは、SSO を使用する他のすべてのアドインと同様に開発できます。 完全な説明については、「[Office アドインのシングル サインオンを有効化する](../develop/sso-in-office-add-ins.md)」を参照してください。違いは、アドインにサーバー側の Web API があることが必須であることと、この記事でアクセス トークンと呼ばれているものは “ブートストラップ アクセス トークン” と呼ばれているという点です。

使用する言語とフレームワークによっては、記述する必要のあるサーバー側コードが簡単になるライブラリが使用できることがあります。 コードでは、次に示す操作を実行する必要があります。

* ブートストラップアクセストークン、ユーザーに関するメタデータ、およびアドインの資格情報 (ID とシークレット) を含む Azure AD v2.0 エンドポイントへの呼び出しを使用して、"代理" フローを開始します。
* Microsoft Graph に (おそらくキャッシュされる) アクセス トークンを渡し、Microsoft Graph データを取得する 1 つ以上の Web API メソッドを作成します。
* 必要に応じて、フローを開始する前に、以前に作成したトークン ハンドラーから受け取ったアドイン ブートストラップ アクセス トークンを検証します。 詳細については、[アクセス トークンの検証](sso-in-office-add-ins.md#validate-the-access-token)に関するページを参照してください。 
* 必要に応じて、フローが完了した後に、Microsoft Graph のアクセス トークンをキャッシュします。 アドインで Microsoft Graph を複数回呼び出す場合に、これが行われます。 このフローの詳細については、「[Azure Active Directory v2.0 と OAuth 2.0 の On-Behalf-Of フロー](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of)」を参照してください。

> [!NOTE]
> “代理” フローで取得した Microsoft Graph のデコードされたアクセス トークンの例については、「[Azure Active Directory v2.0 と OAuth 2.0 の On-Behalf-Of フロー](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of)」を参照してください。

詳細なチュートリアルとシナリオについては、次を参照してください。

* [シングル サインオンを使用する Node.js Office アドインを作成する](create-sso-office-add-ins-nodejs.md)
* [シングル サインオンを使用する ASP.NET Office アドインを作成する](create-sso-office-add-ins-aspnet.md)
* [シナリオ: Outlook アドインでサービスにシングル サインオンを実装する](../outlook/implement-sso-in-outlook-add-in.md)

## <a name="distributing-sso-enabled-add-ins-in-microsoft-appsource"></a>Microsoft AppSource で SSO が有効なアドインを配布する

Microsoft 365 管理者が [Appsource](https://appsource.microsoft.com)からアドインを取得すると、管理者は [一元展開](../publish/centralized-deployment.md) によってアドインを再配布し、microsoft Graph スコープにアクセスするためにアドインに管理者の同意を与えることができます。 ただし、エンドユーザーがアドインを AppSource から直接取得することもできます。この場合、ユーザーはアドインに同意を与える必要があります。 これにより、ソリューションを提供したパフォーマンスの問題が発生する可能性があります。

`allowConsentPrompt`のように、の呼び出しでオプションを渡した場合 `getAccessToken` 、office は、 `OfficeRuntime.auth.getAccessToken( { allowConsentPrompt: true } );` 同意がまだアドインに付与されていない場合に、ユーザーに同意を求めるメッセージを表示します。 ただし、セキュリティ上の理由から、Office はユーザーに対して Azure AD スコープへの同意を求めることしかでき `profile` ません。 *Office では、Microsoft Graph のスコープに対する同意を求めるダイアログを表示することはできません* `User.Read` 。 これは、ユーザーがプロンプトに同意を与えると、Office はブートストラップトークンを返します。 しかし、Microsoft graph へのアクセストークンのブートストラップトークンの交換は、エラー AADSTS65001 で失敗します。つまり、同意 (Microsoft Graph のスコープへのアクセス許可) が付与されていません。

コードでは、別の認証システムにフォールバックすることによって、このエラーを処理することができます。これにより、Microsoft Graph スコープへの同意を求めるメッセージが表示されます。 コード例については、「 [シングルサインオンを使用する Office アドインを Node.js 作成](create-sso-office-add-ins-nodejs.md) し、 [シングルサインオンを使用する ASP.NET Office アドインを作成](create-sso-office-add-ins-aspnet.md) する」と、リンク先のサンプルを参照してください。ただし、プロセス全体では、Azure AD への複数のラウンドトリップが必要です。 の呼び出しにオプションを含めることによって、このパフォーマンスペナルティを避けることができ `forMSGraphAccess` `getAccessToken` ます (例:) `OfficeRuntime.auth.getAccessToken( { forMSGraphAccess: true } )` 。  これにより、アドインに Microsoft Graph のスコープが必要であることが Office に通知されます。 Office は、Microsoft Graph スコープへの同意がアドインに既に付与されているかどうかを確認するために、Azure AD に要求します。 その場合は、ブートストラップトークンが返されます。 存在しない場合は、の呼び出しで `getAccessToken` エラー13012が返されます。 コードでこのエラーを処理するには、Azure AD を使用してトークンを交換することなく、すぐに認証の代替システムにフォールバックします。

ベストプラクティスとして、 `forMSGraphAccess` `getAccessToken` アドインが appsource で配布され、Microsoft Graph のスコープが必要になる場合は、常にを渡します。

> [!TIP]
> SSO を使用する Outlook アドインを開発し、テスト用にサイドロードすると、 *always* `forMSGraphAccess` `getAccessToken` 管理者の同意が与えられている場合でも、Office は常にエラー13012を返します。 このため、 `forMSGraphAccess` Outlook アドインを **開発する際に** は、オプションをコメントにする必要があります。 運用のために展開するときは、オプションのコメントを解除してください。 Bogus 13012 は、Outlook のサイドロード時にのみ発生します。
