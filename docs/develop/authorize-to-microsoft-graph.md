---
title: Office アドインの Microsoft Graph への承認
description: ''
ms.date: 04/10/2018
ms.openlocfilehash: 7631be6900020e4be78a8590b3b1c237088eea02
ms.sourcegitcommit: 3d8454055ba4d7aae12f335def97357dea5beb30
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/14/2018
ms.locfileid: "27270671"
---
# <a name="authorize-to-microsoft-graph-in-your-office-add-in-preview"></a>Office アドインの Microsoft Graph への承認 (プレビュー)

ユーザーは個人用の Microsoft アカウントまたは職場や学校の (Office 365) アカウントのいずれかを使用して、Office (オンライン、モバイル、およびデスクトップ プラットフォーム) にサインインします。 Office アドインの [Microsoft Graph](https://developer.microsoft.com/graph/docs) へのアクセスの承認には、ユーザーの Office サインオン資格証明を使用するのが最良の方法です。 これにより、2 回目はサインインする必要なく Microsoft Graph データにアクセスできます。 

> [!NOTE]
> 現在、シングル サインオン API は Word、Excel、Outlook、PowerPoint のプレビューでサポートされています。 シングル サインオン API の現在のサポート状態に関する詳細は、「[IdentityAPI の要件セット](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js)」を参照してください。
> Outlook アドインで作業している場合は、Office 365 テナントの先進認証が有効になっていることを確認してください。 この方法の詳細については、「[Exchange Online: テナントの先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)」を参照してください。

## <a name="add-in-architecture-for-sso-and-microsoft-graph"></a>SSO と Microsoft Graph のアドイン アーキテクチャ

Web アプリケーションのページと JavaScript をホスティングすることに加え、アドインでは、同一の[完全修飾ドメイン名](https://docs.microsoft.com/windows/desktop/DNS/f-gly#_dns_fully_qualified_domain_name_fqdn__gly)で、Microsoft Graph へのアクセス トークンを取得して、要求を送信するための 1 つ以上の Web API をホストする必要もあります。

アドイン マニフェストには、アドインを Azure Active Directory (Azure AD) v2.0 エンドポイントに登録する方法と、アドインが必要とする Microsoft Graph へのアクセス許可を指定するマークアップが含まれています。

### <a name="how-it-works-at-runtime"></a>実行時の動作のしくみ

次の図は、Microsoft Graph へのサインインおよびアクセスの処理がどのように行われるのかを示しています。

![SSO プロセスを示す図](../images/sso-access-to-microsoft-graph.png)

1. アドインの JavaScript により、新しい Office.js API [getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) が呼び出されます。 これにより、Office ホスト アプリケーションはアドインへのアクセス トークンを取得するように指示されます  (これは処理の後半で 2 つ目のトークンと置換されるので、これ以降では**ブートストラップ アクセス トークン**と呼びます。 デコードされたブートストラップ アクセス トークンの例については、[アクセス トークンの例](sso-in-office-add-ins.md#example-access-token)に関するページを参照してください)。
1. ユーザーがサインインしていない場合、ユーザーにサインインを求めるポップアップ ウィンドウが Office ホスト アプリケーションによって開かれます。
1. 現在のユーザーが初めてアドインを使用する場合は、そのユーザーに同意を求めるダイアログが表示されます。
1. Office ホスト アプリケーションは、Azure AD v2.0 エンドポイントから現在のユーザーの**ブートストラップ アクセス トークン**を要求します。
1. Azure AD が、Office ホスト アプリケーションにブートストラップ アクセス トークンを送信します。
1. Office ホスト アプリケーションが、`getAccessTokenAsync` 呼び出しによって返される結果オブジェクトの一部として、アドインに**ブートストラップ アクセス トークン**を送信します。
1. アドインの JavaScript は、アドインと同じ完全修飾ドメイン名でホストされている Web API に HTTP 要求を送信して、その要求に承認の証拠として**ブートストラップ アクセス トークン**を含めます。  
1. 受け取った**ブートストラップ アクセス トークン**はサーバー側のコードで検証されます。
1. サーバー側のコードは、「代理」フロー (「[OAuth2 Token Exchange](https://tools.ietf.org/html/draft-ietf-oauth-token-exchange-02)」と「[デーモンまたはサーバー アプリケーション対 Web API](https://docs.microsoft.com/azure/active-directory/develop/active-directory-authentication-scenarios)」で定義されているフロー) を使用して、ブートストラップ アクセス トークンと引き換えに Microsoft Graph のアクセス トークンを取得します。
1. Azure AD は、アドインに Microsoft Graph へのアクセス トークンを返します (アドインが *offline_access* アクセス許可を要求した場合は、更新トークンも返します)。
1. サーバー側のコードが Microsoft Graph へのアクセス トークンをキャッシュします。
1. サーバー側のコードが Microsoft Graph に要求を送信します。そのコードには、Microsoft Graph へのアクセス トークンを含めます。
1. Microsoft Graph は、アドインの UI に渡されるデータをアドインに返します。
1. Microsoft Graph へのアクセス トークンの期限が切れた場合、サーバー側のコードがその更新トークンを使用して、Microsoft Graph への新しいアクセス トークンを取得します。

## <a name="develop-an-sso-add-in-that-accesses-microsoft-graph"></a>Microsoft Graph にアクセスする SSO を開発する

Microsoft Graph にアクセスするアドインは、SSO を使用する他のすべてのアドインと同様に開発できます。 完全な説明については、「[Office アドインのシングル サインオンを有効化する](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins)」を参照してください。違いは、アドインにサーバー側の Web API があることが必須であることと、この記事でアクセス トークンと呼ばれているものは “ブートストラップ アクセス トークン” と呼ばれているという点です。 

使用する言語とフレームワークによっては、記述する必要のあるサーバー側コードが簡単になるライブラリが使用できることがあります。 コードでは、次に示す操作を実行する必要があります。

* 前に作成したトークン ハンドラーから受け取ったアドイン ブートストラップ アクセス トークンを検証します。 詳細については、[アクセス トークンの検証](sso-in-office-add-ins.md#validate-the-access-token)に関するページを参照してください。 
* Azure AD v2.0 エンドポイントを呼び出して、“代理” フローを開始します。これには、ブートストラップ アクセス トークン、ユーザーに関するメタデータ、およびアドインの資格情報 (ID とシークレット) を含めます。
* Microsoft Graph に返されたアクセス トークンをキャッシュします。 このフローの詳細については、「[Azure Active Directory v2.0 と OAuth 2.0 の On-Behalf-Of フロー](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of)」を参照してください。
* Microsoft Graph にキャッシュされたアクセス トークンを渡し、Microsoft Graph データを取得する 1 つ以上の Web API メソッドを作成します。

> [!NOTE]
> “代理” フローで取得した Microsoft Graph のデコードされたアクセス トークンの例については、「[Azure Active Directory v2.0 と OAuth 2.0 の On-Behalf-Of フロー](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of)」を参照してください。

詳細なチュートリアルとシナリオについては、次を参照してください。

* [シングル サインオンを使用する Node.js Office アドインを作成する](create-sso-office-add-ins-nodejs.md)
* [シングル サインオンを使用する ASP.NET Office アドインを作成する](create-sso-office-add-ins-aspnet.md)
* [シナリオ: Outlook アドインでサービスにシングル サインオンを実装する](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in)



