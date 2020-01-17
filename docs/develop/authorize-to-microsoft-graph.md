---
title: SSO を使用した Microsoft Graph への承認
description: ''
ms.date: 01/14/2020
localization_priority: Priority
ms.openlocfilehash: b4c813bb3ffd2bf9025ac4cbf51096e1d7a72d8b
ms.sourcegitcommit: 212c810f3480a750df779777c570159a7f76054a
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 01/17/2020
ms.locfileid: "41217352"
---
# <a name="authorize-to-microsoft-graph-with-sso-preview"></a>SSO を使用した Microsoft Graph への承認 (プレビュー)

ユーザーは個人用の Microsoft アカウントまたは職場や学校の (Office 365) アカウントのいずれかを使用して、Office (オンライン、モバイル、およびデスクトップ プラットフォーム) にサインインします。 Office アドインの [Microsoft Graph](https://developer.microsoft.com/graph/docs) へのアクセスの承認には、ユーザーの Office サインオン資格証明を使用するのが最良の方法です。 これにより、2 回目はサインインする必要なく Microsoft Graph データにアクセスできます。 

> [!NOTE]
> 現在、シングル サインオン API は Word、Excel、Outlook、PowerPoint のプレビューでサポートされています。 シングル サインオン API の現在のサポート状態に関する詳細は、「[IdentityAPI の要件セット](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets)」を参照してください。 Outlook アドインで作業している場合は、Office 365 テナントの先進認証が有効になっていることを確認してください。 この方法の詳細については、「[Exchange Online: テナントの先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)」を参照してください。

## <a name="add-in-architecture-for-sso-and-microsoft-graph"></a>SSO と Microsoft Graph のアドイン アーキテクチャ

Web アプリケーションのページと JavaScript をホスティングすることに加え、アドインでは、同一の[完全修飾ドメイン名](/windows/desktop/DNS/f-gly#_dns_fully_qualified_domain_name_fqdn__gly)で、Microsoft Graph へのアクセス トークンを取得して、要求を送信するための 1 つ以上の Web API をホストする必要もあります。

アドイン マニフェストには、アドインを Azure Active Directory (Azure AD) v2.0 エンドポイントに登録する方法と、アドインが必要とする Microsoft Graph へのアクセス許可を指定するマークアップが含まれています。

### <a name="how-it-works-at-runtime"></a>実行時の動作のしくみ

次の図は、Microsoft Graph へのサインインおよびアクセスの処理がどのように行われるのかを示しています。

![SSO プロセスを示す図](../images/sso-access-to-microsoft-graph.png)

1. アドインで、JavaScript は新しい Office.js API [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-) を呼び出します。 これにより、Office ホスト アプリケーションはアドインへのアクセス トークンを取得するように指示されます  (これは処理の後半で 2 つ目のトークンと置換されるので、これ以降では**ブートストラップ アクセス トークン**と呼びます。 デコードされたブートストラップ アクセス トークンの例については、[アクセス トークンの例](sso-in-office-add-ins.md#example-access-token)に関するページを参照してください)。
2. ユーザーがサインインしていない場合、ユーザーにサインインを求めるポップアップ ウィンドウが Office ホスト アプリケーションによって開かれます。
3. 現在のユーザーが初めてアドインを使用する場合は、そのユーザーに同意を求めるダイアログが表示されます。
4. Office ホスト アプリケーションは、Azure AD v2.0 エンドポイントから現在のユーザーの**ブートストラップ アクセス トークン**を要求します。
5. Azure AD が、Office ホスト アプリケーションにブートストラップ アクセス トークンを送信します。
6. Office ホスト アプリケーションが、`getAccessToken` 呼び出しによって返される結果オブジェクトの一部として、アドインに**ブートストラップ アクセス トークン**を送信します。
7. アドインの JavaScript は、アドインと同じ完全修飾ドメイン名でホストされている Web API に HTTP 要求を送信して、その要求に承認の証拠として**ブートストラップ アクセス トークン**を含めます。
8. 受け取った**ブートストラップ アクセス トークン**はサーバー側のコードで検証されます。
9. サーバー側のコードは、「代理」フロー (「[OAuth2 Token Exchange](https://tools.ietf.org/html/draft-ietf-oauth-token-exchange-02)」と「[デーモンまたはサーバー アプリケーション対 Web API](/azure/active-directory/develop/active-directory-authentication-scenarios)」で定義されているフロー) を使用して、ブートストラップ アクセス トークンと引き換えに Microsoft Graph のアクセス トークンを取得します。
10. Azure AD は、アドインに Microsoft Graph へのアクセス トークンを返します (アドインが *offline_access* アクセス許可を要求した場合は、更新トークンも返します)。
11. サーバー側のコードが Microsoft Graph へのアクセス トークンをキャッシュします。
12. サーバー側のコードが Microsoft Graph に要求を送信します。そのコードには、Microsoft Graph へのアクセス トークンを含めます。
13. Microsoft Graph は、アドインの UI に渡されるデータをアドインに返します。
14. Microsoft Graph へのアクセス トークンの期限が切れた場合、サーバー側のコードがその更新トークンを使用して、Microsoft Graph への新しいアクセス トークンを取得します。

## <a name="develop-an-sso-add-in-that-accesses-microsoft-graph"></a>Microsoft Graph にアクセスする SSO を開発する

Microsoft Graph にアクセスするアドインは、SSO を使用する他のすべてのアドインと同様に開発できます。 完全な説明については、「[Office アドインのシングル サインオンを有効化する](/office/dev/add-ins/develop/sso-in-office-add-ins)」を参照してください。違いは、アドインにサーバー側の Web API があることが必須であることと、この記事でアクセス トークンと呼ばれているものは “ブートストラップ アクセス トークン” と呼ばれているという点です。

使用する言語とフレームワークによっては、記述する必要のあるサーバー側コードが簡単になるライブラリが使用できることがあります。 コードでは、次に示す操作を実行する必要があります。

* Azure AD v2.0 エンドポイントを呼び出して、“代理” フローを開始します。これには、ブートストラップ アクセス トークン、ユーザーに関するメタデータ、およびアドインの資格情報 (ID とシークレット) を含めます。
* Microsoft Graph に (おそらくキャッシュされる) アクセス トークンを渡し、Microsoft Graph データを取得する 1 つ以上の Web API メソッドを作成します。
* 必要に応じて、フローを開始する前に、以前に作成したトークン ハンドラーから受け取ったアドイン ブートストラップ アクセス トークンを検証します。 詳細については、[アクセス トークンの検証](sso-in-office-add-ins.md#validate-the-access-token)に関するページを参照してください。 
* 必要に応じて、フローが完了した後に、Microsoft Graph のアクセス トークンをキャッシュします。 アドインで Microsoft Graph を複数回呼び出す場合に、これが行われます。 このフローの詳細については、「[Azure Active Directory v2.0 と OAuth 2.0 の On-Behalf-Of フロー](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of)」を参照してください。

> [!NOTE]
> “代理” フローで取得した Microsoft Graph のデコードされたアクセス トークンの例については、「[Azure Active Directory v2.0 と OAuth 2.0 の On-Behalf-Of フロー](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of)」を参照してください。

詳細なチュートリアルとシナリオについては、次を参照してください。

* [シングル サインオンを使用する Node.js Office アドインを作成する](create-sso-office-add-ins-nodejs.md)
* [シングル サインオンを使用する ASP.NET Office アドインを作成する](create-sso-office-add-ins-aspnet.md)
* [シナリオ: Outlook アドインでサービスにシングル サインオンを実装する](/outlook/add-ins/implement-sso-in-outlook-add-in)
