---
title: Office アドインで外部サービスを承認する
description: OAuth 2.0 認証コード フローおよび暗黙的フローを使用して、Google、Facebook、LinkedIn、SalesForce、および GitHub などの Microsoft 以外のデータソースに対する承認を取得します。
ms.date: 08/07/2019
localization_priority: Normal
ms.openlocfilehash: fd180e11106e7e1e2f20f539746535c4310ad81e
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093743"
---
# <a name="authorize-external-services-in-your-office-add-in"></a>Office アドインで外部サービスを承認する

Popular online services, including Microsoft 365, Google, Facebook, LinkedIn, SalesForce, and GitHub, let developers give users access to their accounts in other applications. This gives you the ability to include these services in your Office Add-in.

> [!NOTE]
> この記事の残りの部分では、Microsoft 以外のサービスへのアクセスについて説明します。 Microsoft Graph (Microsoft 365 を含む) へのアクセスに関する詳細については、「 [sso を使用した Microsoft graph へのアクセス](overview-authn-authz.md#access-to-microsoft-graph-with-sso)」および「sso を使用し[ない microsoft Graph](overview-authn-authz.md#access-to-microsoft-graph-without-sso)へのアクセス」を参照してください。

The industry standard framework for enabling web application access to an online service is **OAuth 2.0**. In most situations, you don't need to know the details of how the framework works to use it in your add-in. Many libraries are available that simplify the details for you.

OAuth の基本的な考え方は、ユーザーやグループと同様に、アプリケーションは専用の ID とアクセス許可のセットによって、それ自体が[セキュリティ プリンシパル](/windows/security/identity-protection/access-control/security-principals)になり得るということです。 通常のシナリオでは、ユーザーがオンライン サービスを必要とする Office アドインのアクションを実行すると、アドインは、ユーザーのアカウントへ特定のセットのアクセス許可を付与するよう求める要求をサービスに送信します。 サービスは、該当するアクセス許可をアドインに付与するように求めるプロンプトをユーザーに表示します。 アクセス許可が付与されると、サービスは小さなエンコードされた*アクセス トークン*をアドインに送信します。 アドインは、サービスの API へのすべての要求にトークンを含めることで、サービスを使用できるようになります。 ただし、そのアドインが実行できるアクションは、ユーザーが付与したアクセス許可の範囲内に限定されます。 また、トークンは特定の時間が経過すると期限切れになります。

さまざまなシナリオに向けて、*フロー*または*許可の種類*と呼ばれる、いくつかの OAuth パターンが設計されています。 最も一般的な実装パターンは次の 2 つです。

- **暗黙的フロー**: アドインとオンライン サービスとの通信は、クライアント側の JavaScript で実装されます。 このフローは、シングル ページ アプリケーション (SPA) で一般的に使用されます。
- **認証コード フロー**:アドインの Web アプリケーションとオンライン サービスとの通信は、*サーバー間*で行われます。 そのため、これはサーバー側のコードで実装されます。

OAuth フローの目的は、アプリケーションの ID と承認の安全を確保することです。 認証コード フローでは、*クライアント シークレット*が提供されます。これは、秘密にしておく必要があります。 SPA など、サーバー側のバックエンドがないアプリケーションには、秘密を保護する方法がありませんので、SPA では暗黙的フローを使用することをお勧めします。

暗黙的フローと認証コード フローのメリットとデメリットについて理解しておく必要があります。 これら 2 つのフローの詳細については、「[認証コード フロー](https://tools.ietf.org/html/rfc6749#section-1.3.1)」と「[暗黙的フロー](https://tools.ietf.org/html/rfc6749#section-1.3.2)」を参照してください。

> [!NOTE]
> You also have the option of using a middleman service to perform authorization and pass the access token to your add-in. For details about this scenario, see the **Middleman services** section later in this article.

## <a name="using-the-implicit-flow-in-office-add-ins"></a>Office アドインに暗黙的フローを使用する

オンライン サービスが暗黙的フローをサポートしているかどうかを判断する最良の方法は、サービスのドキュメントを調べることです。

暗黙的フローをサポートするライブラリの詳細については、後述の「**ライブラリ**」セクションを参照してください。

## <a name="using-the-authorization-code-flow-in-office-add-ins"></a>Office アドインに認証コード フローを使用する

Many libraries are available for implementing the Authorization Code flow in various languages and frameworks. For more information about some of these libraries, see the **Libraries** section later in this article.

## <a name="libraries"></a>ライブラリ

Libraries are available for many languages and platforms, for both the Implicit flow and the Authorization Code flow. Some libraries are general purpose, while others are for specific online services.

**Google**: Search [GitHub.com/Google](https://github.com/google) for "auth" or the name of your language. Most of the relevant repos are named `google-auth-library-[name of language]`.

**Facebook**:[Facebook for Developers](https://developers.facebook.com) で "library" または "sdk" を検索します。

**General OAuth 2.0**: A page of links to libraries for over a dozen languages is maintained by the IETF OAuth Working Group at: [OAuth Code](https://oauth.net/code/). Note that some of these libraries are for implementing an OAuth compliant service. The libraries of interest to you as a an add-in developer are called *client* libraries on this page because your web server is a client of the OAuth compliant service.

## <a name="middleman-services"></a>仲介者サービス

Your add-in can use a middleman service such as [OAuth.io](https://oauth.io) or [Auth0](https://auth0.com) to perform authorization. A middleman service may either provide access tokens for popular online services or simplify the process of enabling social login for your add-in, or both. With very little code, your add-in can use either client-side script or server-side code to connect to the middleman service and it will send your add-in any required tokens for the online service. All of the authorization implementation code is in the middleman service. 

アドインの認証/承認用 UI では、ダイアログ API を使用してログイン ページを開くようにしてください。 詳細については、「[認証フローでダイアログ API を使用する](dialog-api-in-office-add-ins.md#use-the-dialog-apis-in-an-authentication-flow)」を参照してください。 この方法で Office ダイアログを開くと、そのダイアログは、親ページのインスタンス (アドインの作業ウィンドウや FunctionFile など) とはまったく別の新しいブラウザーと JavaScript エンジンのインスタンスを持ちます。 文字列に変換できるトークンなどの情報は、`messageParent` という API を使用して親に戻されます。 そうすることで、親ページはトークンを使用してリソースへの権限のある呼び出しを実行できます。 こうしたアーキテクチャのため、仲介者サービスによって提供される API の使用方法には注意が必要になります。 多くの場合、このサービスはコードでトークンを取得し、そのトークンを使用して後続のリソースへの呼び出しを実行する、ある種のコンテキスト オブジェクトを作成する API セットを提供します。 通常、このサービスには、そのコンテキスト オブジェクトの最初の呼び出し*および*作成を実行する単一の API メソッドがあります。 このようなオブジェクトは、完全に文字列化することができないため、Office ダイアログから親ページに渡せません。 一般に、仲介者サービスは抽象度の低い第 2 の API セット (REST API など) を提供しています。 この第 2 のセットには、トークンを使用してリソースへの権限のあるアクセスを実行するために、サービスからトークンを取得する API と、そのトークンをサービスに渡す別の API があります。 この抽象度の低い API は、Office ダイアログでトークンを取得し、そのトークンを親ページに `messageParent` を使用して渡すために操作する必要があります。 

## <a name="what-is-cors"></a>CORS とは

CORS stands for [Cross Origin Resource Sharing](https://developer.mozilla.org/docs/Web/HTTP/Access_control_CORS). For information about how to use CORS inside add-ins, see [Addressing same-origin policy limitations in Office Add-ins](addressing-same-origin-policy-limitations.md).
