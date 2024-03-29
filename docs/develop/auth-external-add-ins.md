---
title: Microsoft 以外の ID プロバイダーでの承認
description: OAuth 2.0 と承認コードと暗黙的フローを使用して、Microsoft 以外のデータ ソースへの承認を取得します。
ms.date: 01/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 873bf0ad86490670db7a4733db971e377748babf
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743641"
---
# <a name="authorization-with-non-microsoft-identity-providers"></a>Microsoft 以外の ID プロバイダーでの承認

アドインで使用できるサービスには、Microsoft ID プラットフォームに加えて、多くの一般的な ID 提供サービスがあります。 ユーザーは、ユーザーやアプリケーション (Officeアドインなど) に、他のアプリケーションのユーザーのアカウントにアクセスできます。

Web アプリケーションからオンライン サービスへのアクセスを可能にするための業界標準のフレームワークは **OAuth 2.0** です。ほとんどの場合、このフレームワークをアドインで使用するために、その動作のしくみを詳しく知る必要はありません。開発者は、この詳細を簡略化している多数のライブラリを使用できます。

OAuth の基本的な考え方は、ユーザーやグループと同様に、アプリケーションは専用の ID とアクセス許可のセットによって、それ自体が[セキュリティ プリンシパル](/windows/security/identity-protection/access-control/security-principals)になり得るということです。 通常のシナリオでは、ユーザーがオンライン サービスを必要とする Office アドインのアクションを実行すると、アドインは、ユーザーのアカウントへ特定のセットのアクセス許可を付与するよう求める要求をサービスに送信します。 サービスは、該当するアクセス許可をアドインに付与するように求めるプロンプトをユーザーに表示します。 アクセス許可が付与されると、サービスは小さなエンコードされた *アクセス トークン* をアドインに送信します。 アドインは、サービスの API へのすべての要求にトークンを含めることで、サービスを使用できるようになります。 ただし、そのアドインが実行できるアクションは、ユーザーが付与したアクセス許可の範囲内に限定されます。 また、トークンは特定の時間が経過すると期限切れになります。

さまざまなシナリオに向けて、*フロー* または *許可の種類* と呼ばれる、いくつかの OAuth パターンが設計されています。 次の 2 つのパターンが最も一般的に実装されています。

- **暗黙的フロー**: アドインとオンライン サービスとの通信は、クライアント側の JavaScript で実装されます。 このフローは、シングル ページ アプリケーション (SPA) で一般的に使用されます。
- **認証コード フロー**: アドインの Web アプリケーションとオンライン サービスとの通信は、*サーバー間* で行われます。そのため、これはサーバー側のコードで実装されます。

OAuth フローの目的は、アプリケーションの ID と承認の安全を確保することです。 認証コード フローでは、*クライアント シークレット* が提供されます。これは、秘密にしておく必要があります。 SPA など、サーバー側のバックエンドがないアプリケーションには、秘密を保護する方法がありませんので、SPA では暗黙的フローを使用することをお勧めします。

暗黙的フローと認証コード フローのメリットとデメリットについて理解しておく必要があります。 これら 2 つのフローの詳細については、「[認証コード フロー](https://tools.ietf.org/html/rfc6749#section-1.3.1)」と「[暗黙的フロー](https://tools.ietf.org/html/rfc6749#section-1.3.2)」を参照してください。

> [!NOTE]
> 仲介者サービスを使用するというオプションもあります。このサービスは、自動的に承認を行い、アドインにアクセス トークンを渡します。 このシナリオの詳細については、後述の「**仲介者サービス**」セクションを参照してください。

## <a name="use-the-implicit-flow-in-office-add-ins"></a>アドインで暗黙的フロー Office使用する

オンライン サービスが暗黙的フローをサポートしているかどうかを判断する最良の方法は、サービスのドキュメントを調べることです。

暗黙的フローをサポートするライブラリの詳細については、後述の「**ライブラリ**」セクションを参照してください。

## <a name="use-the-authorization-code-flow-in-office-add-ins"></a>アドインの認証コード フロー Office使用する

各種の言語とフレームワークで認証コード フローを実装するために利用できるライブラリは多数あります。これらのライブラリの詳細については、後述の「**ライブラリ**」セクションを参照してください。

## <a name="libraries"></a>ライブラリ

各種の言語とプラットフォームで暗黙的フローと認証コード フローを実装するために利用できるライブラリが多数あります。 ライブラリには汎用のものや、特定のオンライン サービス向けのものがあります。

**Facebook**:[Facebook for Developers](https://developers.facebook.com) で "library" または "sdk" を検索します。

**汎用の OAuth 2.0**:数十の言語に対応したライブラリへのリンクが、「[OAuth Code](https://oauth.net/code/)」のページに掲載されています。このページは、IETF OAuth 作業部会によって維持されています。これらのライブラリの一部は、OAuth 準拠のサービスを実装するためのものです。アドイン開発者にとって重要なライブラリは、このページに記載された *クライアント* と呼ばれるライブラリです。これは、目的の Web サーバーが OAuth 準拠のサービスのクライアントになるためです。

## <a name="middleman-services"></a>仲介者サービス

アドインでは、[OAuth.io](https://oauth.io) や [Auth0](https://auth0.com) などの仲介者サービスを使用して、承認を実行できます。このサービスは、大手オンライン サービスに対応したアクセス トークンを提供するものか、アドインでソーシャル ログインできるようにするプロセスを簡単にするもの (またはその両方) です。短いコードを使用することで、仲介者サービスに接続するクライアント側スクリプトやサーバー側コードをアドインで使用できるようになり、仲介者サービスがオンライン サービスに必要なトークンをアドインに送信します。すべての承認の実装コードは、仲介者サービスに含まれています。

アドインの認証/承認用 UI では、ダイアログ API を使用してログイン ページを開くようにしてください。 詳細については、「[認証フローでダイアログ API を使用する](dialog-api-in-office-add-ins.md#use-the-dialog-apis-in-an-authentication-flow)」を参照してください。 この方法で Office ダイアログを開くと、そのダイアログは、親ページのインスタンス (アドインの作業ウィンドウや FunctionFile など) とはまったく別の新しいブラウザーと JavaScript エンジンのインスタンスを持ちます。 文字列に変換できるトークンなどの情報は、`messageParent` という API を使用して親に戻されます。 そうすることで、親ページはトークンを使用してリソースへの権限のある呼び出しを実行できます。 こうしたアーキテクチャのため、仲介者サービスによって提供される API の使用方法には注意が必要になります。 多くの場合、このサービスはコードでトークンを取得し、そのトークンを使用して後続のリソースへの呼び出しを実行する、ある種のコンテキスト オブジェクトを作成する API セットを提供します。 通常、このサービスには、そのコンテキスト オブジェクトの最初の呼び出し *および* 作成を実行する単一の API メソッドがあります。 このようなオブジェクトは、完全に文字列化することができないため、Office ダイアログから親ページに渡せません。 一般に、仲介者サービスは抽象度の低い第 2 の API セット (REST API など) を提供しています。 この第 2 のセットには、トークンを使用してリソースへの権限のあるアクセスを実行するために、サービスからトークンを取得する API と、そのトークンをサービスに渡す別の API があります。 この抽象度の低い API は、Office ダイアログでトークンを取得し、そのトークンを親ページに `messageParent` を使用して渡すために操作する必要があります。

## <a name="what-is-cors"></a>CORS とは

CORS は [Cross Origin Resource Sharing](https://developer.mozilla.org/docs/Web/HTTP/Access_control_CORS) の略です。アドイン内で CORS を使用する方法の詳細については、「[Office アドインにおける同一生成元ポリシーの制限への対処](addressing-same-origin-policy-limitations.md)」を参照してください。

## <a name="see-also"></a>関連項目

- [アドインでの認証と承認Office概要](overview-authn-authz.md)。