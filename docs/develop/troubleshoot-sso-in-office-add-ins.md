---
title: シングル サインオン (SSO) のエラー メッセージのトラブルシューティング
description: ''
ms.date: 12/08/2017
ms.openlocfilehash: 5abf10d8281ea54be9a172c3f45b742fb33991df
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/12/2018
ms.locfileid: "25506071"
---
# <a name="troubleshoot-error-messages-for-single-sign-on-sso-preview"></a>シングル サインオン (SSO) のエラー メッセージのトラブルシューティング (プレビュー)

この記事では、Office アドインのシングル サインオン (SSO) に関する問題のトラブルシューティング方法と、SSO が有効なアドインによって特別な条件やエラーを確実に処理する方法について説明します。

> [!NOTE]
> Word、Excel、Outlook、および PowerPoint のプレビューでは、単一のサインオンの API はサポートされて現在。シングル サインオンの API がサポートされている現在の詳細については、「IdentityAPI 要件の設定」https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets) を参照してください。SSO を使用するには、アドインのスタートアップの HTML ページのhttps://appsforoffice.microsoft.com/lib/beta/hosted/office.jsからベータ版の Office の JavaScript ライブラリを読み込む必要があります。Outlook のアドインで作業している場合は、Office 365 テナントの先進認証を必ず有効にします。詳細な方法については、「[Exchange Online: テナントで先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx) 」を参照してください。

## <a name="debugging-tools"></a>デバッグ ツール

開発時は、アドインの Web サービスからの HTTP 要求および応答を傍受して表示することができるツールを使用することを強くお勧めします。最も一般的なものは、次の 2 つです。 

- [Fiddler](http://www.telerik.com/fiddler): 無料 ([ドキュメント](http://docs.telerik.com/fiddler/configure-fiddler/tasks/configurefiddler))
- [Charles](https://www.charlesproxy.com/): 30 日間無料。([ドキュメント](https://www.charlesproxy.com/documentation/))

サービス API を開発する際には、次のツールを試してみることもできます。

- [Postman](http://www.getpostman.com/postman): 無料 ([ドキュメント](https://www.getpostman.com/docs/))

## <a name="causes-and-handling-of-errors-from-getaccesstokenasync"></a>getAccessTokenAsync からのエラーの原因と処理

このセクションで説明するエラー処理の例については、次を参照してください。
- [Office-Add-in-ASPNET-SSO の Home.js](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js)
- [Office-Add-in-NodeJS-SSO の program.js](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js)

> [!NOTE]
> ここで行った提案だけでなく Outlook のアドインを 13*nnn* エラーに応答する別の方法には。詳細については、「 [シナリオ: Outlook のアドインで、サービスへのシングル サインオンを実装](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in) と [AttachmentsDemo のサンプル アドイン](https://github.com/OfficeDev/outlook-add-in-attachments-demo)です。 

### <a name="13000"></a>13000

[getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) API は、このアドインまたは Office バージョンではサポートされていません。 

- このバージョンの Office は、SSO をサポートしていません。所要のバージョンは Office 2016 バージョン 1710、ビルド 8629.nnnn 以降 (「クイック実行」と呼ばれることもある Office 365 のサブスクリプション バージョン) です。このバージョンを入手するには、Office Insider への参加が必要になることがあります。詳細については、「[Office Insider](https://products.office.com/office-insider?tab=tab-1)」を参照してください。 
- アドインのマニフェストに適切な [WebApplicationInfo](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/webapplicationinfo?view=office-js) セクションがありません。

アドインは、このエラーに対処するため、ユーザー認証の代替システムに戻る必要があります。詳細については、 [要件とベスト プラクティス](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#requirements-and-best-practices)を参照してください。

### <a name="13001"></a>13001

ユーザーが Office にサインインしていません。コードでは`getAccessTokenAsync` メソッドを呼び出し、[オプション](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) パラメータのオプション`forceAddAccount: true` を渡す必要がありますが、1 回のみ行ってください。ユーザーがサインインしなかった可能性があります。

ユーザーの Cookie が失効している場合、Office Online はエラー 13006 を返します。 

### <a name="13002"></a>13002

ユーザーはサインインまたは同意を中止しました。（コンセント ダイアログで **Cancel** を選択した場合など） 

- アドインがユーザーのサインイン (または同意) の必要がない機能を提供している場合、コードはこのエラーをキャッチし、アドインが継続して実行するようにしなければなりません。
- アドインで、同意を得たサインイン ユーザーが必要な場合、コードはユーザーに操作を繰り返すよう 1 度だけ要求する必要があります。 

### <a name="13003"></a>13003

ユーザーの種類がサポートされていません。ユーザーが、有効な Microsoft アカウントまたは Office 365 (「職場または学校」) アカウントで Office にサインインしていません。これは、オンプレミス ドメイン アカウントを使用して Office を実行した場合などで起こります。コードでは、ユーザーに Office にサインインするよう要求するか、ユーザー認証の代替システムに戻る必要があります。詳細については、 [要件とベスト プラクティス](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins##requirements-and-best-practices)を参照してください。


### <a name="13004"></a>13004

無効なリソースです。アドインのマニフェストが正しく構成されていません。マニフェストを更新してください。詳細については、[マニフェストに関する問題の検証とトラブルシューティング](../testing/troubleshoot-manifest.md)を参照してください。最も一般的な問題は、 ** Resource** 要素（ ** WebApplicationInfo** 要素の中にあるもの）に、アドインのドメインと一致しないドメインがあることです。リソース値のプロトコルの部分は、"https" ではなく "api" である必要がありますが、ドメイン名 (もしあれば、ポートを含む) の他のすべての部分は、アドインの場合と同じにする必要があります。

### <a name="13005"></a>13005

無効な付与です。通常、Office がアドインの Web サービスに事前承認されなかったことを表します。詳細については、「[サービス アプリケーションを作成する](sso-in-office-add-ins.md#create-the-service-application)」および「[Azure AD V2.0 エンドポイントにアドインを登録する](create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v20-endpoint)」(ASP.NET) または「[Azure AD V2.0 エンドポイントにアドインを登録する](create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v20-endpoint)」(Node JS) を参照してください。また、ユーザーがサービス アプリケーションのアクセス許可を付与しなかった可能性もあります`profile`。

### <a name="13006"></a>13006

クライアント エラー。コードでは、ユーザーにサインアウトしてから Office を再起動するように指示するか、Office Online セッションを再開する必要があります。

### <a name="13007"></a>13007

Office ホストは、アドインの Web サービスへのアクセス トークンを取得できませんでした。

- 開発時にこのエラーが発生した場合は、アドイン登録とアドイン マニフェストが`openid` と `profile` のアクセス許可を指定していることを確認してください。詳細については、「[Azure AD V2.0 エンドポイントにアドインを登録する](create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v20-endpoint)」(ASP.NET) または「[Azure AD V2.0 エンドポイントにアドインを登録する](create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v20-endpoint)」(Node JS)、および「[アドインを構成する](create-sso-office-add-ins-aspnet.md#configure-the-add-in)」(ASP.NET) または「[アドインを構成する](create-sso-office-add-ins-nodejs.md#configure-the-add-in)」(Node JS) を参照してください。
- 運用環境では、このエラーの原因となることがいくつかあります。
    - ユーザーが以前に承諾した同意を無効にしました。コードは `getAccessTokenAsync` オプション付きのメソッド `forceConsent: true` を再び呼び出す必要があります。しかし、1 度のみです。
    - ユーザーは、Microsoft アカウント (MSA) の ID を持っています。職場または学校のアカウントで、他の13nnnエラーの1つが発生する状況によっては、MSAを使用した場合に13007が発生することがあります。 

  これらすべてのケースについて、すでに `forceConsent` オプションを試行した場合には、ユーザーがのちに操作を再試行することをコードで示唆することができます。

### <a name="13008"></a>13008

ユーザーは、前回の `getAccessTokenAsync` の呼び出しが完了する前に `getAccessTokenAsync` を呼び出す操作をトリガーしました。コードでは、ユーザーに前の操作が完了した後に操作を繰り返すよう、要求する必要があります。

### <a name="13009"></a>13009

アドインは  `forceConsent: true` オプションを使用して`getAccessTokenAsync` メソッド を呼び出しましたが、アドインのマニフェストが強制の承認をサポートしていないカタログの種類に展開されています。コードは`getAccessTokenAsync` メソッドを取り消して[オプション](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) パラメータのオプション`forceConsent: false` を渡す必要があります。ただし、`forceConsent: true`を指定した  `getAccessTokenAsync` の呼び出し自体が、`forceConsent: false`  を指定した `getAccessTokenAsync`  の失敗した呼び出しに対する自動的な応答の可能性があるため、コードでは、`forceConsent: false`  を指定した `getAccessTokenAsync`  が既に呼び出されているかどうかを追跡記録する必要があります。その場合、コードは、Office をサインアウトし、もう一度サインインするユーザーを知らせるか、または別のシステムのユーザー認証にフォールバックする必要があります。詳細については、 [要件とベスト プラクティス](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#requirements-and-best-practices)を参照してください。

> [!NOTE]
> Microsoft は、必ずしもすべての種類のアドイン カタログに、この制約を課す予定はありません。その場合、このエラーは表示されません。

### <a name="13010"></a>13010

ユーザーは、Office Online でアドインを実行し、Edge または Internet Explorer を使用しています。ユーザーの Office 365 ドメイン、および login.microsoftonline.com ドメインは、ブラウザーの設定で異なるセキュリティ ゾーンにあります。このエラーが返された場合、ユーザーはこれを説明し、ゾーン構成を変更する方法のページへリンクするエラーをすでに認識しています。アドインがユーザーのサインインを必要としない機能を提供している場合、コードでは、このエラーをキャッチして、アドインの実行を続行する必要があります。

### <a name="13012"></a>13012

アドインが、`getAccessTokenAsync` API をサポートしていないプラットフォームで実行されています。たとえば、iPad 上でサポートされていません。 [ユーザー API の要件セット](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets)を参照してください。

### <a name="50001"></a>50001

このエラー（`getAccessTokenAsync` 特有のものではありません）は、ブラウザが office.js ファイルの古いコピーをキャッシュしたことを示している可能性があります。開発するとき、ブラウザーのキャッシュをオフにします。別の可能性としては、Office のバージョンが最新ではないため SSO をサポートしていない場合です。 [前提条件](create-sso-office-add-ins-aspnet.md#prerequisites)を参照してください。

アドインは、このエラーに対処するため、ユーザー認証の代替システムに戻る必要があります。詳細については、 [要件とベスト プラクティス](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins##requirements-and-best-practices)を参照してください。


## <a name="errors-on-the-server-side-from-azure-active-directory"></a>Azure Active Directory からのサーバー側のエラー

このセクションで説明するエラー処理の例については、次を参照してください。
- [OfficeアドインASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO)
- [OfficeアドインNodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO)


### <a name="conditional-access--multifactor-authentication-errors"></a>条件付きアクセスおよび多要素認証のエラー
 
AAD および Office 365 の特定の ID 構成では、Microsoft Graph でアクセス可能な一部のリソースで、ユーザーの Office 365 のテナントでは必要ない場合でも、多要素認証 (MFA) が必要な場合があります。AAD がフローの代理でMFA で保護されたリソースへのトークンの要求を受信した場合、アドインの Web サービスに、`claims` プロパティを含む JSON メッセージを返します。要求プロパティには、さらにどのような認証要素が必要かについての情報があります。 

サーバー側コードは、このメッセージをテストし、クライアント側コードにクレームの値を中継する必要があります。Office が SSO アドインの認証を処理するため、この情報がクライアントで必要になります。クライアントへのメッセージは、エラー (`500 Server Error` や `401 Unauthorized` など) または成功応答の本文 (`200 OK` など) のいずれかになります。どちらの場合も、アドインの Web API へのコードのクライアント側の AJAX の (失敗または成功の) コールバックが、この応答のテストをする必要があります。クレームの値が中継されると、コードは`getAccessTokenAsync`を呼び出し、[オプション](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)パラメータのオプション`authChallenge: CLAIMS-STRING-HERE`を渡す必要があります。AAD は、この文字列を見て追加の要素をユーザーに促し、代理でのフローでは受け入れられる新しいアクセス トークンを返します。

### <a name="consent-missing-errors"></a>同意なしエラー

AAD に、ユーザー (またはテナント管理者) がアドインに (Microsoft Graph リソースに対して) 同意した記録がない場合、AAD はエラー メッセージを Web サービスに送信します。コードは、`forceConsent: true` オプションで `getAccessTokenAsync` を再呼び出しするよう、(`403 Forbidden` 応答の本文などで) クライアントに指示する必要があります。

### <a name="invalid-or-missing-scope-permission-errors"></a>無効または不足したスコープ (アクセス許可) のエラー

- サーバー側のコードでは、`403 Forbidden` 応答をクライアントに送って、ユーザーにわかりやすいメッセージを提示する必要があります。可能な場合は、そのエラーをコンソールに出力するか、ログに記録します。
- アドイン マニフェスト[スコープ](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/scopes?view=office-js) セクションが、すべての必要なアクセス許可を指定していることを確認します。同じアクセス許可を指定するアドイン Web サービスを登録してください。スペル チェックもできます。詳細については、「[ Azure AD V2.0 エンドポイントにアドインを登録する](create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v20-endpoint)」(ASP.NET) または「[ Azure AD V2.0 エンドポイントにアドインを登録する](create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v20-endpoint)」(Node JS)、および「[ アドインを構成する](create-sso-office-add-ins-aspnet.md#configure-the-add-in)」(ASP.NET) または「[ アドインを構成する](create-sso-office-add-ins-nodejs.md#configure-the-add-in)」(Node JS) を参照してください。

### <a name="expired-or-invalid-token-errors-when-calling-microsoft-graph"></a>Microsoft Graph 呼び出し時の期限切れまたは無効なトークンのエラー

MSAL を含む一部の認証および承認ライブラリは、必要に応じてキャッシュされた更新トークンを使用することにより、期限切れトークンのエラーが発生しないようにします。独自のトークン キャッシュ システムをコーディングすることもできます。コーディングを行うサンプルについては、[Office アドイン NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO)、特にファイル [auth.ts](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/src/auth.ts) を参照してください。

しかし、期限切れトークンや無効なトークンのエラーが発生した場合、コードは `getAccessTokenAsync` を再呼び出しして、アドインの Web API のエンドポイントへの呼び出しを繰り返すよう、(`401 Unauthorized` 応答の本文などで) クライアントに指示しなければなりません。つまり、Microsoft Graph に対する新しいトークンを取得する代理フローを繰り返します。 

### <a name="invalid-token-error-when-calling-microsoft-graph"></a>Microsoft Graph 呼び出し時の無効なトークンのエラー

このエラーは、期限切れトークンのエラーと同様に処理します。前のセクションを参照してください。

### <a name="invalid-audience-error"></a>無効な対象ユーザーのエラー

サーバー側のコードは、`403 Forbidden` 応答をクライアントに送って、ユーザーにわかりやすいメッセージを提示しなければなりません。場合によっては、エラーについて、コンソールでログを作成するか、ログに記録する必要もあります。

トークン検証のためのマルチテナント サポートの追加の詳細については、[Azure マルチテナント サンプル](https://github.com/Azure-Samples/active-directory-dotnet-webapp-webapi-multitenant-openidconnect)に関する記事をご覧ください。
