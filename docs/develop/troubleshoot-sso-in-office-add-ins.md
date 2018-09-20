---
title: シングル サインオン (SSO) のエラー メッセージのトラブルシューティング
description: ''
ms.date: 12/08/2017
ms.openlocfilehash: a0eb0839596bad0dfe45c2cbbc05c2c3d74eda24
ms.sourcegitcommit: 3da2038e827dc3f274d63a01dc1f34c98b04557e
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/19/2018
ms.locfileid: "24016319"
---
# <a name="troubleshoot-error-messages-for-single-sign-on-sso-preview"></a>シングル サインオン (SSO) のエラー メッセージのトラブルシューティング (プレビュー)

この記事では、Office アドインのシングル サインオン (SSO) に関する問題のトラブルシューティング方法と、SSO が有効なアドインによって特別な条件やエラーを確実に処理する方法について説明します。

> [!NOTE]
> 現在、シングル サインオン API は Word、Excel、Outlook、PowerPoint のプレビューでサポートされています。 シングル サインオン API の現在のサポート状態に関する詳細は、「IdentityAPI の要件セット」https://docs.microsoft.com/javascript/office/requirement-sets/identity-api-requirement-sets)をご覧ください。
> SSO を使用するには、アドインの HTML 起動ページの https://appsforoffice.microsoft.com/lib/beta/hosted/office.js からベータ版 Office の JavaScript ライブラリを読み込む必要があります。
> Outlook アドインで作業している場合は、Office 365 テナントの先進認証が有効になっていることを確認してください。 この方法の詳細については、「[Exchange Online: テナントの先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)」をご覧ください。

## <a name="debugging-tools"></a>デバッグ ツール

開発時は、アドインの Web サービスからの HTTP 要求および応答を傍受して表示することができるツールを使用することを強くお勧めします。最も一般的なものは、次の 2 つです。 

- [Fiddler](http://www.telerik.com/fiddler): 無料 ([ドキュメント](http://docs.telerik.com/fiddler/configure-fiddler/tasks/configurefiddler))
- [Charles](https://www.charlesproxy.com/): 30 日間無料 ([ドキュメント](https://www.charlesproxy.com/documentation/))

サービス API を開発する際には、次のツールを試してみることもできます。

- [Postman](http://www.getpostman.com/postman):無料 ([ドキュメント](https://www.getpostman.com/docs/))

## <a name="causes-and-handling-of-errors-from-getaccesstokenasync"></a>getAccessTokenAsync からのエラーの原因と処理

このセクションで説明するエラー処理の例については、次を参照してください。
- [Office-Add-in-ASPNET-SSO の Home.js](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js)
- [Office-Add-in-NodeJS-SSO の program.js](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js)

> [!NOTE]
> このセクションでの提案に加えて、Outlook アドインには、どのような 13*nnn* エラーにも応答する追加の機能があります。 詳細については、「[シナリオ: Outlook アドインでサービスにシングル サインオンを実装する](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in)」および「[AtachmentDemo サンプル アドイン](https://github.com/OfficeDev/outlook-add-in-attachments-demo)」をご覧ください。 

### <a name="13000"></a>13000

[getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) API は、このアドインまたは Office バージョンではサポートされていません。 

- このバージョンの Office は、SSO をサポートしていません。所要のバージョンは Office 2016 バージョン 1710、ビルド 8629.nnnn 以降 (「クイック実行」と呼ばれることもある Office 365 のサブスクリプション バージョン) です。このバージョンを入手するには、Office Insider への参加が必要になることがあります。詳細については、「[Office Insider](https://products.office.com/office-insider?tab=tab-1)」を参照してください。 
- アドインのマニフェストに適切な [WebApplicationInfo](https://docs.microsoft.com/javascript/office/manifest/webapplicationinfo?view=office-js) セクションがありません。

### <a name="13001"></a>13001

ユーザーは Office にサインインしていません。 コードでは、`getAccessTokenAsync` メソッドを再度呼び出して、[options](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) パラメーターでオプション `forceAddAccount: true` を渡す必要があります。 しかし、これを一回以上実行しないでください。 ユーザーがサインインしないと決定した可能性があります。

Office Online では、このエラーが発生することはありません。 ユーザーの Cookie が失効している場合、Office Online はエラー 13006 を返します。 

### <a name="13002"></a>13002

ユーザーはサインインまたは同意を中止しました。（コンセント ダイアログで **Cancel** を選択した場合など） 
- アドインがユーザーのサインイン (または同意) の必要がない機能を提供している場合、コードはこのエラーをキャッチし、アドインが継続して実行するようにしなければなりません。
- アドインで、同意を得たサインイン ユーザーが必要な場合、コードはユーザーに操作を繰り返すよう 1 度だけ要求する必要があります。 

### <a name="13003"></a>13003

ユーザーの種類がサポートされていません。 ユーザーは、有効な Microsoft アカウント、職場または学校のアカウントで Office にサインインしていません。 このエラーは、Office がオンプレミス ドメイン アカウントで実行されている場合に発生する可能性があります。 コードは、ユーザーに Office にサインインするよう要求する必要があります。

### <a name="13004"></a>13004

無効なリソースです。 アドイン マニフェストは、正しく構成されていません。 マニフェストを更新してください。 詳細は、「[マニフェストの問題を検証し、トラブルシューティングする](../testing/troubleshoot-manifest.md)」をご覧ください。 最も一般的な問題は、 **Resource** 要素（ **WebApplicationInfo** 要素の中にあるもの）に、アドインのドメインと一致しないドメインがあることです。 リソース値のプロトコルの部分は、"https" ではなく "api" である必要がありますが、ドメイン名 (もしあれば、ポートを含む) の他のすべての部分は、アドインの場合と同じにする必要があります。

### <a name="13005"></a>13005

無効な許可です。 このエラーは、通常、Office がアドインの Web サービスに対して事前に承認されていないことを意味します。 詳細については、「[サービス アプリケーションを作成する](sso-in-office-add-ins.md#create-the-service-application)」、および「[Azure AD V2.0 エンドポイントにアドインを登録する](create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v20-endpoint)」(ASP.NET) または「[Azure AD V2.0 エンドポイントにアドインを登録する](create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v20-endpoint)」(Node JS) を参照してください。 このエラーは、ユーザーが `profile` にサービス アプリケーションのアクセス許可を与えていない場合にも発生する可能性があります。

### <a name="13006"></a>13006

クライアント エラー。コードでは、ユーザーにサインアウトしてから Office を再起動するように指示するか、Office Online セッションを再開する必要があります。

### <a name="13007"></a>13007

Office ホストは、アドインの Web サービスへのアクセス トークンを取得できませんでした。
- 開発中にこのエラーが発生した場合は、アドインの登録とアドインマニフェストが、 `openid` および `profile` のアクセス許可を指定していることを確認してください。 詳細については、「[Azure AD V2.0 エンドポイントにアドインを登録する](create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v20-endpoint)」(ASP.NET) または「[Azure AD V2.0 エンドポイントにアドインを登録する](create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v20-endpoint)」(Node JS)、および「[アドインを構成する](create-sso-office-add-ins-aspnet.md#configure-the-add-in)」(ASP.NET) または「[アドインを構成する](create-sso-office-add-ins-nodejs.md#configure-the-add-in)」(Node JS) を参照してください。
- 運用環境では、このエラーの原因となることがいくつかあります。 以下はその中のいくつかの例です。
    - ユーザーは、事前に同意し、のちにその同意を取り消しました。 コードは `getAccessTokenAsync` オプション付きのメソッド `forceConsent: true` を再び呼び出す必要があります。しかし、1 度のみです。
    - ユーザーは、Microsoft Account (MSA) ID を持っています。 職場または学校のアカウントで、他の13nnnエラーの1つが発生する状況によっては、MSAを使用した場合に13007が発生することがあります。 

  これらすべてのケースについて、すでに `forceConsent` オプションを試行した場合には、ユーザーがのちに操作を再試行することをコードで示唆することができます。

### <a name="13008"></a>13008

ユーザーは、前回の `getAccessTokenAsync` の呼び出しが完了する前に `getAccessTokenAsync` を呼び出す操作をトリガーしました。 コードでは、ユーザーに前の操作が完了した後に操作を繰り返すよう、要求する必要があります。

### <a name="13009"></a>13009

アドインは、オプション `forceConsent: true` を指定して `getAccessTokenAsync` メソッドを呼び出しましたが、アドインのマニフェストが強制的な同意をサポートしていない種類のカタログに展開されています。 コードでは、`getAccessTokenAsync` メソッドを再度呼び出して、[options](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) パラメーターでオプション `forceConsent: false` を渡す必要があります。 ただし、`forceConsent: true` を指定した `getAccessTokenAsync` の呼び出し自体が、`forceConsent: false` を指定した `getAccessTokenAsync` の失敗した呼び出しに対する自動的な応答の可能性があるため、コードでは、`forceConsent: false` を指定した `getAccessTokenAsync` が既に呼び出されているかどうかを追跡記録する必要があります。 既に呼び出されていた場合、コードでは、ユーザーに Office からサインアウトして、再度サインインするように求める必要があります。

> [!NOTE]
> Microsoft は、必ずしもすべての種類のアドイン カタログに、この制約を課す予定はありません。 制約が課されなかった場合、このエラーが発生することはなくなります。

### <a name="13010"></a>13010

ユーザーが Office Online でアドインを実行していて、Edge または Internet Explorer を使用しています。 ユーザーの Office 365 ドメインと、login.microsoftonline.com ドメインは、ブラウザー設定で異なるセキュリティ ゾーンに含まれています。 このエラーが返された場合、ユーザーには、これについて説明するエラーとゾーンの構成を変更する方法に関するページへのリンクが表示されています。 アドインがユーザーのサインインを必要としない機能を提供している場合、コードでは、このエラーをキャッチして、アドインの実行を続行する必要があります。

### <a name="13012"></a>13012

アドインが、`getAccessTokenAsync` API をサポートしていないプラットフォームで実行されています。 たとえば、この APIは iPad 上ではサポートされていません。 「[ユーザー API の要件セット](https://docs.microsoft.com/javascript/office/requirement-sets/identity-api-requirement-sets)」をご覧ください。

### <a name="50001"></a>50001

このエラー（`getAccessTokenAsync` 特有のものではありません）は、ブラウザが office.js ファイルの古いコピーをキャッシュしたことを示している可能性があります。 ブラウザのキャッシュをクリアしてください。 もう 1 つの可能性は、Office のバージョンが SSO をサポートできる新しいものでないということです。 [Prerequisites](create-sso-office-add-ins-aspnet.md#prerequisites) を参照してください。

## <a name="errors-on-the-server-side-from-azure-active-directory"></a>Azure Active Directory からのサーバー側のエラー

このセクションで説明するエラー処理の例については、次を参照してください。
- [OfficeアドインASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO)
- [OfficeアドインNodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO)


### <a name="conditional-access--multifactor-authentication-errors"></a>条件付きアクセスおよび多要素認証のエラー
 
AAD および Office 365 の特定の ID 構成では、Microsoft Graph でアクセス可能な一部のリソースで、ユーザーの Office 365 のテナントでは必要ない場合でも、多要素認証 (MFA) が必要な場合があります。 AAD は、MFA で保護されたリソースへのトークンの要求を、代理フロー経由で受け取ると、アドインの Web サービスに `claims` プロパティを含む JSON メッセージを返します。 claims プロパティには、さらに必要となる認証要素の情報が含まれています。 

サーバー側のコードはこのメッセージをテストし、クライアント側のコードに claims 値を中継する必要があります。 Office が SSO アドインの認証を処理するため、この情報がクライアントで必要になります。クライアントへのメッセージは、エラー (`500 Server Error` や `401 Unauthorized` など) または成功応答の本文 (`200 OK` など) のいずれかになります。 どちらの場合でも、アドインの Web API に対する、コードによるクライアント側の AJAX 呼び出しのコールバック (失敗または成功) が、この応答をテストする必要があります。 claims 値が中継されている場合、コードは `getAccessTokenAsync` を再呼び出しして [options](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) パラメーターのオプション `authChallenge: CLAIMS-STRING-HERE` を渡す必要があります。 AAD がこの文字列を認識すると、ユーザーに追加の要素を入力するよう促してから、代理フローで受け入れられる新しいアクセス トークンを返します。

### <a name="consent-missing-errors"></a>同意なしエラー

AAD に、ユーザー (またはテナント管理者) がアドインに (Microsoft Graph リソースに対して) 同意した記録がない場合、AAD はエラー メッセージを Web サービスに送信します。 コードは、`forceConsent: true` オプションで `getAccessTokenAsync` を再呼び出しするよう、(`403 Forbidden` 応答の本文などで) クライアントに指示する必要があります。

### <a name="invalid-or-missing-scope-permission-errors"></a>無効または不足した範囲 (アクセス許可) のエラー

- サーバー側のコードでは、`403 Forbidden` 応答をクライアントに送って、ユーザーにわかりやすいメッセージを提示する必要があります。可能な場合は、そのエラーをコンソールに出力するか、ログに記録します。
- アドイン マニフェストの [Scopes](https://docs.microsoft.com/javascript/office/manifest/scopes?view=office-js) セクションで、必要なすべてのアクセス許可が指定されていることを確認してください。 また、アドインの Web サービスの登録で同じアクセス許可が指定されていることを確認してください。 スペルミスもチェックしてください。 詳細については、「[Azure AD V2.0 エンドポイントにアドインを登録する](create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v20-endpoint)」(ASP.NET) または「[Azure AD V2.0 エンドポイントにアドインを登録する](create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v20-endpoint)」(Node JS)、および「[アドインを構成する](create-sso-office-add-ins-aspnet.md#configure-the-add-in)」(ASP.NET) または「[アドインを構成する](create-sso-office-add-ins-nodejs.md#configure-the-add-in)」(Node JS) を参照してください。

### <a name="expired-or-invalid-token-errors-when-calling-microsoft-graph"></a>Microsoft Graph 呼び出し時の期限切れまたは無効なトークンのエラー

MSAL を含む一部の認証および承認ライブラリは、必要に応じてキャッシュされた更新トークンを使用することにより、期限切れトークンのエラーが発生しないようにします。 独自のトークン キャッシュ システムをコーディングすることもできます。 コーディングを行うサンプルについては、[Office Add-in NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO)、特にファイル [auth.ts](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/src/auth.ts) を参照してください。

しかし、期限切れトークンや無効なトークンのエラーが発生した場合、コードは `getAccessTokenAsync` を再呼び出しして、アドインの Web API のエンドポイントへの呼び出しを繰り返すよう、(`401 Unauthorized` 応答の本文などで) クライアントに指示しなければなりません。つまり、Microsoft Graph に対する新しいトークンを取得する代理フローを繰り返します。 

### <a name="invalid-token-error-when-calling-microsoft-graph"></a>Microsoft Graph 呼び出し時の無効なトークンのエラー

このエラーは、期限切れトークンのエラーと同様に処理します。前のセクションを参照してください。

### <a name="invalid-audience-error"></a>無効な対象ユーザーのエラー

サーバー側のコードは、`403 Forbidden` 応答をクライアントに送って、ユーザーにわかりやすいメッセージを提示しなければなりません。場合によっては、エラーについて、コンソールでログを作成するか、ログに記録する必要もあります。

トークン検証のためのマルチテナント サポートの追加の詳細については、[Azure マルチテナント サンプル](https://github.com/Azure-Samples/active-directory-dotnet-webapp-webapi-multitenant-openidconnect)に関する記事をご覧ください。
