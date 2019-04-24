---
title: シングル サインオン (SSO) のエラー メッセージのトラブルシューティング
description: ''
ms.date: 03/22/2019
localization_priority: Priority
ms.openlocfilehash: 1b885834304ebedd62eea206f02dae4bacefba5c
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449961"
---
# <a name="troubleshoot-error-messages-for-single-sign-on-sso-preview"></a>シングル サインオン (SSO) のエラー メッセージのトラブルシューティング (プレビュー)

この記事では、Office アドインのシングル サインオン (SSO) に関する問題のトラブルシューティング方法と、SSO が有効なアドインによって特別な条件やエラーを確実に処理する方法について説明します。

> [!NOTE]
> 現在、シングル サインオン API は Word、Excel、Outlook、PowerPoint のプレビューでサポートされています。 シングル サインオン API の現在のサポート状態に関する詳細は、「[IdentityAPI の要件セット](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets)」を参照してください。
> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]
> Outlook アドインで作業している場合は、Office 365 テナントの先進認証が有効になっていることを確認してください。 この方法の詳細については、「[Exchange Online: テナントの先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)」を参照してください。

## <a name="debugging-tools"></a>デバッグ ツール

開発時は、アドインの Web サービスからの HTTP 要求および応答を傍受して表示することができるツールを使用することを強くお勧めします。最も一般的なものは、次の 2 つです。

- [Fiddler](https://www.telerik.com/fiddler): 無料 ([ドキュメント](https://docs.telerik.com/fiddler/configure-fiddler/tasks/configurefiddler))
- [Charles](https://www.charlesproxy.com/): 30 日間無料  ([ドキュメント](https://www.charlesproxy.com/documentation/))。

サービス API を開発する際には、次のツールを試してみることもできます。

- [Postman](https://www.getpostman.com/postman):無料 ([ドキュメント](https://www.getpostman.com/docs/))

## <a name="causes-and-handling-of-errors-from-getaccesstokenasync"></a>getAccessTokenAsync からのエラーの原因と処理

このセクションで説明するエラー処理の例については、次を参照してください。
- [Office-Add-in-ASPNET-SSO の Home.js](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/Home.js)
- [Office-Add-in-NodeJS-SSO の program.js](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/public/program.js)

> [!NOTE]
> このセクションの提案以外にも、Outlook アドインには任意の 13*nnn* のエラーに対応するその他の方法があります。 詳細については、「[シナリオ: Outlook アドインでサービスにシングル サインオンを実装する](/outlook/add-ins/implement-sso-in-outlook-add-in)」と [AttachmentsDemo サンプル アドイン](https://github.com/OfficeDev/outlook-add-in-attachments-demo)を参照してください。

### <a name="13000"></a>13000

[getAccessTokenAsync](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) API は、このアドインまたは Office バージョンではサポートされていません。

- この Office のバージョンは、SSO をサポートしていません。 必要なバージョンは Office 365 (Office のサブスクリプション バージョン)、バージョン 1710、ビルド 8629.nnnn 以降。 このバージョンを入手するには、Office Insider への参加が必要になることがあります。 詳細については、「[Office Insider](https://products.office.com/office-insider?tab=tab-1)」を参照してください。
- アドインのマニフェストに適切な [WebApplicationInfo](/office/dev/add-ins/reference/manifest/webapplicationinfo) セクションがありません。

アドインがこのエラーに対応するには、ユーザー認証の代替システムにフォールバックする必要があります。 詳細については、「[要件とベスト プラクティス](/office/dev/add-ins/develop/sso-in-office-add-ins#requirements-and-best-practices)」を参照してください。

### <a name="13001"></a>13001

ユーザーは Office にサインインしていません。 コードでは、`getAccessTokenAsync` メソッドを再度呼び出して、[options](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) パラメーターでオプション `forceAddAccount: true` を渡す必要があります。 ただし、この操作は複数回実行しないでください。 ユーザーはサインインしないと判断する可能性があります。

Office Online では、このエラーが発生することはありません。 ユーザーの Cookie が失効している場合、Office Online はエラー 13006 を返します。

### <a name="13002"></a>13002

ユーザーが、同意ダイアログの **[キャンセル]** を選択するなどして、サインインまたは同意を中止しました。

- アドインがユーザーのサインイン (または同意) の必要がない機能を提供している場合、コードはこのエラーをキャッチし、アドインが継続して実行するようにしなければなりません。
- アドインで、同意を得たサインイン ユーザーが必要な場合、コードはユーザーに操作を繰り返すよう 1 度だけ要求する必要があります。

### <a name="13003"></a>13003

ユーザーの種類がサポートされていません。 ユーザーは、有効な Microsoft アカウントまたは Office 365 ("職場または学校") アカウントで Office にサインインしていません。 このエラーは、Office がオンプレミス ドメイン アカウントで実行されている場合に発生する可能性があります。 Office にサインインするか、代わりのユーザー認証システムにフォールバックするようにユーザーに求めるコードを作成します。 詳細については、「[要件とベスト プラクティス](/office/dev/add-ins/develop/sso-in-office-add-ins##requirements-and-best-practices)」を参照してください。

### <a name="13004"></a>13004

無効なリソースです。 アドイン マニフェストは、正しく構成されていません。 マニフェストを更新してください。 詳細については、「[マニフェストの問題を検証し、トラブルシューティングする](../testing/troubleshoot-manifest.md)」を参照してください。 最も一般的な問題は、**Resource** 要素 (**WebApplicationInfo** 要素内) にアドインのドメインと一致しないドメインがあることです。 Resource 値のプロトコル部分は "https" ではなく "api" である必要があります。ドメイン名の他のすべての部分は (ポートがある場合はそれも含めて)、アドインと同じである必要があります。

### <a name="13005"></a>13005

無効な許可です。 このエラーは、通常、Office がアドインの Web サービスに対して事前に承認されていないことを意味します。 詳細については、「[サービス アプリケーションを作成する](sso-in-office-add-ins.md#create-the-service-application)」、および「[Azure AD V2.0 エンドポイントにアドインを登録する](create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v20-endpoint)」(ASP.NET) または「[Azure AD V2.0 エンドポイントにアドインを登録する](create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v20-endpoint)」(Node JS) を参照してください。 このエラーは、ユーザーが自分の `profile` に対するアクセス許可をサービス アプリケーションに与えていない場合にも発生する可能性があります。

### <a name="13006"></a>13006

クライアント エラー。コードでは、ユーザーにサインアウトしてから Office を再起動するように指示するか、Office Online セッションを再開する必要があります。

### <a name="13007"></a>13007

Office ホストは、アドインの Web サービスへのアクセス トークンを取得できませんでした。

- 開発中にこのエラーが発生する場合は、アドインの登録とアドインのマニフェストで `openid` および `profile` のアクセス許可が指定されていることを確認してください。 詳細については、「[Azure AD V2.0 エンドポイントにアドインを登録する](create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v20-endpoint)」(ASP.NET) または「[Azure AD V2.0 エンドポイントにアドインを登録する](create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v20-endpoint)」(Node JS)、および「[アドインを構成する](create-sso-office-add-ins-aspnet.md#configure-the-add-in)」(ASP.NET) または「[アドインを構成する](create-sso-office-add-ins-nodejs.md#configure-the-add-in)」(Node JS) を参照してください。
- 運用環境では、このエラーの原因として考えられることがいくつかあります。 その一部を次に示します。
    - ユーザーが、以前は同意していた内容を取り消しました。 コードでオプション `forceConsent: true` を指定して `getAccessTokenAsync` メソッドをもう一度呼び出す必要があるのに、複数回の呼び出を行いました。
    - ユーザーは Microsoft アカウント (MSA) ID を使用しています。 職場または学校アカウントで他の 13nnn エラーのいずれかが発生する状況の場合、MSA を使用すると 13007 が発生することがあります。

  以上のいずれの場合でも、`forceConsent` オプションを既に一度試行している場合は、操作を後でやり直すようにユーザーに提案するコードを作成することをお勧めします。

### <a name="13008"></a>13008

ユーザーは、前回の `getAccessTokenAsync` の呼び出しが完了する前に `getAccessTokenAsync` を呼び出す操作をトリガーしました。 コードでは、ユーザーに前の操作が完了した後に操作を繰り返すよう、要求する必要があります。

### <a name="13009"></a>13009

アドインは、オプション `forceConsent: true` を指定して `getAccessTokenAsync` メソッドを呼び出しましたが、アドインのマニフェストが強制的な同意をサポートしていない種類のカタログに展開されています。 コードでは、`getAccessTokenAsync` メソッドを再度呼び出して、[options](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) パラメーターでオプション `forceConsent: false` を渡す必要があります。 ただし、`forceConsent: true` を指定した `getAccessTokenAsync` の呼び出し自体が、`forceConsent: false` を指定した `getAccessTokenAsync` の失敗した呼び出しに対する自動的な応答の可能性があるため、コードでは、`forceConsent: false` を指定した `getAccessTokenAsync` が既に呼び出されているかどうかを追跡記録する必要があります。 この場合、Office からサインアウトしてサインインし直すか、代わりのユーザー認証システムにフォールバックするようにユーザーに指示するコードを作成します。 詳細については、「[要件とベスト プラクティス](/office/dev/add-ins/develop/sso-in-office-add-ins#requirements-and-best-practices)」を参照してください。

> [!NOTE]
> Microsoft は、必ずしもすべての種類のアドイン カタログに、この制約を課す予定はありません。 制約が課されていない場合は、このエラーが発生することはありません。

### <a name="13010"></a>13010

ユーザーが Office Online でアドインを実行していて、Edge または Internet Explorer を使用しています。 ユーザーの Office 365 ドメインと、login.microsoftonline.com ドメインは、ブラウザー設定で異なるセキュリティ ゾーンに含まれています。 このエラーが返された場合、ユーザーには、これについて説明するエラーとゾーンの構成を変更する方法に関するページへのリンクが表示されています。 アドインがユーザーのサインインを必要としない機能を提供している場合、コードでは、このエラーをキャッチして、アドインの実行を続行する必要があります。

### <a name="13012"></a>13012

アドインは、`getAccessTokenAsync` API をサポートしていないプラットフォーム上で実行されています。 たとえば、iPad 上ではサポートされていません。 「[Identity API の要件セット](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets)」も参照してください。

### <a name="50001"></a>50001

このエラー (`getAccessTokenAsync` に固有ではありません) は、ブラウザーが office.js ファイルの古いコピーをキャッシュしていることを示す可能性があります。 開発中の場合は、ブラウザーのキャッシュをクリアしてください。 また、Office のバージョンが古いため、SSO をサポートしていない可能性も考えられます。 「[前提条件](create-sso-office-add-ins-aspnet.md#prerequisites)」を参照してください。

運用環境のアドインの場合、アドインがこのエラーに対応するには、ユーザー認証の代替システムにフォールバックする必要があります。 詳細については、「[要件とベスト プラクティス](/office/dev/add-ins/develop/sso-in-office-add-ins##requirements-and-best-practices)」を参照してください。


## <a name="errors-on-the-server-side-from-azure-active-directory"></a>Azure Active Directory からのサーバー側のエラー

このセクションで説明するエラー処理の例については、次を参照してください。
- [Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO)
- [Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO)


### <a name="conditional-access--multifactor-authentication-errors"></a>条件付きアクセスおよび多要素認証のエラー

AAD および Office 365 の特定の ID 構成では、Microsoft Graph でアクセス可能な一部のリソースで、ユーザーの Office 365 のテナントでは必要ない場合でも、多要素認証 (MFA) が必要な場合があります。 AAD は、MFA で保護されたリソースへのトークンの要求を、代理フロー経由で受け取ると、アドインの Web サービスに `claims` プロパティを含む JSON メッセージを返します。 claims プロパティには、さらに必要となる認証要素の情報が含まれています。

サーバー側のコードはこのメッセージをテストし、クライアント側のコードに claims 値を中継する必要があります。 Office が SSO アドインの認証を処理するため、この情報がクライアントで必要になります。クライアントへのメッセージは、エラー (`500 Server Error` や `401 Unauthorized` など) または成功応答の本文 (`200 OK` など) のいずれかになります。 どちらの場合でも、アドインの Web API に対する、コードによるクライアント側の AJAX 呼び出しのコールバック (失敗または成功) が、この応答をテストする必要があります。 claims 値が中継されている場合、コードは `getAccessTokenAsync` を再呼び出しして `authChallenge: CLAIMS-STRING-HERE` パラメーターのオプション [](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) を渡す必要があります。 AAD がこの文字列を認識すると、ユーザーに追加の要素を入力するよう促してから、代理フローで受け入れられる新しいアクセス トークンを返します。

### <a name="consent-missing-errors"></a>同意なしエラー

AAD に、ユーザー (またはテナント管理者) がアドインに (Microsoft Graph リソースに対して) 同意した記録がない場合、AAD はエラー メッセージを Web サービスに送信します。 コードは、`403 Forbidden` オプションで `getAccessTokenAsync` を再呼び出しするよう、(`forceConsent: true` 応答の本文などで) クライアントに指示する必要があります。

### <a name="invalid-or-missing-scope-permission-errors"></a>無効または不足した範囲 (アクセス許可) のエラー

- サーバー側のコードでは、`403 Forbidden` 応答をクライアントに送って、ユーザーにわかりやすいメッセージを提示する必要があります。可能な場合は、そのエラーをコンソールに出力するか、ログに記録します。
- アドイン マニフェストの[範囲](/office/dev/add-ins/reference/manifest/scopes)セクションで、必要なすべてのアクセス許可が指定されていることを確認してください。 また、アドインの Web サービスの登録で同じアクセス許可が指定されていることを確認してください。 スペルミスもチェックしてください。 詳細については、「[Azure AD V2.0 エンドポイントにアドインを登録する](create-sso-office-add-ins-aspnet.md#register-the-add-in-with-azure-ad-v20-endpoint)」(ASP.NET) または「[Azure AD V2.0 エンドポイントにアドインを登録する](create-sso-office-add-ins-nodejs.md#register-the-add-in-with-azure-ad-v20-endpoint)」(Node JS)、および「[アドインを構成する](create-sso-office-add-ins-aspnet.md#configure-the-add-in)」(ASP.NET) または「[アドインを構成する](create-sso-office-add-ins-nodejs.md#configure-the-add-in)」(Node JS) を参照してください。

### <a name="expired-or-invalid-token-errors-when-calling-microsoft-graph"></a>Microsoft Graph 呼び出し時の期限切れまたは無効なトークンのエラー

MSAL を含む一部の認証および承認ライブラリは、必要に応じてキャッシュされた更新トークンを使用することにより、期限切れトークンのエラーが発生しないようにします。 独自のトークン キャッシュ システムをコーディングすることもできます。 コーディングを行うサンプルについては、[Office Add-in NodeJS SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO)、特にファイル [auth.ts](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Completed/src/auth.ts) を参照してください。

しかし、期限切れトークンや無効なトークンのエラーが発生した場合、コードは `getAccessTokenAsync` を再呼び出しして、アドインの Web API のエンドポイントへの呼び出しを繰り返すよう、(`401 Unauthorized` 応答の本文などで) クライアントに指示しなければなりません。つまり、Microsoft Graph に対する新しいトークンを取得する代理フローを繰り返します。

### <a name="invalid-token-error-when-calling-microsoft-graph"></a>Microsoft Graph 呼び出し時の無効なトークンのエラー

このエラーは、期限切れトークンのエラーと同様に処理します。前のセクションを参照してください。

### <a name="invalid-audience-error"></a>無効な対象ユーザーのエラー

サーバー側のコードは、`403 Forbidden` 応答をクライアントに送って、ユーザーにわかりやすいメッセージを提示しなければなりません。場合によっては、エラーについて、コンソールでログを作成するか、ログに記録する必要もあります。

トークン検証のためのマルチテナント サポートの追加の詳細については、[Azure マルチテナント サンプル](https://github.com/Azure-Samples/active-directory-dotnet-webapp-webapi-multitenant-openidconnect)に関する記事をご覧ください。
