---
title: シングル サインオン (SSO) のエラー メッセージのトラブルシューティング
description: シングル サインオン (SSO) に関する問題をトラブルシューティングし、Officeの条件やエラーを処理する方法について説明します。
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: 1587f747ab3128904d4b287581f6f909f3fdb216ec03a711b30d994ca696eb03
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/07/2021
ms.locfileid: "57079961"
---
# <a name="troubleshoot-error-messages-for-single-sign-on-sso"></a>シングル サインオン (SSO) のエラー メッセージのトラブルシューティング

この記事では、Office アドインのシングル サインオン (SSO) に関する問題のトラブルシューティング方法と、SSO が有効なアドインによって特別な条件やエラーを確実に処理する方法について説明します。

> [!NOTE]
> 現在、シングル サインオン API は Word、Excel、Outlook, および PowerPoint でサポートされています。 シングル サインオン API の現在のサポート状態に関する詳細は、「[IdentityAPI の要件セット](../reference/requirement-sets/identity-api-requirement-sets.md)」を参照してください。
> Outlook アドインで作業している場合は、Microsoft 365 テナントの先進認証が有効になっていることを確認してください。 この方法の詳細については、「[Exchange Online: テナントの先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)」を参照してください。

## <a name="debugging-tools"></a>デバッグ ツール

開発時は、アドインの Web サービスからの HTTP 要求および応答を傍受して表示することができるツールを使用することを強くお勧めします。最も一般的なものは、次の 2 つです。

- [Fiddler](https://www.telerik.com/fiddler): 無料 ([ドキュメント](https://docs.telerik.com/fiddler/configure-fiddler/tasks/configurefiddler))
- [Charles](https://www.charlesproxy.com): 30 日間無料  ([ドキュメント](https://www.charlesproxy.com/documentation/))。

## <a name="causes-and-handling-of-errors-from-getaccesstoken"></a>getAccessToken からのエラーの原因と処理

このセクションで説明するエラー処理の例については、次を参照してください。

- [Office-Add-in-ASPNET-SSO の HomeES6.js](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/HomeES6.js)
- [Office-Add-in-NodeJS-SSO の ssoAuthES6.js](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/public/javascripts/ssoAuthES6.js)

### <a name="13000"></a>13000

[getAccessToken](../develop/sso-in-office-add-ins.md#sso-api-reference) API は、このアドインまたは Office バージョンではサポートされていません。

- この Office のバージョンは、SSO をサポートしていません。 必要なバージョンは、Microsoft 365チャネルでサブスクリプションに追加されます。
- アドインのマニフェストに適切な [WebApplicationInfo](../reference/manifest/webapplicationinfo.md) セクションがありません。

アドインがこのエラーに対応するには、ユーザー認証の代替システムにフォールバックする必要があります。 詳細については、「[要件とベスト プラクティス](../develop/sso-in-office-add-ins.md#requirements-and-best-practices)」を参照してください。

### <a name="13001"></a>13001

ユーザーは Office にサインインしていません。 ほとんどのシナリオでは、`AuthOptions` パラメーターでオプション `allowSignInPrompt: true` を渡して、このエラーが表示されないようにすることをお勧めします。

ただし、例外がある場合があります。 たとえば、ユーザーが *すでに* Office にログインしている *場合にのみ*、ユーザーがログインしていることが必要な機能とともにアドインが開くようにするとします。 ユーザーがログインしていない場合は、ユーザーがサインインしていることが必要のない別の機能のセットとともにアドインが開くようにするとします。 この場合、アドインが起動する際に実行されるロジックでは、`allowSignInPrompt: true` を含めずに `getAccessToken` が呼び出されます。 別の機能のセットを表示するようにアドインに指示するためのフラグとして、13001エラーを使用します。

別の方法として、13001 に対応するために、ユーザー認証の代替システムにフォールバックすることもできます。 これにより、ユーザーを Office ではなく AAD にサインインさせられます。

このエラーは **Office on the web** では一度も確認されていません。 ユーザーの Cookie が失効すると、**Office on the web** はエラー 13006 を返します。

### <a name="13002"></a>13002

ユーザーが、同意ダイアログの **[キャンセル]** を選択するなどして、サインインまたは同意を中止しました。

- アドインがユーザーのサインイン (または同意) の必要がない機能を提供している場合、コードはこのエラーをキャッチし、アドインが継続して実行するようにしなければなりません。
- 同意済みのサインインしているユーザーがアドインで必要な場合は、コードは [サインイン] ボタンを表示させる必要があります。

### <a name="13003"></a>13003

ユーザーの種類がサポートされていません。 ユーザーが有効な Microsoft アカウントOfficeまたは仕事用アカウントを使用Microsoft 365 Educationサインインしていない。 このエラーは、Office がオンプレミス ドメイン アカウントで実行されている場合に発生する可能性があります。 コードでは、ユーザー認証の代替システムにフォールバックする必要があります。 このOutlook、ユーザーのテナントでモダン認証が無効になっている[](/exchange/clients-and-mobile-in-exchange-online/enable-or-disable-modern-authentication-in-exchange-online)場合にも、このエラーが発生Exchange Online。 詳細については、「[要件とベスト プラクティス](../develop/sso-in-office-add-ins.md#requirements-and-best-practices)」を参照してください。

### <a name="13004"></a>13004

無効なリソースです。 (このエラーは開発でのみ表示される必要があります)。アドイン マニフェストが正しく構成されていません。 マニフェストを更新してください。 詳細については、「[Office アドインのマニフェストを検証する](../testing/troubleshoot-manifest.md)」を参照してください。 最も一般的な問題は、**Resource** 要素 (**WebApplicationInfo** 要素内) にアドインのドメインと一致しないドメインがあることです。 Resource 値のプロトコル部分は "https" ではなく "api" である必要があります。ドメイン名の他のすべての部分は (ポートがある場合はそれも含めて)、アドインと同じである必要があります。

### <a name="13005"></a>13005

無効な許可です。 このエラーは、通常、Office がアドインの Web サービスに対して事前に承認されていないことを意味します。 詳細については、「[サービス アプリケーションを作成する](sso-in-office-add-ins.md#create-the-service-application)」および「[Azure AD v2.0 エンドポイントにアドインを登録する](register-sso-add-in-aad-v2.md)」を参照してください。 このエラーは、ユーザーが自分の `profile` に対するアクセス許可をサービス アプリケーションに与えていない場合、または同意を取り消した場合にも発生する可能性があります。 コードでは、ユーザー認証の代替システムにフォールバックする必要があります。

開発中の場合、別の原因として、アドインを使用する Internet Explorer およびユーザーが自己署名証明書を使用していることが考えられます。 (アドインによって使用されているブラウザーを特定するには、「[Office アドインによって使用されるブラウザー](../concepts/browsers-used-by-office-web-add-ins.md)」を参照してください)。

### <a name="13006"></a>13006

クライアント エラーです。 このエラーは **Office on the web** でのみ確認されています。 コードでは、サイン アウトし、次に Office のブラウザー セッションを再起動することをユーザーに促す必要があります。

### <a name="13007"></a>13007

アプリケーションOfficeアドインの Web サービスへのアクセス トークンを取得できなかった。

- 開発中にこのエラーが発生する場合は、アドインの登録とアドイン マニフェストで `profile` のアクセス許可および (MSAL.NET を使用している場合は) `openid` のアクセス許可が指定されていることを確認してください。 詳細については、「[Azure AD v2.0 エンドポイントにアドインを登録する](register-sso-add-in-aad-v2.md)」を参照してください。
- 運用環境では、このエラーの原因として考えられることがいくつかあります。 その一部を次に示します。
  - ユーザーは Microsoft アカウント ID を持っています。
  - Microsoft 365 Education または作業用アカウントで他の 13xxx エラーの 1 つが発生する状況によっては、MSA を使用すると 13007 が発生します。

  これらのすべてのケースでは、コードでは、ユーザー認証の代替システムにフォールバックする必要があります。

### <a name="13008"></a>13008

ユーザーは、前回の `getAccessToken` の呼び出しが完了する前に `getAccessToken` を呼び出す操作をトリガーしました。 このエラーは **Office on the web** でのみ確認されています。 コードでは、ユーザーに前の操作が完了した後に操作を繰り返すよう要求する必要があります。

### <a name="13010"></a>13010

ユーザーが Microsoft Edge または Internet Explorer で Office のアドインを実行しています。 ユーザーのドメインMicrosoft 365ドメインは、ブラウザーの設定で別の `login.microsoftonline.com` セキュリティ 領域に入っています。 このエラーは **Office on the web** でのみ確認されています。 このエラーが返された場合、ユーザーには、これについて説明するエラーとゾーンの構成を変更する方法に関するページへのリンクが表示されています。 アドインがユーザーのサインインを必要としない機能を提供している場合、コードでは、このエラーをキャッチして、アドインの実行を続行する必要があります。

### <a name="13012"></a>13012

考えられる原因は複数あります。

- アドインは、`getAccessToken` API をサポートしていないプラットフォーム上で実行されています。 たとえば、iPad 上ではサポートされていません。 「Identity [API の要件セット」も参照してください](../reference/requirement-sets/identity-api-requirement-sets.md)。
- `getAccessToken` への呼び出しで `forMSGraphAccess` オプションが渡され、ユーザーが AppSource からアドインを取得しました。 このシナリオでは、アドインが必要とする Microsoft Graph スコープ (権限) について、テナント管理者はアドインに同意していません。 Office では、ユーザーに求めることができるのは AAD `profile` スコープへの同意のみであるため、`allowConsentPrompt` を使用して `getAccessToken` を取り消しても問題は解決できません。

コードでは、ユーザー認証の代替システムにフォールバックする必要があります。

開発中は、アドインは Outlook でサイドロードされ、`getAccessToken` への呼び出しで `forMSGraphAccess` オプションが渡されます。

### <a name="13013"></a>13013

短い時間で何度も呼び出されたので、Officeの呼び `getAccessToken` 出しを調整しました。 これは通常、メソッドの呼び出しの無限ループによって発生します。 メソッドの呼び出しが推奨されるシナリオがあります。 ただし、コードはカウンター変数またはフラグ変数を使用して、メソッドが繰り返し呼び出されていないことを確認する必要があります。 同じ "再試行" コード パスが再び実行されている場合、コードはユーザー認証の別のシステムに戻る必要があります。 コード例については、変数がどのように使用されるのかHomeES6.js`retryGetAccessToken` を[](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Scripts/HomeES6.js)参照[ssoAuthES6.js。 ](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/public/javascripts/ssoAuthES6.js)

### <a name="50001"></a>50001

このエラー (`getAccessToken` に固有ではありません) は、ブラウザーが office.js ファイルの古いコピーをキャッシュしていることを示す可能性があります。 開発中の場合は、ブラウザーのキャッシュをクリアしてください。 また、Office のバージョンが古いため、SSO をサポートしていない可能性も考えられます。 Windows での最小バージョンは、16.0.12215.20006 です。 Mac では、16.32.19102902 です。

運用環境のアドインの場合、アドインがこのエラーに対応するには、ユーザー認証の代替システムにフォールバックする必要があります。 詳細については、「[要件とベスト プラクティス](../develop/sso-in-office-add-ins.md#requirements-and-best-practices)」を参照してください。

## <a name="errors-on-the-server-side-from-azure-active-directory"></a>Azure Active Directory からのサーバー側のエラー

このセクションで説明するエラー処理の例については、次を参照してください。

- [Office-Add-in-ASPNET-SSO](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO)
- [Office-Add-in-NodeJS-SSO](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO)

### <a name="conditional-access--multifactor-authentication-errors"></a>条件付きアクセスおよび多要素認証のエラー

AAD および Microsoft 365 の ID の特定の構成では、Microsoft Graph でアクセスできる一部のリソースで、ユーザーの Microsoft 365 テナントが要求しない場合でも、多要素認証 (MFA) を必要とすることができます。 AAD は、MFA で保護されたリソースへのトークンの要求を、代理フロー経由で受け取ると、アドインの Web サービスに `claims` プロパティを含む JSON メッセージを返します。 claims プロパティには、さらに必要となる認証要素の情報が含まれています。

コードは、この `claims` プロパティについてテストする必要があります。 アドインのアーキテクチャによっては、クライアント側でテストすることができます。または、サーバー側でテストし、クライアントにリレーすることができます。 SSO アドインの認証は Office によって処理されるため、この情報がクライアントで必要になります。この情報をサーバー側からリレーする場合、クライアントへのメッセージは、エラー (`500 Server Error` や `401 Unauthorized` など) または成功応答の本文 (`200 OK` など) のいずれかになります。 どちらの場合でも、アドインの Web API に対する、コードによるクライアント側の AJAX 呼び出しのコールバック (失敗または成功) が、この応答をテストする必要があります。

アーキテクチャに関係なく、クレーム値が AAD から送信されている場合、コードはパラメーターでオプションを呼び出して `getAccessToken` `authChallenge: CLAIMS-STRING-HERE` 渡す必要 `options` があります。 AAD がこの文字列を認識すると、ユーザーに追加の要素を入力するよう促してから、代理フローで受け入れられる新しいアクセス トークンを返します。

### <a name="consent-missing-errors"></a>同意なしエラー

AAD に、ユーザー (またはテナント管理者) がアドインに (Microsoft Graph リソースに対して) 同意した記録がない場合、AAD はエラー メッセージを Web サービスに送信します。 コードは、 (`403 Forbidden` 応答の本文などで) クライアントに指示する必要があります。

管理者のみが同意できる Microsoft Graph のスコープがアドインで必要な場合は、コードはエラーをスローする必要があります。 唯一必要なスコープに対して同意できるのがユーザーである場合は、コードはユーザー認証の代替システムにフォールバックする必要があります。

### <a name="invalid-or-missing-scope-permission-errors"></a>無効または見つからないスコープ (アクセス許可) のエラー

この種類のエラーが表示されるのは、開発中のみである必要があります。

- サーバー側のコードでは、`403 Forbidden` 応答をクライアントに送り、そのエラーのログをコンソールで作成するか、ログに記録する必要があります。
- アドイン マニフェストの[範囲](../reference/manifest/scopes.md)セクションで、必要なすべてのアクセス許可が指定されていることを確認してください。 また、アドインの Web サービスの登録で同じアクセス許可が指定されていることを確認してください。 スペルミスもチェックしてください。 詳細については、「[Azure AD v2.0 エンドポイントにアドインを登録する](register-sso-add-in-aad-v2.md)」を参照してください。

### <a name="invalid-audience-error-in-the-access-token-not-the-bootstrap-token"></a>(ブートストラップ トークンではなく) アクセス トークンでの無効な対象ユーザーのエラー

サーバー側のコードは、`403 Forbidden` 応答をクライアントに送って、ユーザーにわかりやすいメッセージを提示しなければなりません。場合によっては、エラーについて、コンソールでログを作成するか、ログに記録する必要もあります。
