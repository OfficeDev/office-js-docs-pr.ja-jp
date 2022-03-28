---
title: Office アドインにおける認証と承認の概要
description: Office アドインでの認証と承認のしくみについて説明します。
ms.date: 01/25/2022
ms.localizationpriority: high
ms.openlocfilehash: ba7f55a0b8ca163b994bcfb91879c675b777a7c9
ms.sourcegitcommit: b66ba72aee8ccb2916cd6012e66316df2130f640
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2022
ms.locfileid: "64483642"
---
# <a name="overview-of-authentication-and-authorization-in-office-add-ins"></a>Office アドインにおける認証と承認の概要

Office アドインでは、既定で匿名アクセスが許可されますが、ユーザーに対して、アドインを使用するために Microsoft アカウント、Microsoft 365 Education または職場アカウント、その他の一般的なアカウントでサインインすることを要求できます。 これによりユーザーの確認がアドインで可能になることから、このタスクはユーザー認証と呼ばれています。

アドインは、Microsoft Graph データ (Microsoft 365 プロファイル、OneDrive ファイル、SharePoint データなど) または Google、Facebook、LinkedIn、SalesForce、GitHub などの他の外部ソースのデータにアクセスすることにユーザーの同意を得ることもできます。このタスクは、ユーザーではなく、承認されている *アドイン* であるため、アドイン (またはアプリ) 承認と呼ばれます。

## <a name="key-resources-for-authentication-and-authorization"></a>認証と承認のための主要なリソース

このドキュメントでは、認証と承認を正常に実装するために Office アドインを作成および構成する方法について説明します。 ただし、言及されている多くの概念とセキュリティ テクノロジは、このドキュメントの範囲外です。 たとえば、OAuth フロー、トークン キャッシュ、ID 管理などの一般的なセキュリティの概念については、ここでは説明しません。 また、このドキュメントには、Microsoft Azure または Microsoft ID プラットフォームに固有の情報は記載されていません。 これらの領域の詳細情報が必要な場合は、次のリソースを参照することをお勧めします。

- [Microsoft ID プラットフォーム](/azure/active-directory/develop)
- [Microsoft ID プラットフォームのサポートとヘルプ オプション](/azure/active-directory/develop/developer-support-help-options)
- [Microsoft ID プラットフォームにおける OAuth 2.0 と OpenID Connect プロトコル](/azure/active-directory/develop/active-directory-v2-protocols)

## <a name="sso-scenarios"></a>SSO のシナリオ

シングル サインオン (SSO) を使用すると、Office に 1 回サインインするだけで済むため、ユーザーにとって便利です。 アドインに個別にサインインする必要はありません。 SSO はすべてのバージョンの Office でサポートされているわけではないため、[Microsoft ID プラットフォームを使用して](#authenticate-with-the-microsoft-identity-platform)別のサインイン 方法を実装する必要があります。 サポートされている Office バージョンの詳細については、「[Identity API 要件セット](/javascript/api/requirement-sets/identity-api-requirement-sets)」を参照してください。

### <a name="get-the-users-identity-through-sso"></a>SSO を使用してユーザーの ID を取得する

多くの場合、アドインではユーザーの ID のみが必要です。 たとえば、アドインをカスタマイズして、ユーザーの名前を作業ウィンドウに表示するだけという場合があります。 または、データベース内のデータにユーザーを関連付けるために一意の ID が必要な場合があります。 これは、Office からユーザーのアクセス トークンを取得するだけで実現できます。

SSO を使用してユーザーの ID を取得するには、[getAccessToken](/javascript/api/office-runtime/officeruntime.auth#office-runtime-officeruntime-auth-getaccesstoken-member(1)) メソッドを呼び出します。 このメソッドは、`preferred_username`、`name`、`sub`、`oid` などの、現在サインインしているユーザーに固有のいくつかの要求を含む ID トークンでもあるアクセス トークンを返します。 これらのプロパティの詳細については、「[Microsoft ID プラットフォームの ID トークン](/azure/active-directory/develop/id-tokens)」を参照してください。 [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#office-runtime-officeruntime-auth-getaccesstoken-member(1)) によって返されるトークンの例については、「[アクセス トークンの例](sso-in-office-add-ins.md#example-access-token)」を参照してください。

ユーザーがサインインしていない場合、Office はダイアログ ボックスを開き、Microsoft ID プラットフォームを介してユーザーにサインインを要求します。 その後、メソッドによってアクセス トークンが返されるか、ユーザーをサインインできない場合はエラーがスローされます。

ユーザーのデータを格納する必要があるシナリオでは、トークンから値を取得してユーザーを一意に識別する方法の詳細について、「[Microsoft ID プラットフォームの ID トークン](/azure/active-directory/develop/id-tokens)」を参照してください。 この値を使用して、管理しているユーザー テーブルまたはユーザー データベース内のユーザーを参照します。 ユーザー設定やユーザーのアカウントの状態などのユーザー関連情報を格納するには、データベースを使用します。 SSO を使用しているため、ユーザーは個別にアドインにサインインを行いません。このため、ユーザーのパスワードを保存する必要はありません。

SSO を使用するユーザー認証を実装する前に、「[Office アドインのシングル サインオンを有効化する](sso-in-office-add-ins.md)」の記事を十分に理解しておく必要があります。

### <a name="access-your-web-apis-through-sso"></a>SSO を使用して Web API にアクセスする

アドインに、承認されたユーザーを必要とするサーバー側 API がある場合は、[getAccessToken](/javascript/api/office-runtime/officeruntime.auth#office-runtime-officeruntime-auth-getaccesstoken-member(1)) メソッドを呼び出してアクセス トークンを取得します。 アクセス トークンは、独自の Web サーバーへのアクセスを提供します ([Microsoft Azure アプリ登録](register-sso-add-in-aad-v2.md)を使用して構成されます)。Web サーバーで API を呼び出すときは、アクセス トークンも渡してユーザーを承認します。

次のコードは、アドインの Web サーバー API に対して HTTPS GET 要求を作成してデータを取得する方法を示しています。 このコードは、作業ウィンドウなどで、クライアント側で実行されます。 最初に `getAccessToken` を呼び出してアクセス トークンを取得します。 次に、サーバー API の正しい承認ヘッダーと URL を使用して AJAX 呼び出しを作成します。

```javascript
function getOneDriveFileNames() {

    let accessToken = await Office.auth.getAccessToken();

    $.ajax({
        url: "/api/data",
        headers: { "Authorization": "Bearer " + accessToken },
        type: "GET"
    })
        .done(function (result) {
            //... work with data from the result...
        });
}
```

次のコードは、前のコード例の REST 呼び出しの /api/data ハンドラーの例を示しています。 このコードは、Web サーバーで実行されている ASP.NET コードです。 `[Authorize]` 属性により、有効なアクセス トークンをクライアントから渡すか、クライアントにエラーを返すことが求められます。

```csharp
    [Authorize]
    // GET api/data
    public async Task<HttpResponseMessage> Get()
    {
        //... obtain and return data to the client-side code...
    }
```

### <a name="access-microsoft-graph-through-sso"></a>SSO を使用して Microsoft Graph にアクセスする

一部のシナリオでは、ユーザーの ID が必要なだけでなく、ユーザーの代わりに [Microsoft Graph](/graph) リソースにアクセスする必要がある場合があります。 たとえば、ユーザーの代わりにメールを送信したり、Teams でチャットを作成したりする必要がある場合があります。 これらのアクションは、Microsoft Graph を通じて実行できます。 次の手順に従う必要があります。

1. [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#office-runtime-officeruntime-auth-getaccesstoken-member(1)) を呼び出して、SSO を使用して現在のユーザーのアクセス トークンを取得します。 ユーザーがサインインしていない場合、Office はダイアログ ボックスを開き、Microsoft ID プラットフォームを使用してユーザーをサインインします。 ユーザーがサインインする、またはユーザーが既にサインインしている場合、メソッドによりアクセス トークンが返されます。
1. アクセス トークンをサーバー側のコードに渡します。
1. サーバー側で [OAuth 2.0 On-Behalf-Of フロー](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow)を使用して、アクセス トークンを、Microsoft Graph を呼び出すために必要な委任されたユーザーの ID とアクセス許可を含む新しいアクセス トークンに交換します。

> [!NOTE]
> アクセス トークンの漏洩を避けるための最善のセキュリティ対策として、常にサーバー側で On-Behalf-Of フローを実行します。 クライアントではなく、サーバーから Microsoft Graph API を呼び出します。 クライアント側のコードにアクセス トークンを返す必要はありません。

アドインで Microsoft Graph にアクセスするための SSO の実装を開始する前に、次の記事を十分に理解しておく必要があります。

- [Office アドインのシングル サインオンを有効化する](sso-in-office-add-ins.md)
- [SSO を使用した Microsoft Graph への承認](authorize-to-microsoft-graph.md)

また、SSO を使用して Microsoft Graph にアクセスするための Office アドインの作成について説明している以下の記事のうち少なくとも 1 つに目を通してください。 その手順を実行しない場合でも、SSO と On-Behalf-Of フローの実装方法に関する重要な情報が含まれています。

- 「[シングル サインオンを使用する ASP.NET Office アドインを作成する](create-sso-office-add-ins-aspnet.md)」では、[Office アドイン ASP.NET SSO](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO) のサンプルについて説明します。
- 「[シングル サインオンを使用する Node.js Office アドインを作成する](create-sso-office-add-ins-nodejs.md)」では、[Office アドイン NodeJS SSO](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO) のサンプルについて説明します。

## <a name="non-sso-scenarios"></a>SSO 以外のシナリオ

一部のシナリオでは、SSO を使用しない場合があります。 たとえば、Microsoft ID プラットフォームとは異なる ID プロバイダーを使用して認証する必要がある場合があります。 また、SSO はすべてのシナリオでサポートされているわけではありません。 たとえば、以前のバージョンの Office では SSO がサポートされていません。 この場合は、アドインの代替認証システムに戻る必要があります。

### <a name="authenticate-with-the-microsoft-identity-platform"></a>Microsoft ID プラットフォームを使用して認証する

アドインは、認証プロバイダーとして [Microsoft ID プラットフォーム](/azure/active-directory/develop)を使用してユーザーをサインインすることができます。 ユーザーをサインインしたら、Microsoft ID プラットフォームを使用して、Microsoft が管理する [Microsoft Graph](/graph) またはその他のサービスに対してアドインを承認できます。 Office を介した SSO が利用できない場合は、この方法を代替サインイン方法として使用します。 また、SSO が使用可能な場合でも、ユーザーをアドインに個別にサインインさせるシナリオもあります。たとえば、現在 Office にサインインしている ID とは異なる ID でアドインにサインインするオプションをユーザーに与える場合などです。

Microsoft ID プラットフォームでは、サインイン ページを iframe で開くことが許可されていないことに注意してください。 Office アドインが *Office on the web* で実行されている場合、作業ウィンドウとして iFrame が使用されます。 これは、Office ダイアログ API で開かれるダイアログ ボックスを使用して、サインイン ページを開く必要があることを意味します。 このことは、認証ヘルパー ライブラリの使用方法に影響します。 詳細については、「[Office ダイアログ API を使用して認証および承認する](auth-with-office-dialog-api.md)」を参照してください。

Microsoft ID プラットフォームを使用した認証の実装の詳細については、「[Microsoft ID プラットフォーム (v2.0) の概要](/azure/active-directory/develop/v2-overview)」を参照してください。 ドキュメントには、多くのチュートリアルとガイドのほか、関連するサンプルとライブラリへのリンクが含まれています。 「[Office ダイアログ API を使用して認証および承認する](auth-with-office-dialog-api.md)」の説明にあるように、Office ダイアログ ボックスで実行するサンプル内のコードを調整する必要がある場合があります。

### <a name="access-to-microsoft-graph-without-sso"></a>SSO を使用しないで Microsoft Graph にアクセスする

Microsoft ID プラットフォームから Microsoft Graph へのアクセス トークンを取得することで、アドインの Microsoft Graph データに対する承認を取得できます。 これは、Office を介した SSO に依存することなく (または、SSO が失敗した場合、またはサポートされていない場合に) 実行できます。 詳細については、「[SSO を使用せずに Microsoft Graph にアクセスする](authorize-to-microsoft-graph-without-sso.md)」を参照してください。詳細情報とサンプルへのリンクが含まれています。

### <a name="access-to-non-microsoft-data-sources"></a>Microsoft 以外のデータ ソースへのアクセス

大手のオンライン サービス (Google、Facebook、LinkedIn、SalesForce、GitHub など) では、開発者は、ユーザーが自分のアカウントに別のアプリケーションからアクセスできるようにすることが可能です。 これにより、開発者はこれらのサービスを Office アドインに含めることができます。 アドインでこれを実行する方法の概要については、「[Authorize external services in your Office Add-in (Office アドインで外部サービスを承認する)](auth-external-add-ins.md)」を参照してください。

> [!IMPORTANT]
> コーディングを開始する前に、データ ソースのサインイン ページを iframe で開くことが許可されているかどうか確認してください。 Office アドインが *Office on the web* で実行されている場合、作業ウィンドウとして iFrame が使用されます。 データ ソースのサインイン ページを iframe で開くことが許可されていない場合は、Office ダイアログ API で開かれるダイアログ ボックスでサインイン ページを開く必要があります。 詳細については、「[Office ダイアログ API を使用して認証および承認する](auth-with-office-dialog-api.md)」を参照してください。

## <a name="see-also"></a>関連項目

- [Microsoft ID プラットフォームのドキュメント](/azure/active-directory/develop/)
- [Microsoft ID プラットフォームのアクセス トークン](/azure/active-directory/develop/access-tokens)
- [Microsoft ID プラットフォームにおける OAuth 2.0 と OpenID Connect プロトコル](/azure/active-directory/develop/active-directory-v2-protocols)
- [Microsoft ID プラットフォームと OAuth 2.0 On-Behalf-Of フロー](/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow)
- [JSON Web Token (JWT)](https://en.wikipedia.org/wiki/JSON_Web_Token)
- [JSON Web Token ビューアー](https://jwt.ms/)
