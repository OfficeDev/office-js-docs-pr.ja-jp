---
title: 'シナリオ: サービスにシングル サインオンを実装する'
description: Outlook アドインが提供するシングル サインオン トークンと Exchange ID トークンを使用して、サービスに SSO を実装する方法について説明します。
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: 07e3672860a063f8bf14ef9a653f97a8244e2c3e
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671451"
---
# <a name="scenario-implement-single-sign-on-to-your-service-in-an-outlook-add-in"></a>シナリオ: Outlook アドインでサービスにシングル サインオンを実装する

この記事では、独自のバックエンド サービスにシングル サインオンの実装を提供するために、[シングル サインオン アクセス トークン](authenticate-a-user-with-an-sso-token.md)と [Exchange ID トークン](authenticate-a-user-with-an-identity-token.md)を同時に使用する推奨の方法について説明します。 両方のトークンを同時に使用することで、SSO アクセス トークンが使用できる場合はその利点を活用し、そのトークンが使用できない場合でもアドインが確実に動作するようにします。SSO アクセス トークンは、ユーザーがそのトークンをサポートしていないクライアントに切り替えたときや、ユーザーのメールボックスがオンプレミスの Exchange サーバーにある場合などは使用できません。

この記事のアイデアを実装するサンプル アドインについては、「Outlook [SSO」を参照してください](https://github.com/OfficeDev/Outlook-Add-in-SSO)。


> [!NOTE]
> 現在、シングル サインオン API は Word、Excel、Outlook, および PowerPoint でサポートされています。 シングル サインオン API の現在のサポート状態に関する詳細は、「[IdentityAPI の要件セット](../reference/requirement-sets/identity-api-requirement-sets.md)」を参照してください。
> Outlook アドインで作業している場合は、Microsoft 365 テナントの先進認証が有効になっていることを確認してください。 この方法の詳細については、「[Exchange Online: テナントの先進認証を有効にする方法](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)」を参照してください。


## <a name="why-use-the-sso-access-token"></a>SSO アクセス トークンを使用する理由

Exchange ID トークンはアドインの API のすべての要件セットで使用できるため、このトークンだけを使用して、SSO トークンは完全に無視してしまうことがあります。 ただし、SSO トークンには Exchange ID トークンよりも優れた点がいくつかあるため、使用できるときには SSO トークンが推奨の方法になります。

- SSO トークンは、標準の OpenID 形式を使用して、Azure が発行します。 そのため、このトークンの検証プロセスは、とても簡単になります。 それに比べて、Exchange ID トークンは JSON Web トークン標準に基づいたカスタム形式を使用するため、トークンの検証にカスタムの操作が必要になります。
- SSO トークンは、Microsoft Graph のトークンを取得するためにバックエンドで使用できます。このとき、ユーザーが追加のサインイン操作を実行する必要はありません。
- SSO トークンは、ユーザーの表示名など豊富な ID 情報を提供します。

## <a name="add-in-scenario"></a>アドインのシナリオ

この例では、アドイン UI およびスクリプト (HTML + JavaScript) の両方で構成されているアドインと、そのアドインで呼び出すバックエンド Web API について考えてみます。 バックエンド Web API は、[Microsoft Graph API](/graph/overview) と Contoso Data API (架空のサード パーティ製 API) の両方を呼び出します。 Microsoft Graph API と同じように、Contoso Data API も OAuth 認証を必要とします。 要件は、アクセス トークンの有効期限が切れるたびに バックエンド Web API がユーザーに資格情報を求めるダイアログを表示することなく、両方の API を呼び出せるようにすることです。

そのために、バックエンド API は、ユーザーに関するセキュリティで保護されたデータベースを作成します。 それぞれのユーザーごとに、このデータベース内のエントリが割り当てられます。バックエンドは、このデータベースに Microsoft Graph API と Contoso Data API の長期間有効な更新トークンを保存します。 次に示す JSON マークアップは、データベース内のユーザーのエントリを表しています。

```JSON
{
  "userDisplayName": "...",
  "ssoId": "...",
  "exchangeId": "...",
  "graphRefreshToken": "...",
  "contosoRefreshToken": "..."
}
```

アドインは、バックエンド Web API を呼び出すたびに、SSO アクセス トークン (使用可能な場合) または Exchange ID トークン (SSO トークンが使用不可の場合) のどちらかを含めます。

### <a name="add-in-startup"></a>アドインのスタートアップ

1. アドインの起動時に、アドインはバックエンド Web API に要求を送信して、ユーザーが登録されているかどうか (ユーザー データベースに関連付けられたレコードがあるか) と、その API が Graph および Contoso の更新トークンを保持しているかを判断します。 アドインは、この呼び出に SSO トークン (使用可能な場合) と ID トークンの両方を含めます。

1. Web API は、「[Outlook アドインでシングル サインオン トークンを使用してユーザーを認証する](authenticate-a-user-with-an-sso-token.md)」と「[Exchange の ID トークンを使用してユーザーを認証する](authenticate-a-user-with-an-identity-token.md)」に示した方法を使用して、両方のトークンを検証して一意の ID を生成します。

1. SSO トークンが提供された場合、Web API は SSO トークンから生成された一意の ID と一致する `ssoId` 値を保持するエントリについて、ユーザー データベースを照会します。
   - エントリが存在しない場合は、次の手順に進みます。
   - エントリが存在する場合は、手順 5 に進みます。

1. Web API は、Exchange ID トークンから生成された一意の ID と一致する `exchangeId` 値を保持するエントリについて、データベースを照会します。
   - エントリが存在しているときに、SSO トークンが提供されていた場合は、データベース内のユーザーのレコードを更新して、`ssoId` の値を SSO トークンから生成された一意の ID に設定して、手順 5 に進みます。
   - エントリが存在しているときに、SSO トークンが提供されていなかった場合は、手順 5 に進みます。
   - エントリが存在しないときには、新しいエントリを作成します。 `ssoId` を SSO トークン (使用可能な場合) から生成された一意の ID に設定し、`exchangeId` を Exchange ID トークンから生成された一意の ID に設定します。

1. ユーザーの `graphRefreshToken` 値で有効な更新トークンについて調べます。
   - この値が無効または見つからないときに、SSO トークンが提供されていた場合は、[OAuth2 On-Behalf-Of フロー](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of)を使用して Graph のアクセス トークンと更新トークンを取得します。 ユーザーの `graphRefreshToken` 値に更新トークンを保存します。

1. `graphRefreshToken` と `contosoRefreshToken` の両方で有効な更新トークンについて調べます。
   - 両方の値が有効な場合は、ユーザーが既に登録および構成されていることを示す応答をアドインに返します。
   - どちらかの値が無効な場合は、ユーザーの設定が必要なことと、構成が必要になるサービスがどちらか (Graph または Contoso) を示す応答をアドインに返します。

1. アドインは、この応答を確認します。
   - ユーザーが既に登録および構成されている場合、アドインは通常の操作を続行します。
   - ユーザーの設定が必要な場合、アドインは「セットアップ」モードに移行して、ユーザーにアドインを承認するように求めるダイアログを表示します。

### <a name="authorize-the-backend-web-api"></a>バックエンド Web API の承認

Microsoft Graph API と Contoso Data API を呼び出すバックエンド Web API を承認する手順は、1 回だけ実施されるようにすることが理想的です。そうすることで、ユーザーにサインインを求めるダイアログの表示を最小限に抑えるようにします。

バックエンド Web API からの応答に基づいて、アドインは Microsoft Graph API または Contoso Data API、またはその両方のユーザーを承認することが必要になる場合があります。 どちらの API も OAuth2 認証を使用するため、その方法はどちらも同様になります。

1. アドインは、API の使用を承認する必要があることをユーザーに通知して、そのプロセスを開始するためにリンクまたはボタンをクリックするように求めます。

    > [!NOTE]
    > [Outlook アドイン SSO](https://github.com/OfficeDev/Outlook-Add-in-SSO)のアドインの例は、ダイアログ API と[office-js-helpers](https://github.com/OfficeDev/office-js-helpers)ライブラリをオプションとして使用して[、API](/javascript/api/office/office.ui#displayDialogAsync_startAddress__options__callback_)の[OAuth2 承認](/azure/active-directory/develop/active-directory-protocols-oauth-code)コード フローを開始する方法を示しています。

1. このフローが完了すると、アドインは更新トークンをバックエンド Web API に送信し、SSO トークン (使用可能な場合) または Exchange ID トークンを含めます。

1. バックエンド Web API はデータベース内でユーザーを見つけて、該当する更新トークンを更新します。

1. アドインは、通常の操作を続行します。

### <a name="normal-operation"></a>通常の操作

アドインがバックエンド Web API を呼び出すときには、SSO トークンまたは Exchange ID トークンを必ず含めます。 バックエンド Web API は、このトークンでユーザーを見つけて、保存されている更新トークンを使用して Microsoft Graph API と Contoso Data API のアクセス トークンを取得します。 更新トークンが有効な間、ユーザーは再度サインインする必要がなくなります。
