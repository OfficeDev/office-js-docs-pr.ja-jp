---
title: SSO を使用せずに Microsoft Graph を承認する
description: SSO を使用せずに Microsoft Graph を承認する方法
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: 99d300d0155ba9a117efda5d31ef068a41eb86a9
ms.sourcegitcommit: 8546889a759590c3798ce56e311d9e46f0171413
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/04/2021
ms.locfileid: "50104834"
---
# <a name="authorize-to-microsoft-graph-without-sso"></a>SSO を使用せずに Microsoft Graph を承認する

アドインは、Azure Active Directory (Azure AD) から Microsoft Graph へのアクセス トークンを取得することで、Microsoft Graph データへの承認を取得できます。 認証コード フローまたは暗黙的フローは、他の Web アプリケーションの場合と同様に使用しますが、1 つの例外があります。Azure AD では、iframe でサインイン ページを開くことができません。 Office アドインが *Office on the web* で実行されている場合、作業ウィンドウとして iFrame が使用されます。 つまり、Azure AD ログイン画面を開き、そのダイアログ API で開くダイアログ Office必要があります。 これは、認証と承認ヘルパー ライブラリの使用方法に影響します。 詳細については、「[Office ダイアログ API を使用して認証および承認する](auth-with-office-dialog-api.md)」を参照してください。

Azure AD を使用した認証のプログラミングの詳細については、Microsoft ID プラットフォーム [(v2.0)](/azure/active-directory/develop/v2-overview)の概要を参照してください。このドキュメント セットには、チュートリアルとガイド、および関連するサンプルへのリンクが含まれています。 繰り返しますが、Office ダイアログ ボックスで実行するようにサンプルのコードを調整して、Office ダイアログ ボックスが作業ウィンドウとは別のプロセスで実行されるという事実を説明する必要がある場合があります。

コードが Microsoft Graph へのアクセス トークンを取得した後、アクセス トークンをダイアログ ボックスから作業ウィンドウに渡すか、トークンをデータベースに格納して、トークンが使用可能な作業ウィンドウにシグナルを送信します。 (詳細 [については、「Office ダイアログ API を使用した認証](auth-with-office-dialog-api.md) 」を参照してください)。作業ウィンドウのコードは、Microsoft Graph からデータを要求し、それらの要求にトークンを含します。 Microsoft Graph と Microsoft Graph SDK の呼び出しの詳細については [、Microsoft Graph のドキュメントを参照してください](/graph/)。

## <a name="recommended-libraries-and-samples"></a>推奨されるライブラリおよびサンプル

SSO を使用せずに Microsoft Graph にアクセスする場合は、次のライブラリを使用することをお勧めします。

- .NET ベースのフレームワーク (.NET Core や ASP.NET など) のサーバー側を使用するアドインの場合は、[MSAL.NET](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki#conceptual-documentation) を使用します。
- NodeJS ベースのサーバー側を使用するアドインの場合は、[Passport Azure AD](https://github.com/AzureAD/passport-azure-ad) を使用します。
- 暗黙的なフローを使用するアドインの場合は、[msal.js](https://github.com/AzureAD/microsoft-authentication-library-for-js/wiki) を使用します。

Microsoft ID プラットフォーム (以前は AAD v.2.0) を使用するための推奨ライブラリの詳細については、「[Microsoft ID プラットフォームの認証ライブラリ](/azure/active-directory/develop/reference-v2-libraries)」を参照してください。

次のサンプルでは、Office アドインから Microsoft Graph のデータを取得します。

- [Office アドイン Microsoft Graph ASP.NET](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)
- [Outlook アドイン Microsoft Graph ASP.NET](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)
- [Office アドイン Microsoft Graph React](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Samples/auth/Office-Add-in-Microsoft-Graph-React)
