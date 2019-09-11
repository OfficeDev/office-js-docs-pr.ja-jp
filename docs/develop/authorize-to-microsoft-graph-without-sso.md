---
title: SSO を使用せずに Microsoft Graph を承認する
description: ''
ms.date: 08/07/2019
localization_priority: Priority
ms.openlocfilehash: 9636077553904e7250cf1d6dc740febe9eac61e2
ms.sourcegitcommit: 24303ca235ebd7144a1d913511d8e4fb7c0e8c0d
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/11/2019
ms.locfileid: "36838488"
---
# <a name="authorize-to-microsoft-graph-without-sso"></a>SSO を使用せずに Microsoft Graph を承認する

Azure Active Directory (AAD) から Graph へのアクセス トークンを取得することで、アドイン用の Microsoft Graph データへの認証を取得できます。 これを行うには、他の Web アプリケーションの場合と同様に (1 つの例外を除く)、認証コード フローか暗黙的なフローのいずれかを使用します。AAD では、ログイン ページを iframe で開くことを許可しません。 Office アドインが *Office on the web* で実行されている場合、作業ウィンドウとして iFrame が使用されます。 これは、AAD のログイン画面は、Office ダイアログ API を使用して開かれるダイアログで開く必要があることを意味します。 これは、認証と承認ヘルパー ライブラリの使用方法に影響します。 詳細については、「[Office ダイアログ API を使用して認証と承認を行う](auth-with-office-dialog-api.md)」を参照してください。

AAD での認証のプログラミングの詳細については、「[Microsoft ID プラットフォーム (v2.0) の概要](/azure/active-directory/develop/v2-overview)」を参照してください。 このドキュメント セットには、多くのチュートリアルやガイドのほか、関連するサンプルへのリンクが含まれています。 繰り返しますが、Office ダイアログで実行するようにサンプルのコードを調整して、ダイアログが作業ウィンドウとは別のプロセスで実行されるという事実を説明する必要がある場合があります。

Microsoft Graph へのアクセス トークンを取得した後、コードはアクセス トークンをダイアログから作業ウィンドウに渡すか、トークンをデータベースに保存し、そこでトークンが使用可能であることを作業ウィンドウに通知します。 (詳細については、「[Office ダイアログ API を使用して認証と承認を行う](auth-with-office-dialog-api.md)」を参照してください)。作業ウィンドウのコードは、Microsoft Graph のデータを要求し、それらの要求にトークンを含めます。 Microsoft Graph と Microsoft Graph の SDK の呼び出しの詳細については、「[Microsoft Graph ドキュメント](/graph/)」を参照してください。

## <a name="recommended-libraries-and-samples"></a>推奨されるライブラリおよびサンプル

SSO を使用せずに Microsoft Graph にアクセスする場合は、次のライブラリを使用することをお勧めします。

- .NET ベースのフレームワーク (.NET Core や ASP.NET など) のサーバー側を使用するアドインの場合は、[MSAL.NET](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki#conceptual-documentation) を使用します。
- NodeJS ベースのサーバー側を使用するアドインの場合は、[Passport Azure AD](https://github.com/AzureAD/passport-azure-ad) を使用します。
- 暗黙的なフローを使用するアドインの場合は、[msal.js](https://github.com/AzureAD/microsoft-authentication-library-for-js/wiki) を使用します。

Microsoft ID プラットフォーム (以前は AAD v.2.0) を使用するための推奨ライブラリの詳細については、「[Microsoft ID プラットフォームの認証ライブラリ](/azure/active-directory/develop/reference-v2-libraries)」を参照してください。

次のサンプルでは、Office アドインから Microsoft Graph のデータを取得します。

- [Office アドイン Microsoft Graph ASP.NET](https://github.com/OfficeDev/office-add-in-microsoft-graph-aspnet)
- [Outlook アドイン Microsoft Graph ASP.NET](https://github.com/OfficeDev/outlook-add-in-microsoft-graph-aspnet)

