---
title: SSO を使用せずに Microsoft Graph を承認する
description: SSO を使用せずに Microsoft Graph を承認する方法
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: 828779a766c41088435ff5fdfa693e1d9939c710
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/12/2020
ms.locfileid: "41949662"
---
# <a name="authorize-to-microsoft-graph-without-sso"></a>SSO を使用せずに Microsoft Graph を承認する

Azure Active Directory (AAD) から Graph へのアクセス トークンを取得することで、アドインの Microsoft Graph データへの承認を取得できます。 他の Web アプリケーションの場合と同様に (1 つの例外を除く)、認証コード フローか暗黙的なフローを使用します。AAD では、ログイン ページを iframe で開くことを許可しません。 Office アドインが *Office on the web* で実行されている場合、作業ウィンドウとして iFrame が使用されます。 これは、AAD のログイン画面は、Office ダイアログ API を使用して開かれるダイアログ ボックスで開く必要があることを意味します。 これは、認証と承認ヘルパー ライブラリの使用方法に影響します。 詳細については、「[Office ダイアログ API を使用して認証および承認する](auth-with-office-dialog-api.md)」を参照してください。

AAD を使用した認証のプログラミングの詳細については、「[Microsoft ID プラットフォーム (v 2.0) の概要](/azure/active-directory/develop/v2-overview)」を参照してください。このドキュメント セットでは、チュートリアルやガイド、関連するサンプルへのリンクを見つけることができます。 繰り返しますが、Office ダイアログ ボックスで実行するようにサンプルのコードを調整して、Office ダイアログ ボックスが作業ウィンドウとは別のプロセスで実行されるという事実を説明する必要がある場合があります。

Graph へのアクセス トークンを取得した後、コードはアクセス トークンをダイアログから作業ウィンドウに渡すか、トークンをデータベースに保存し、トークンが使用可能であることを作業ウィンドウに通知します。 (詳細については、「[Office ダイアログ API を使用して認証と承認を行う](auth-with-office-dialog-api.md)」を参照してください)。作業ウィンドウのコードは、Graph のデータを要求し、それらの要求にトークンを含めます。 Graph とGraph SDK の呼び出しの詳細については、「[Microsoft Graph ドキュメント](/graph/)」を参照してください。

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
