---
title: Office アドインから Microsoft Graph に対する承認
description: Office アドインから Microsoft Graph に承認する方法について説明します。
ms.date: 01/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 37dd4be3acb92dc7884972de923d94936fa870f4
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810170"
---
# <a name="authorize-to-microsoft-graph-from-an-office-add-in"></a>Office アドインから Microsoft Graph に対する承認

アドインは、Microsoft ID プラットフォームから Microsoft Graph へのアクセス トークンを取得することで、Microsoft Graph データに対する承認を取得できます。 承認コード フローまたは暗黙的フローは、他の Web アプリケーションと同じように使用しますが、1 つの例外があります。Microsoft ID プラットフォームでは、サインイン ページが iframe で開くことを許可しません。 Office アドインが *Office on the web* で実行されている場合、作業ウィンドウは iframe です。 つまり、Office ダイアログ API を使用して、ダイアログ ボックスでサインイン ページを開く必要があります。 これは、認証と承認ヘルパー ライブラリの使用方法に影響します。 詳細については、「[Office ダイアログ API を使用して認証および承認する](auth-with-office-dialog-api.md)」を参照してください。

> [!NOTE]
> SSO を実装し、Microsoft Graph にアクセスする予定の場合は、「 [SSO を使用して Microsoft Graph に承認する」](authorize-to-microsoft-graph.md)を参照してください。

Microsoft ID プラットフォームを使用した認証のプログラミングについては、[Microsoft ID プラットフォームドキュメントを参照してください](/azure/active-directory/develop)。 そのドキュメント セットには、チュートリアルとガイドのほか、関連するサンプルへのリンクがあります。 もう一度、作業ウィンドウとは別のプロセスで実行される Office ダイアログ ボックスを考慮して、Office ダイアログ ボックスで実行するようにサンプルのコードを調整する必要がある場合があります。

コードが Microsoft Graph へのアクセス トークンを取得した後、ダイアログ ボックスから作業ウィンドウにアクセス トークンを渡すか、データベースにトークンを格納し、トークンが使用可能であることを作業ウィンドウに通知します。 (詳細については、「 [Office ダイアログ API を使用した認証](auth-with-office-dialog-api.md) 」を参照してください)。作業ウィンドウのコードは、Microsoft Graph からデータを要求し、それらの要求にトークンを含めます。 Microsoft Graph と Microsoft Graph SDK の呼び出しの詳細については、 [Microsoft Graph のドキュメントを参照してください](/graph/)。

## <a name="recommended-libraries-and-samples"></a>推奨されるライブラリおよびサンプル

Microsoft Graph にアクセスするときは、次のライブラリを使用することをお勧めします。

- .NET ベースのフレームワーク (.NET Core や ASP.NET など) のサーバー側を使用するアドインの場合は、[MSAL.NET](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki#conceptual-documentation) を使用します。
- NodeJS ベースのサーバー側を使用するアドインの場合は、[Passport Azure AD](https://github.com/AzureAD/passport-azure-ad) を使用します。
- 暗黙的なフローを使用するアドインの場合は、[msal.js](https://github.com/AzureAD/microsoft-authentication-library-for-js/wiki) を使用します。

Microsoft ID プラットフォーム (以前は AAD v.2.0) を使用するための推奨ライブラリの詳細については、「[Microsoft ID プラットフォームの認証ライブラリ](/azure/active-directory/develop/reference-v2-libraries)」を参照してください。

次のサンプルでは、Office アドインから Microsoft Graph データを取得します。

- [Office アドイン Microsoft Graph ASP.NET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)
- [Outlook アドイン Microsoft Graph ASP.NET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)
- [Office アドイン Microsoft Graph React](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React)
