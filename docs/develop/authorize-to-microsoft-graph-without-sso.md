---
title: アドインから Microsoft GraphにOffice承認する
description: アドインから Microsoft GraphにOfficeする方法について説明します。
ms.date: 01/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 8166b7a71767abd0456662dbe8573f59bb2c7e82
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743584"
---
# <a name="authorize-to-microsoft-graph-from-an-office-add-in"></a>アドインから Microsoft GraphにOffice承認する

アドインは、Microsoft Graphへのアクセス トークンを取得することで、Microsoft Graphデータに対する承認をMicrosoft ID プラットフォーム。 他の Web アプリケーションと同様に、承認コード フローまたは暗黙的フローを使用しますが、1 つの例外を除きます。Microsoft ID プラットフォーム では、サインイン ページを iframe で開くことができません。 Office アドインが *Office on the web* で実行されている場合、作業ウィンドウとして iFrame が使用されます。 つまり、ダイアログ API を使用してダイアログ ボックスでサインイン ページを開くOffice必要があります。 これは、認証と承認ヘルパー ライブラリの使用方法に影響します。 詳細については、「[Office ダイアログ API を使用して認証および承認する](auth-with-office-dialog-api.md)」を参照してください。

> [!NOTE]
> SSO を実装し、Microsoft サービスにアクセスする計画がある場合Graph SSO を使用して Microsoft にGraph[を参照してください](authorize-to-microsoft-graph.md)。

アプリケーションを使用したプログラミング認証の詳細については、Microsoft ID プラットフォームドキュメント[Microsoft ID プラットフォームしてください](/azure/active-directory/develop)。 このドキュメント セットには、チュートリアルとガイド、および関連するサンプルへのリンクがあります。 もう一度、Office ダイアログ ボックスで実行するサンプル内のコードを調整して、作業ウィンドウとは別のプロセスで実行される Office ダイアログ ボックスを考慮する必要があります。

コードが Microsoft Graph へのアクセス トークンを取得した後、アクセス トークンをダイアログ ボックスから作業ウィンドウに渡すか、またはトークンをデータベースに格納し、トークンが使用可能な作業ウィンドウにシグナルを送信します。 (詳細[については、「Office API を使用した認証」](auth-with-office-dialog-api.md)を参照してください。作業ウィンドウ内のコードは、Microsoft Graphデータを要求し、それらの要求にトークンを含む。 Microsoft Graph SDK と Microsoft Graph呼び出しの詳細については、「[Microsoft Graph」を参照してください](/graph/)。

## <a name="recommended-libraries-and-samples"></a>推奨されるライブラリおよびサンプル

Microsoft サーバーにアクセスする場合は、次のライブラリを使用することをおGraph。

- .NET ベースのフレームワーク (.NET Core や ASP.NET など) のサーバー側を使用するアドインの場合は、[MSAL.NET](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet/wiki#conceptual-documentation) を使用します。
- NodeJS ベースのサーバー側を使用するアドインの場合は、[Passport Azure AD](https://github.com/AzureAD/passport-azure-ad) を使用します。
- 暗黙的なフローを使用するアドインの場合は、[msal.js](https://github.com/AzureAD/microsoft-authentication-library-for-js/wiki) を使用します。

Microsoft ID プラットフォーム (以前は AAD v.2.0) を使用するための推奨ライブラリの詳細については、「[Microsoft ID プラットフォームの認証ライブラリ](/azure/active-directory/develop/reference-v2-libraries)」を参照してください。

次のサンプルでは、Microsoft GraphアドインからOfficeデータを取得します。

- [Office アドイン Microsoft Graph ASP.NET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)
- [Outlook アドイン Microsoft Graph ASP.NET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)
- [Office アドイン Microsoft Graph React](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React)
