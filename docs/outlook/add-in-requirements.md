---
title: Outlook アドインの要件
description: Outlook アドインが正しく読み込まれて機能するためには、サーバーとクライアントの両方に関していくつかの要件があります。
ms.date: 02/09/2021
ms.localizationpriority: high
ms.openlocfilehash: 33e24c731b78f302b6fe655c15ddc683343afe88
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154246"
---
# <a name="outlook-add-in-requirements"></a>Outlook アドインの要件

Outlook アドインが正しく読み込まれて機能するためには、サーバーとクライアントの両方に関していくつかの要件があります。

## <a name="client-requirements"></a>クライアント要件

- クライアントは、Outlook アドインをサポートするアプリケーションのいずれかでなければなりません。以下のクライアントがアドインをサポートしています。

  - Windows 用 Outlook 2013 以降
  - Mac 用 Outlook 2016 以降
  - Outlook on iOS
  - Outlook on Android
  - Outlook on the web (Exchange 2016 以降用)
  - Exchange 2013 向け Outlook on the web
  - Outlook.com

- クライアントは、直接接続を使用して Exchange サーバーまたは Microsoft 365 に接続する必要があります。ユーザーはクライアントを構成するときに、アカウントの種類として **Exchange**、**Office**、**Outlook.com** のいずれかを選択する必要があります。POP3 または IMAP を使用して接続するようにクライアントが構成されている場合、アドインは読み込まれません。

## <a name="mail-server-requirements"></a>メール サーバーの要件

ユーザーが Microsoft 365 または Outlook.com に接続している場合は、既にメール サーバーの要件をすべて満たしています。ただし、オンプレミスの Exchange Server インストール環境に接続しているユーザーの場合は、以下の要件が適用されます。

- サーバーは、Exchange 2013 以降である必要があります。
- Exchange Web サービス (EWS) が有効で、インターネットに公開されている必要があります。 多くのアドインでは、EWS が正しく機能する必要があります。
- 有効な ID トークンをサーバーが発行するためには、有効な認証証明書がサーバーになければなりません。 新しくインストールした Exchange Server には、既定の認証証明書が含まれます。 詳細については、「[Exchange 2016 のデジタル証明書と暗号化](/Exchange/architecture/client-access/certificates)」と「[Set-AuthConfig](/powershell/module/exchange/organization/Set-AuthConfig)」を参照してください。
- [AppSource](https://appsource.microsoft.com/marketplace/apps?product=office&page=1&src=office&corrid=a35323d5-0e3d-4cc0-ba44-57537d74aae8&omexanonuid=581941df-1c6f-4eda-89e7-651af8aeaeb2) からアドインにアクセスするためには、クライアント アクセス サーバーが AppSource と通信できなければなりません。

## <a name="add-in-server-requirements"></a>アドイン サーバーの要件

アドイン ファイル (HTML、JavaScript など) は、目的の Web サーバー プラットフォームでホストできます。唯一の要件は、HTTPS を使用するようにサーバーが構成されていなければならないこと、および SSL 証明書がクライアントで信頼されなければならないことです。

## <a name="see-also"></a>関連項目

- [Office アドインを実行するための要件](../concepts/requirements-for-running-office-add-ins.md)
- [Office アドインの Office クライアント アプリケーションとプラットフォームの可用性 (Outlook セクション)](../overview/office-add-in-availability.md#outlook)
- [Outlook JavaScript API の要件セットのサポート](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)
