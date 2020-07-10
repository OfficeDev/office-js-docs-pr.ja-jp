---
title: Outlook アドインの要件
description: Outlook アドインが正しく読み込まれて機能するためには、サーバーとクライアントの両方に関していくつかの要件があります。
ms.date: 07/07/2020
localization_priority: Priority
ms.openlocfilehash: 700e0efd2ab2655de61d37d42038fa2c15a99cb4
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093995"
---
# <a name="outlook-add-in-requirements"></a>Outlook アドインの要件

Outlook アドインが正しく読み込まれて機能するためには、サーバーとクライアントの両方に関していくつかの要件があります。

## <a name="client-requirements"></a>クライアント要件

- クライアントは、Outlook アドインをサポートするホストのいずれかでなければなりません。以下のクライアントがアドインをサポートしています。

   - Windows 用 Outlook 2013 以降
   - Mac 用 Outlook 2016 以降
   - Outlook on iOS
   - Outlook on Android
   - Outlook on the web (Exchange 2016 以降および Office 365 用)
   - Exchange 2013 向け Outlook on the web
   - Outlook.com

- The client must be connected to an Exchange server or Microsoft 365 using a direct connection. When configuring the client, the user must choose an **Exchange**, **Office 365**, or **Outlook.com** account type. If the client is configured to connect with POP3 or IMAP, add-ins will not load.

## <a name="mail-server-requirements"></a>メール サーバーの要件

If the user is connected to Microsoft 365 or Outlook.com, mail server requirements are all taken care of already. However, for users connected to on-premises installations of Exchange Server, the following requirements apply.

- サーバーは、Exchange 2013 以降である必要があります。
- Exchange Web サービス (EWS) が有効で、インターネットに公開されている必要があります。 多くのアドインでは、EWS が正しく機能する必要があります。
- 有効な ID トークンをサーバーが発行するためには、有効な認証証明書がサーバーになければなりません。 新しくインストールした Exchange Server には、既定の認証証明書が含まれます。 詳細については、「[Exchange 2016 のデジタル証明書と暗号化](/Exchange/architecture/client-access/certificates)」と「[Set-AuthConfig](/powershell/module/exchange/organization/Set-AuthConfig)」を参照してください。
- [AppSource](https://appsource.microsoft.com/marketplace/apps?product=office&page=1&src=office&corrid=a35323d5-0e3d-4cc0-ba44-57537d74aae8&omexanonuid=581941df-1c6f-4eda-89e7-651af8aeaeb2) からアドインにアクセスするためには、クライアント アクセス サーバーが AppSource と通信できなければなりません。

## <a name="add-in-server-requirements"></a>アドイン サーバーの要件

Add-in files (HTML, JavaScript, etc.) can be hosted on any web server platform desired. The only requirement is that the server must be configured to use HTTPS, and the SSL certificate must be trusted by the client.

## <a name="see-also"></a>関連項目

- [Office アドインを実行するための要件](../concepts/requirements-for-running-office-add-ins.md)
- [Office アドインのホストとプラットフォームの可用性 (Outlook セクション)](../overview/office-add-in-availability.md#outlook)
- [Outlook JavaScript API の要件セットのサポート](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)
