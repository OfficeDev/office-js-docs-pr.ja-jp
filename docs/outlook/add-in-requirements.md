---
title: Outlook アドインの要件
description: Outlook アドインが正しく読み込まれて機能するためには、サーバーとクライアントの両方に関していくつかの要件があります。
ms.date: 02/09/2021
localization_priority: Priority
ms.openlocfilehash: dd7831ce8ebd1165f920fe24775f46cd8cd7f91c
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234297"
---
# <a name="outlook-add-in-requirements"></a><span data-ttu-id="fa63a-103">Outlook アドインの要件</span><span class="sxs-lookup"><span data-stu-id="fa63a-103">Outlook add-in requirements</span></span>

<span data-ttu-id="fa63a-104">Outlook アドインが正しく読み込まれて機能するためには、サーバーとクライアントの両方に関していくつかの要件があります。</span><span class="sxs-lookup"><span data-stu-id="fa63a-104">For Outlook add-ins to load and function properly, there are a number of requirements for both the servers and the clients.</span></span>

## <a name="client-requirements"></a><span data-ttu-id="fa63a-105">クライアント要件</span><span class="sxs-lookup"><span data-stu-id="fa63a-105">Client requirements</span></span>

- <span data-ttu-id="fa63a-106">クライアントは、Outlook アドインをサポートするアプリケーションのいずれかでなければなりません。以下のクライアントがアドインをサポートしています:</span><span class="sxs-lookup"><span data-stu-id="fa63a-106">The client must be one of the supported applications for Outlook add-ins. The following clients support add-ins:</span></span>

   - <span data-ttu-id="fa63a-107">Windows 用 Outlook 2013 以降</span><span class="sxs-lookup"><span data-stu-id="fa63a-107">Outlook 2013 or later on Windows</span></span>
   - <span data-ttu-id="fa63a-108">Mac 用 Outlook 2016 以降</span><span class="sxs-lookup"><span data-stu-id="fa63a-108">Outlook 2016 or later on Mac</span></span>
   - <span data-ttu-id="fa63a-109">Outlook on iOS</span><span class="sxs-lookup"><span data-stu-id="fa63a-109">Outlook on iOS</span></span>
   - <span data-ttu-id="fa63a-110">Outlook on Android</span><span class="sxs-lookup"><span data-stu-id="fa63a-110">Outlook on Android</span></span>
   - <span data-ttu-id="fa63a-111">Outlook on the web (Exchange 2016 以降用)</span><span class="sxs-lookup"><span data-stu-id="fa63a-111">Outlook on the web for Exchange 2016 or later</span></span>
   - <span data-ttu-id="fa63a-112">Exchange 2013 向け Outlook on the web</span><span class="sxs-lookup"><span data-stu-id="fa63a-112">Outlook on the web for Exchange 2013</span></span>
   - <span data-ttu-id="fa63a-113">Outlook.com</span><span class="sxs-lookup"><span data-stu-id="fa63a-113">Outlook.com</span></span>

- <span data-ttu-id="fa63a-p101">クライアントは、直接接続を使用して Exchange サーバーまたは Microsoft 365 に接続する必要があります。ユーザーはクライアントを構成するときに、アカウントの種類として **Exchange**、**Office**、**Outlook.com** のいずれかを選択する必要があります。POP3 または IMAP を使用して接続するようにクライアントが構成されている場合、アドインは読み込まれません。</span><span class="sxs-lookup"><span data-stu-id="fa63a-p101">The client must be connected to an Exchange server or Microsoft 365 using a direct connection. When configuring the client, the user must choose an **Exchange**, **Office**, or **Outlook.com** account type. If the client is configured to connect with POP3 or IMAP, add-ins will not load.</span></span>

## <a name="mail-server-requirements"></a><span data-ttu-id="fa63a-117">メール サーバーの要件</span><span class="sxs-lookup"><span data-stu-id="fa63a-117">Mail server requirements</span></span>

<span data-ttu-id="fa63a-p102">ユーザーが Microsoft 365 または Outlook.com に接続している場合は、既にメール サーバーの要件をすべて満たしています。ただし、オンプレミスの Exchange Server インストール環境に接続しているユーザーの場合は、以下の要件が適用されます。</span><span class="sxs-lookup"><span data-stu-id="fa63a-p102">If the user is connected to Microsoft 365 or Outlook.com, mail server requirements are all taken care of already. However, for users connected to on-premises installations of Exchange Server, the following requirements apply.</span></span>

- <span data-ttu-id="fa63a-120">サーバーは、Exchange 2013 以降である必要があります。</span><span class="sxs-lookup"><span data-stu-id="fa63a-120">The server must be Exchange 2013 or later.</span></span>
- <span data-ttu-id="fa63a-121">Exchange Web サービス (EWS) が有効で、インターネットに公開されている必要があります。</span><span class="sxs-lookup"><span data-stu-id="fa63a-121">Exchange Web Services (EWS) must be enabled and must be exposed to the Internet.</span></span> <span data-ttu-id="fa63a-122">多くのアドインでは、EWS が正しく機能する必要があります。</span><span class="sxs-lookup"><span data-stu-id="fa63a-122">Many add-ins require EWS to function properly.</span></span>
- <span data-ttu-id="fa63a-123">有効な ID トークンをサーバーが発行するためには、有効な認証証明書がサーバーになければなりません。</span><span class="sxs-lookup"><span data-stu-id="fa63a-123">The server must have a valid authentication certificate in order for the server to issue valid identity tokens.</span></span> <span data-ttu-id="fa63a-124">新しくインストールした Exchange Server には、既定の認証証明書が含まれます。</span><span class="sxs-lookup"><span data-stu-id="fa63a-124">New installations of Exchange Server include a default authentication certificate.</span></span> <span data-ttu-id="fa63a-125">詳細については、「[Exchange 2016 のデジタル証明書と暗号化](/Exchange/architecture/client-access/certificates)」と「[Set-AuthConfig](/powershell/module/exchange/organization/Set-AuthConfig)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="fa63a-125">For more information, see [Digital certificates and encryption in Exchange 2016](/Exchange/architecture/client-access/certificates) and [Set-AuthConfig](/powershell/module/exchange/organization/Set-AuthConfig).</span></span>
- <span data-ttu-id="fa63a-126">[AppSource](https://appsource.microsoft.com/marketplace/apps?product=office&page=1&src=office&corrid=a35323d5-0e3d-4cc0-ba44-57537d74aae8&omexanonuid=581941df-1c6f-4eda-89e7-651af8aeaeb2) からアドインにアクセスするためには、クライアント アクセス サーバーが AppSource と通信できなければなりません。</span><span class="sxs-lookup"><span data-stu-id="fa63a-126">To access add-ins from [AppSource](https://appsource.microsoft.com/marketplace/apps?product=office&page=1&src=office&corrid=a35323d5-0e3d-4cc0-ba44-57537d74aae8&omexanonuid=581941df-1c6f-4eda-89e7-651af8aeaeb2), the client access servers must be able to communicate with AppSource.</span></span>

## <a name="add-in-server-requirements"></a><span data-ttu-id="fa63a-127">アドイン サーバーの要件</span><span class="sxs-lookup"><span data-stu-id="fa63a-127">Add-in server requirements</span></span>

<span data-ttu-id="fa63a-p105">アドイン ファイル (HTML、JavaScript など) は、目的の Web サーバー プラットフォームでホストできます。唯一の要件は、HTTPS を使用するようにサーバーが構成されていなければならないこと、および SSL 証明書がクライアントで信頼されなければならないことです。</span><span class="sxs-lookup"><span data-stu-id="fa63a-p105">Add-in files (HTML, JavaScript, etc.) can be hosted on any web server platform desired. The only requirement is that the server must be configured to use HTTPS, and the SSL certificate must be trusted by the client.</span></span>

## <a name="see-also"></a><span data-ttu-id="fa63a-130">関連項目</span><span class="sxs-lookup"><span data-stu-id="fa63a-130">See also</span></span>

- [<span data-ttu-id="fa63a-131">Office アドインを実行するための要件</span><span class="sxs-lookup"><span data-stu-id="fa63a-131">Requirements for running Office Add-ins</span></span>](../concepts/requirements-for-running-office-add-ins.md)
- [<span data-ttu-id="fa63a-132">Office アドインの Office クライアント アプリケーションとプラットフォームの可用性 (Outlook セクション)</span><span class="sxs-lookup"><span data-stu-id="fa63a-132">Office client application and platform availability for Office Add-ins (Outlook section)</span></span>](../overview/office-add-in-availability.md#outlook)
- [<span data-ttu-id="fa63a-133">Outlook JavaScript API の要件セットのサポート</span><span class="sxs-lookup"><span data-stu-id="fa63a-133">Outlook JavaScript API requirement set support</span></span>](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)
