---
title: アドインで ID トークンを使用してユーザーを認証する
description: サービスで SSO を実装するために、Outlook アドインが提供する ID トークンの使用方法について説明します。
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 839c3db2658ac4f16b965ee399df391ddcfd0385
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348539"
---
# <a name="authenticate-a-user-with-an-identity-token-for-exchange"></a><span data-ttu-id="6bbb4-103">Exchange の ID トークンを使用してユーザーを認証する</span><span class="sxs-lookup"><span data-stu-id="6bbb4-103">Authenticate a user with an identity token for Exchange</span></span>

<span data-ttu-id="6bbb4-104">Exchange のユーザー ID トークンは、アドインがアドイン ユーザーを一意に識別する方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="6bbb4-104">Exchange user identity tokens provide a way for your add-in to uniquely identify an add-in user.</span></span> <span data-ttu-id="6bbb4-105">ユーザーの ID を確立することで、Outlook アドインを使用している顧客がサインインせずにサービスに接続できる、シングル サインオン (SSO) 認証スキームをバック エンド サービスに実装できます。</span><span class="sxs-lookup"><span data-stu-id="6bbb4-105">By establishing the user's identity, you can implement a single sign-on (SSO) authentication scheme for your back-end service that enables customers who are using Outlook add-ins to connect to your service without signing in.</span></span> <span data-ttu-id="6bbb4-106">このトークンの種類を使用する場合の詳細については、「[Exchange のユーザー ID トークン](authentication.md#exchange-user-identity-token)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6bbb4-106">See [Exchange user identity token](authentication.md#exchange-user-identity-token) for more about when to use this token type.</span></span> <span data-ttu-id="6bbb4-107">この記事では、Exchange の ID トークンを使用してバックエンドにユーザーを認証する簡単な方法について説明します。</span><span class="sxs-lookup"><span data-stu-id="6bbb4-107">In this article, we'll take a look at a simplistic method of using the Exchange identity token to authenticate a user to your back-end.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="6bbb4-108">これは、単なる SSO の簡単な実装例です。</span><span class="sxs-lookup"><span data-stu-id="6bbb4-108">This is just a simple example of an SSO implementation.</span></span> <span data-ttu-id="6bbb4-109">従来どおり、ID と認証を処理するときは、コードが組織のセキュリティ要件を満たしていることを確認する必要があります。</span><span class="sxs-lookup"><span data-stu-id="6bbb4-109">As always, when you're dealing with identity and authentication, you have to make sure that your code meets the security requirements of your organization.</span></span>

## <a name="send-the-id-token-with-each-request"></a><span data-ttu-id="6bbb4-110">要求ごとに ID トークンを送信する</span><span class="sxs-lookup"><span data-stu-id="6bbb4-110">Send the ID token with each request</span></span>

<span data-ttu-id="6bbb4-111">最初の手順では、[getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) を呼び出すことにより、アドインでサーバーから Exchange のユーザー ID トークンを取得します。</span><span class="sxs-lookup"><span data-stu-id="6bbb4-111">The first step is for your add-in to obtain the Exchange user identity token from the server by calling [getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods).</span></span> <span data-ttu-id="6bbb4-112">その次に、アドインはこのトークンを、バックエンドに対する各要求とともに送信します。</span><span class="sxs-lookup"><span data-stu-id="6bbb4-112">Then the add-in sends this token with every request it makes to your back-end.</span></span> <span data-ttu-id="6bbb4-113">これはヘッダーか、要求の本文の一部として組み込まれます。</span><span class="sxs-lookup"><span data-stu-id="6bbb4-113">This could be in a header, or as part of the request body.</span></span>

## <a name="validate-the-token"></a><span data-ttu-id="6bbb4-114">トークンを検証する</span><span class="sxs-lookup"><span data-stu-id="6bbb4-114">Validate the token</span></span>

<span data-ttu-id="6bbb4-115">バックエンドは、トークンを検証してから承諾する必要があります。</span><span class="sxs-lookup"><span data-stu-id="6bbb4-115">The back-end MUST validate the token before accepting it.</span></span> <span data-ttu-id="6bbb4-116">これは、トークンがユーザーの Exchange サーバーによって発行されたことを確認する重要な手順です。</span><span class="sxs-lookup"><span data-stu-id="6bbb4-116">This is an important step to ensure that the token was issued by the user's Exchange server.</span></span> <span data-ttu-id="6bbb4-117">Exchange のユーザー ID トークンの検証の詳細については、「[Exchange の ID トークンを検証する](validate-an-identity-token.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="6bbb4-117">For information on validating Exchange user identity tokens, see [Validate an Exchange identity token](validate-an-identity-token.md).</span></span>

<span data-ttu-id="6bbb4-118">検証およびデコードが完了すると、トークンのペイロードは次のようになります。</span><span class="sxs-lookup"><span data-stu-id="6bbb4-118">Once validated and decoded, the payload of the token looks something like the following:</span></span>

```json
{ 
    "aud" : "https://mailhost.contoso.com/IdentityTest.html",
    "iss" : "00000002-0000-0ff1-ce00-000000000000@mailhost.contoso.com",
    "nbf" : "1505749527",
    "exp" : "1505778327",
    "appctxsender":"00000002-0000-0ff1-ce00-000000000000@mailhost.context.com",
    "isbrowserhostedapp":"true",
    "appctx" : {
        "msexchuid" : "53e925fa-76ba-45e1-be0f-4ef08b59d389",
        "version" : "ExIdTok.V1",
        "amurl" : "https://mailhost.contoso.com:443/autodiscover/metadata/json/1"
    }
}
```

## <a name="map-the-token-to-a-user-in-your-backend"></a><span data-ttu-id="6bbb4-119">トークンをバックエンドのユーザーにマップする</span><span class="sxs-lookup"><span data-stu-id="6bbb4-119">Map the token to a user in your backend</span></span>

<span data-ttu-id="6bbb4-120">バックエンド サービスはトークンから一意のユーザー ID を計算し、内部ユーザー システムのユーザーにマップできます。</span><span class="sxs-lookup"><span data-stu-id="6bbb4-120">Your back-end service can calculate a unique user ID from the token and map it to a user in your internal user system.</span></span> <span data-ttu-id="6bbb4-121">たとえば、ユーザーの格納にデータベースを使用する場合は、この一意の ID をデータベース内のユーザーのレコードに追加できます。</span><span class="sxs-lookup"><span data-stu-id="6bbb4-121">For example, if you use a database to store users, you could add this unique ID to the user's record in your database.</span></span>

### <a name="generate-a-unique-id"></a><span data-ttu-id="6bbb4-122">一意の ID を生成する</span><span class="sxs-lookup"><span data-stu-id="6bbb4-122">Generate a unique ID</span></span>

<span data-ttu-id="6bbb4-123">`msexchuid` プロパティと `amurl` プロパティを組み合わせて使用することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="6bbb4-123">We recommend that you use a combination of the `msexchuid` and `amurl` properties.</span></span> <span data-ttu-id="6bbb4-124">たとえば、2 つの値を連結して、Base64 でエンコードされた文字列を生成します。</span><span class="sxs-lookup"><span data-stu-id="6bbb4-124">For example, you could concatenate the two values together and generate a base 64-encoded string.</span></span> <span data-ttu-id="6bbb4-125">この値は毎回トークンから確実に生成できるので、Exchange のユーザー ID トークンをシステム内のユーザーにマップできます。</span><span class="sxs-lookup"><span data-stu-id="6bbb4-125">This value can be reliably generated from the token every time, so you can map an Exchange user identity token back to the user in your system.</span></span>

### <a name="check-the-user"></a><span data-ttu-id="6bbb4-126">ユーザーを確認する</span><span class="sxs-lookup"><span data-stu-id="6bbb4-126">Check the user</span></span>

<span data-ttu-id="6bbb4-127">次の手順では、生成された一意の ID を使用して、関連付けられた ID でシステム内のユーザーを確認します。</span><span class="sxs-lookup"><span data-stu-id="6bbb4-127">With the unique ID generated, the next step is to check for a user in your system with that associated ID.</span></span>

- <span data-ttu-id="6bbb4-128">ユーザーが見つかった場合、バックエンドは要求を認証済みとして処理し、要求の続行を許可します。</span><span class="sxs-lookup"><span data-stu-id="6bbb4-128">If the user is found, the back-end treats the request as authenticated, and allows the request to proceed.</span></span>

- <span data-ttu-id="6bbb4-129">ユーザーが見つからない場合、バックエンドはユーザーがサインインする必要があることを示すエラーを返します。</span><span class="sxs-lookup"><span data-stu-id="6bbb4-129">If the user is not found, then the back-end returns an error indicating that the user needs to sign in.</span></span> <span data-ttu-id="6bbb4-130">その後アドインは、既存の認証方法を使用してバックエンドにサインインするように求めるダイアログを表示します。</span><span class="sxs-lookup"><span data-stu-id="6bbb4-130">The add-in then prompts the user to sign in to the back-end using your existing authentication method.</span></span> <span data-ttu-id="6bbb4-131">ユーザーが認証されると、Exchange のユーザー ID トークンとユーザー認証の詳細が送信されます。</span><span class="sxs-lookup"><span data-stu-id="6bbb4-131">Once the user is authenticated, the Exchange user identity token is submitted with the user authentication details.</span></span> <span data-ttu-id="6bbb4-132">バックエンドはシステム内のユーザーのレコードを一意の ID で更新できます。</span><span class="sxs-lookup"><span data-stu-id="6bbb4-132">The back-end can then update the user's record in your system with the unique ID.</span></span>
