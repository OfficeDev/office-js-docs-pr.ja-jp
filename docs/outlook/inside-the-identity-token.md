---
title: Outlook アドインでの Exchange の ID トークンの内部
description: Outlook アドインから生成される Exchange のユーザー ID トークンの内容について説明します。
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: dee8416660386c25a55caa42b6e5ee8685ee8852
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609091"
---
# <a name="inside-the-exchange-identity-token"></a><span data-ttu-id="3a141-103">Exchange の ID トークンの内部</span><span class="sxs-lookup"><span data-stu-id="3a141-103">Inside the Exchange identity token</span></span>

<span data-ttu-id="3a141-104">[getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) メソッドによって返された Exchange のユーザー ID トークンは、アドイン コードがバックエンド サービスへの呼び出しでユーザー ID を含めるための方法を提供します。</span><span class="sxs-lookup"><span data-stu-id="3a141-104">The Exchange user identity token returned by the [getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method provides a way for your add-in code to include the user's identity with calls to your back-end service.</span></span> <span data-ttu-id="3a141-105">この記事では、トークンの形式と内容について説明します。</span><span class="sxs-lookup"><span data-stu-id="3a141-105">This article will discuss the format and contents of the token.</span></span>

<span data-ttu-id="3a141-106">Exchange ユーザー ID トークンとは、そのトークンを送信する Exchange サーバーによって自己署名された、Base 64 URL 形式でエンコードされた文字列です。</span><span class="sxs-lookup"><span data-stu-id="3a141-106">An Exchange user identity token is a base-64 URL-encoded string that is signed by the Exchange server that sent it.</span></span> <span data-ttu-id="3a141-107">トークンは暗号化されていません。署名の検証に使用する公開キーは、トークンを発行した Exchange サーバーに保存されています。</span><span class="sxs-lookup"><span data-stu-id="3a141-107">The token is not encrypted, and the public key that you use to validate the signature is stored on the Exchange server that issued the token.</span></span> <span data-ttu-id="3a141-108">トークンには 3 つのパーツ (ヘッダー、ペイロード、署名) があります。</span><span class="sxs-lookup"><span data-stu-id="3a141-108">The token has three parts: a header, a payload, and a signature.</span></span> <span data-ttu-id="3a141-109">トークン文字列では、トークンを容易に分割できるように、パーツがピリオド文字 (`.`) で区切られています。</span><span class="sxs-lookup"><span data-stu-id="3a141-109">In the token string, the parts are separated by a period character (`.`) to make it easy for you to split the token.</span></span>

<span data-ttu-id="3a141-110">Exchange では ID トークンに、JSON Web トークン (JWT) 形式を使用します。</span><span class="sxs-lookup"><span data-stu-id="3a141-110">Exchange uses a the JSON Web Token (JWT) format for the identity token.</span></span> <span data-ttu-id="3a141-111">JWT トークンの詳細については、「[RFC 7519 JSON Web Token (JWT)](https://www.rfc-editor.org/rfc/rfc7519.txt)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3a141-111">For information about JWT tokens, see [RFC 7519 JSON Web Token (JWT)](https://www.rfc-editor.org/rfc/rfc7519.txt).</span></span>

## <a name="identity-token-header"></a><span data-ttu-id="3a141-112">ID トークンのヘッダー</span><span class="sxs-lookup"><span data-stu-id="3a141-112">Identity token header</span></span>

<span data-ttu-id="3a141-113">ヘッダーは、トークンの形式に関する情報と、署名情報に関する情報を提供します。</span><span class="sxs-lookup"><span data-stu-id="3a141-113">The header provides information about the format and signature information of the token.</span></span> <span data-ttu-id="3a141-114">次の例は、トークンのヘッダーの外観を示しています。</span><span class="sxs-lookup"><span data-stu-id="3a141-114">The following example shows what the header of the token looks like.</span></span>

```JSON
{
  "typ": "JWT",
  "alg": "RS256",
  "x5t": "Un6V7lYN-rMgaCoFSTO5z707X-4"
}
```

<br/>
 
<span data-ttu-id="3a141-115">トークンのヘッダーのパーツについての説明を、次の表に示します。</span><span class="sxs-lookup"><span data-stu-id="3a141-115">The following table describes the parts of the token header.</span></span>

| <span data-ttu-id="3a141-116">クレーム</span><span class="sxs-lookup"><span data-stu-id="3a141-116">Claim</span></span> | <span data-ttu-id="3a141-117">値</span><span class="sxs-lookup"><span data-stu-id="3a141-117">Value</span></span> | <span data-ttu-id="3a141-118">説明</span><span class="sxs-lookup"><span data-stu-id="3a141-118">Description</span></span> |
|:-----|:-----|:-----|
| `typ` | `JWT` | <span data-ttu-id="3a141-119">トークンを JSON Web トークンとして識別します。</span><span class="sxs-lookup"><span data-stu-id="3a141-119">Identifies the token as a JSON Web Token.</span></span> <span data-ttu-id="3a141-120">Exchange サーバーから提供される ID トークンは、すべて JWT トークンです。</span><span class="sxs-lookup"><span data-stu-id="3a141-120">All identity tokens provided by Exchange server are JWT tokens.</span></span> |
| `alg` | `RS256` | <span data-ttu-id="3a141-121">署名の作成に使用されるハッシュ アルゴリズム。</span><span class="sxs-lookup"><span data-stu-id="3a141-121">The hashing algorithm that is used to create the signature.</span></span> <span data-ttu-id="3a141-122">Exchange サーバーから提供されるトークンは、すべて SHA-256 ハッシュ アルゴリズムの RSASSA-PKCS1-v1_5 を使用します。</span><span class="sxs-lookup"><span data-stu-id="3a141-122">All tokens provided by Exchange server use the RSASSA-PKCS1-v1_5 with SHA-256 hash algorithm.</span></span> |
| `x5t` | <span data-ttu-id="3a141-123">証明書の拇印</span><span class="sxs-lookup"><span data-stu-id="3a141-123">Certificate thumbprint</span></span> | <span data-ttu-id="3a141-124">トークンの X.509 拇印です。</span><span class="sxs-lookup"><span data-stu-id="3a141-124">The X.509 thumbprint of the token.</span></span> |

## <a name="identity-token-payload"></a><span data-ttu-id="3a141-125">ID トークンのペイロード</span><span class="sxs-lookup"><span data-stu-id="3a141-125">Identity token payload</span></span>

<span data-ttu-id="3a141-p107">ペイロードには、電子メール アカウントの識別と、トークンを送信した Exchange サーバーの識別を行う認証クレームが含まれます。以下に、ペイロード セクションの例を示します。</span><span class="sxs-lookup"><span data-stu-id="3a141-p107">The payload contains the authentication claims that identify the email account and identify the Exchange server that sent the token. The following example shows what the payload section looks like.</span></span>

```JSON
{ 
  "aud": "https://mailhost.contoso.com/IdentityTest.html", 
  "iss": "00000002-0000-0ff1-ce00-000000000000@mailhost.contoso.com", 
  "nbf": "1331579055", 
  "exp": "1331607855", 
  "appctxsender": "00000002-0000-0ff1-ce00-000000000000@mailhost.context.com",
  "isbrowserhostedapp": "true",
  "appctx": { 
    "msexchuid": "53e925fa-76ba-45e1-be0f-4ef08b59d389@mailhost.contoso.com",
    "version": "ExIdTok.V1",
    "amurl": "https://mailhost.contoso.com:443/autodiscover/metadata/json/1"
  } 
}
```

<br/>
 
<span data-ttu-id="3a141-128">ID トークンのペイロードのパーツを、次の表に示します。</span><span class="sxs-lookup"><span data-stu-id="3a141-128">The following table lists the parts of the identity token payload.</span></span>

| <span data-ttu-id="3a141-129">クレーム</span><span class="sxs-lookup"><span data-stu-id="3a141-129">Claim</span></span> | <span data-ttu-id="3a141-130">説明</span><span class="sxs-lookup"><span data-stu-id="3a141-130">Description</span></span> |
|:-----|:-----|
| `aud` | <span data-ttu-id="3a141-131">トークンを要求したアドインの URL。</span><span class="sxs-lookup"><span data-stu-id="3a141-131">The URL of the add-in that requested the token.</span></span> <span data-ttu-id="3a141-132">トークンは、クライアントのブラウザー内で実行されているアドインから送信された場合にのみ有効です。</span><span class="sxs-lookup"><span data-stu-id="3a141-132">A token is only valid if it is sent from the add-in that is running in the client's browser.</span></span> <span data-ttu-id="3a141-133">アドインが Office アドイン マニフェスト スキーマ v1.1 を使用する場合、この URL はフォーム タイプ `ItemRead` または `ItemEdit` (アドイン マニフェスト内で [FormSettings](../reference/manifest/formsettings.md) 要素の一部として最初に出現する方) の下にある、最初の `SourceLocation` 要素に指定された URL になります。</span><span class="sxs-lookup"><span data-stu-id="3a141-133">If the add-in uses the Office Add-ins manifests schema v1.1, this URL is the URL specified in the first `SourceLocation` element, under the form type `ItemRead` or `ItemEdit`, whichever occurs first as part of the [FormSettings](../reference/manifest/formsettings.md) element in the add-in manifest.</span></span> |
| `iss` | <span data-ttu-id="3a141-p109">トークンを発行した Exchange サーバーの一意の識別子です。この Exchange サーバーから発行されるトークンはすべて同じ識別子になります。</span><span class="sxs-lookup"><span data-stu-id="3a141-p109">A unique identifier for the Exchange server that issued the token. All tokens issued by this Exchange server will have the same identifier.</span></span> |
| `nbf` | <span data-ttu-id="3a141-p110">トークンの有効期間の開始日時です。この値は 1970 年 1 月 1 日を起点とする秒数です。</span><span class="sxs-lookup"><span data-stu-id="3a141-p110">The date and time that the token is valid starting from. The value is the number of seconds since January 1, 1970.</span></span> |
| `exp` | <span data-ttu-id="3a141-p111">トークンの有効期間の終了日時です。この値も 1970 年 1 月 1 日を起点とする秒数です。</span><span class="sxs-lookup"><span data-stu-id="3a141-p111">The date and time that the token is valid until. The value is the number of seconds since January 1, 1970.</span></span> |
| `appctxsender` | <span data-ttu-id="3a141-140">アプリケーション コンテキストを送信した Exchange サーバーの一意の識別子。</span><span class="sxs-lookup"><span data-stu-id="3a141-140">A unique identifier for the Exchange server that sent the application context.</span></span> |
| `isbrowserhostedapp` | <span data-ttu-id="3a141-141">アドインがブラウザーでホストされるかどうかを指定します。</span><span class="sxs-lookup"><span data-stu-id="3a141-141">Indicates whether the add-in is hosted in a browser.</span></span> |
| `appctx` | <span data-ttu-id="3a141-142">トークンのアプリケーション コンテキスト。</span><span class="sxs-lookup"><span data-stu-id="3a141-142">The application context for the token.</span></span> |

<span data-ttu-id="3a141-143">appctx クレーム内の情報は、アカウントの一意の識別子と、トークンの署名に使用された公開キーの場所を提供します。</span><span class="sxs-lookup"><span data-stu-id="3a141-143">The information in the appctx claim provides you with the unique identifier for the account and the location of the public key used to sign the token.</span></span> <span data-ttu-id="3a141-144">`appctx` クレームのパーツを、次の表に示します。</span><span class="sxs-lookup"><span data-stu-id="3a141-144">The following table lists the parts of the `appctx` claim.</span></span>

| <span data-ttu-id="3a141-145">アプリケーション コンテキスト プロパティ</span><span class="sxs-lookup"><span data-stu-id="3a141-145">Application context property</span></span> | <span data-ttu-id="3a141-146">説明</span><span class="sxs-lookup"><span data-stu-id="3a141-146">Description</span></span> |
|:-----|:-----|
| `msexchuid` | <span data-ttu-id="3a141-147">電子メール アカウントと Exchange サーバーに割り当てられた一意の識別子。</span><span class="sxs-lookup"><span data-stu-id="3a141-147">A unique identifier associated with the email account and the Exchange server.</span></span> |
| `version` | <span data-ttu-id="3a141-148">トークンのバージョン番号。</span><span class="sxs-lookup"><span data-stu-id="3a141-148">The version number of the token.</span></span> <span data-ttu-id="3a141-149">Exchange によって提供されるトークンの値は、すべて `ExIdTok.V1` になります。</span><span class="sxs-lookup"><span data-stu-id="3a141-149">For all tokens provided by Exchange, the value is `ExIdTok.V1`.</span></span> |
| `amurl` | <span data-ttu-id="3a141-150">トークンに署名するために使用された X.509 証明書の公開キーが含まれる認証メタデータ ドキュメントの URL。</span><span class="sxs-lookup"><span data-stu-id="3a141-150">The URL of the authentication metadata document that contains the public key of the X.509 certificate that was used to sign the token.</span></span><br/><br/><span data-ttu-id="3a141-151">認証メタデータ ドキュメントの使用方法については、「[Exchange の ID トークンを検証する](validate-an-identity-token.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3a141-151">For more information about how to use the authentication metadata document, see [Validate an Exchange identity token](validate-an-identity-token.md).</span></span> |

## <a name="identity-token-signature"></a><span data-ttu-id="3a141-152">ID トークンの署名</span><span class="sxs-lookup"><span data-stu-id="3a141-152">Identity token signature</span></span>

<span data-ttu-id="3a141-p114">この署名は、ヘッダーおよびペイロード セクションに対して、ヘッダーで指定されたアルゴリズムを使用したハッシュ処理を行うと共に、ペイロード内で指定された場所にあるサーバー上の自己署名された X509 証明書を使用することで作成されます。Web サービスは、この署名を検証して、ID トークンがその送信元として想定されるサーバーから発行されたものであることを確認できます。</span><span class="sxs-lookup"><span data-stu-id="3a141-p114">The signature is created by hashing the header and payload sections with the algorithm specified in the header and using the self-signed X509 certificate located on the server at the location specified in the payload. Your web service can validate this signature to help make sure that the identity token comes from the server that you expect to send it.</span></span>

## <a name="see-also"></a><span data-ttu-id="3a141-155">関連項目</span><span class="sxs-lookup"><span data-stu-id="3a141-155">See also</span></span>

<span data-ttu-id="3a141-156">Exchange のユーザー ID トークンの解析例については、「[Outlook アドイン トークン ビューアー](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="3a141-156">For an example that parses the Exchange user identity token, see [Outlook-Add-In-Token-Viewer](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer).</span></span>
