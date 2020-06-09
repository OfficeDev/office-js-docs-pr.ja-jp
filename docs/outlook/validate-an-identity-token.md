---
title: Outlook アドイン ID トークンを検証する
description: 使用している Outlook アドインから Exchange のユーザー ID トークンを送信できますが、要求を信頼する前に、トークンを検証して適切な Exchange サーバーからのものであることを確認する必要があります。
ms.date: 05/08/2020
localization_priority: Normal
ms.openlocfilehash: 89be659085dbf35b4ad6644eba3b5bf3acd24a9d
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 06/08/2020
ms.locfileid: "44604581"
---
# <a name="validate-an-exchange-identity-token"></a><span data-ttu-id="09151-103">Exchange の ID トークンを検証する</span><span class="sxs-lookup"><span data-stu-id="09151-103">Validate an Exchange identity token</span></span>

<span data-ttu-id="09151-104">使用している Outlook アドインから Exchange のユーザー ID トークンを送信できますが、要求を信頼する前に、トークンを検証して適切な Exchange サーバーからのものであることを確認する必要があります。</span><span class="sxs-lookup"><span data-stu-id="09151-104">Your Outlook add-in can send you an Exchange user identity token, but before you trust the request you must validate the token to ensure that it came from the Exchange server that you expect.</span></span> <span data-ttu-id="09151-105">Exchange のユーザー ID トークンは、JSON Web トークン (JWT) です。</span><span class="sxs-lookup"><span data-stu-id="09151-105">Exchange user identity tokens are JSON Web Tokens (JWT).</span></span> <span data-ttu-id="09151-106">JWT の検証に必要な手順は、「[RFC 7519 JSON Web Token (JWT)](https://www.rfc-editor.org/rfc/rfc7519.txt)」に記載されています。</span><span class="sxs-lookup"><span data-stu-id="09151-106">The steps required to validate a JWT are described in [RFC 7519 JSON Web Token (JWT)](https://www.rfc-editor.org/rfc/rfc7519.txt).</span></span>

<span data-ttu-id="09151-107">ID トークンの検証およびユーザーの一意識別子の取得は 4 つのステップで進めることをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="09151-107">We suggest that you use a four-step process to validate the identity token and obtain the user's unique identifier.</span></span> <span data-ttu-id="09151-108">まず、Base 64 URL 形式でエンコードされた文字列から、JSON Web トークン (JWT) を抽出します。</span><span class="sxs-lookup"><span data-stu-id="09151-108">First, extract the JSON Web Token (JWT) from a base64 URL-encoded string.</span></span> <span data-ttu-id="09151-109">次に、トークンが整形式であること、使用する Outlook アドイン向けのトークンであること、有効期限が切れていないこと、および認証メタデータ ドキュメントの有効な URL を抽出できることを確認します。</span><span class="sxs-lookup"><span data-stu-id="09151-109">Second, make sure that the token is well-formed, that it is for your Outlook add-in, that it has not expired, and that you can extract a valid URL for the authentication metadata document.</span></span> <span data-ttu-id="09151-110">その後、Exchange サーバーから認証メタデータ ドキュメントを取得し、ID トークンに添付されている署名を検証します。</span><span class="sxs-lookup"><span data-stu-id="09151-110">Next, retrieve the authentication metadata document from the Exchange server and validate the signature attached to the identity token.</span></span> <span data-ttu-id="09151-111">最後に、ユーザーの Exchange ID と認証メタデータドキュメントの URL を連結することによって、ユーザーの一意の識別子を計算します。</span><span class="sxs-lookup"><span data-stu-id="09151-111">Finally, compute a unique identifier for the user by concatenating the user's Exchange ID with the URL of the authentication metadata document.</span></span>

## <a name="extract-the-json-web-token"></a><span data-ttu-id="09151-112">JSON Web トークンを抽出する</span><span class="sxs-lookup"><span data-stu-id="09151-112">Extract the JSON Web Token</span></span>

<span data-ttu-id="09151-113">[getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) から返されたトークンは、トークンのエンコードされた文字列表現です。</span><span class="sxs-lookup"><span data-stu-id="09151-113">The token returned from [getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) is an encoded string representation of the token.</span></span> <span data-ttu-id="09151-114">この形式では、すべての JWT にピリオドで区切られた 3 つの部分があります (RFC 7519 ごと)。</span><span class="sxs-lookup"><span data-stu-id="09151-114">In this form, per RFC 7519, all JWTs have three parts, separated by a period.</span></span> <span data-ttu-id="09151-115">形式は次のようになります。</span><span class="sxs-lookup"><span data-stu-id="09151-115">The format is as follows.</span></span>

```json
{header}.{payload}.{signature}
```

<span data-ttu-id="09151-116">ヘッダーとペイロードは、各部分の JSON 表現を取得するために、Base64 でデコードされる必要があります。</span><span class="sxs-lookup"><span data-stu-id="09151-116">The header and payload should be base64-decoded to obtain a JSON representation of each part.</span></span> <span data-ttu-id="09151-117">署名は、バイナリ シグネチャを含むバイト配列を取得するために、Base64 でデコードされる必要があります。</span><span class="sxs-lookup"><span data-stu-id="09151-117">The signature should be base64-decoded to obtain a byte array containing the binary signature.</span></span>

<span data-ttu-id="09151-118">トークンの内容の詳細については、「[Exchange の ID トークンの内部](inside-the-identity-token.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="09151-118">For more information about the contents of the token, see [Inside the Exchange identity token](inside-the-identity-token.md).</span></span>

<span data-ttu-id="09151-119">デコードされた 3 つのコンポーネントがあれば、トークンの内容の検証を進めることができます。</span><span class="sxs-lookup"><span data-stu-id="09151-119">After you have the three decoded components, you can proceed with validating the content of the token.</span></span>

## <a name="validate-token-contents"></a><span data-ttu-id="09151-120">トークンの内容を検証する</span><span class="sxs-lookup"><span data-stu-id="09151-120">Validate token contents</span></span>

<span data-ttu-id="09151-121">トークンの内容を検証するには、以下を確認する必要があります。</span><span class="sxs-lookup"><span data-stu-id="09151-121">To validate the token contents, you should check the following.</span></span>

- <span data-ttu-id="09151-122">ヘッダーを確認し、次の点を確認します。</span><span class="sxs-lookup"><span data-stu-id="09151-122">Check the header and verify that the:</span></span>
    - <span data-ttu-id="09151-123">`typ`claim はに設定されて `JWT` います。</span><span class="sxs-lookup"><span data-stu-id="09151-123">`typ` claim is set to `JWT`.</span></span>
    - <span data-ttu-id="09151-124">`alg`claim はに設定されて `RS256` います。</span><span class="sxs-lookup"><span data-stu-id="09151-124">`alg` claim is set to `RS256`.</span></span>
    - <span data-ttu-id="09151-125">`x5t`claim が存在します。</span><span class="sxs-lookup"><span data-stu-id="09151-125">`x5t` claim is present.</span></span>

- <span data-ttu-id="09151-126">ペイロードを確認し、次の点を確認します。</span><span class="sxs-lookup"><span data-stu-id="09151-126">Check the payload and verify that the:</span></span>
    - <span data-ttu-id="09151-127">`amurl`内のクレーム `appctx` は、承認済みのトークン署名キーマニフェストファイルの場所に設定されます。</span><span class="sxs-lookup"><span data-stu-id="09151-127">`amurl` claim inside the `appctx` is set to the location of an authorized token signing key manifest file.</span></span> <span data-ttu-id="09151-128">たとえば、 `amurl` Office 365 に対して予想される値は https://outlook.office365.com:443/autodiscover/metadata/json/1 です。</span><span class="sxs-lookup"><span data-stu-id="09151-128">For example, the expected `amurl` value for Office 365 is https://outlook.office365.com:443/autodiscover/metadata/json/1.</span></span> <span data-ttu-id="09151-129">次のセクションを参照してください。詳細については、「[ドメイン」を](#verify-the-domain)参照してください。</span><span class="sxs-lookup"><span data-stu-id="09151-129">See the next section [Verify the domain](#verify-the-domain) for additional information.</span></span>
    - <span data-ttu-id="09151-130">現在の時刻は、およびクレームで指定された時間です `nbf` `exp` 。</span><span class="sxs-lookup"><span data-stu-id="09151-130">Current time is between the times specified in the `nbf` and `exp` claims.</span></span> <span data-ttu-id="09151-131">`nbf` クレームは、トークンが有効と考えられる最も早い時刻を指定し、`exp` クレームはトークンの有効期限を指定します。</span><span class="sxs-lookup"><span data-stu-id="09151-131">The `nbf` claim specifies the earliest time that the token is considered valid, and the `exp` claim specifies the expiration time for the token.</span></span> <span data-ttu-id="09151-132">サーバー間のクロック設定には、ある程度の変動を許可することをお勧めします。</span><span class="sxs-lookup"><span data-stu-id="09151-132">It is recommended to allow for some variation in clock settings between servers.</span></span>
    - <span data-ttu-id="09151-133">`aud`claim は、アドインに必要な URL です。</span><span class="sxs-lookup"><span data-stu-id="09151-133">`aud` claim is the expected URL for your add-in.</span></span>
    - <span data-ttu-id="09151-134">`version`クレーム内のクレーム `appctx` はに設定されてい `ExIdTok.V1` ます。</span><span class="sxs-lookup"><span data-stu-id="09151-134">`version` claim inside the `appctx` claim is set to `ExIdTok.V1`.</span></span>

### <a name="verify-the-domain"></a><span data-ttu-id="09151-135">ドメインを確認する</span><span class="sxs-lookup"><span data-stu-id="09151-135">Verify the domain</span></span>

<span data-ttu-id="09151-136">このセクションで前述した検証ロジックを実装する場合は、要求のドメインが `amurl` ユーザーの自動検出ドメインと一致することも要求する必要があります。</span><span class="sxs-lookup"><span data-stu-id="09151-136">When implementing the verification logic described previously in this section, you should also require that the domain of the `amurl` claim matches the Autodiscover domain for the user.</span></span> <span data-ttu-id="09151-137">これを行うには、自動検出を使用または実装する必要があります。</span><span class="sxs-lookup"><span data-stu-id="09151-137">To do so, you'll need to use or implement Autodiscover.</span></span> <span data-ttu-id="09151-138">詳細については、「 [Exchange の自動検出](/exchange/client-developer/exchange-web-services/autodiscover-for-exchange)を開始する」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="09151-138">To learn more, you can start with [Autodiscover for Exchange](/exchange/client-developer/exchange-web-services/autodiscover-for-exchange).</span></span>

## <a name="validate-the-identity-token-signature"></a><span data-ttu-id="09151-139">ID トークンの署名を検証する</span><span class="sxs-lookup"><span data-stu-id="09151-139">Validate the identity token signature</span></span>

<span data-ttu-id="09151-140">JWT に必要なクレームが含まれていることを確認したら、トークンの署名の検証を進めることができます。</span><span class="sxs-lookup"><span data-stu-id="09151-140">After you know that the JWT contains the required claims, you can proceed with validating the token signature.</span></span>

### <a name="retrieve-the-public-signing-key"></a><span data-ttu-id="09151-141">公開署名キーを取得する</span><span class="sxs-lookup"><span data-stu-id="09151-141">Retrieve the public signing key</span></span>

<span data-ttu-id="09151-142">最初のステップでは、Exchange サーバーがトークンの署名に使用した証明書に対応する公開キーを取得します。</span><span class="sxs-lookup"><span data-stu-id="09151-142">The first step is to retrieve the public key that corresponds to the certificate that the Exchange server used to sign the token.</span></span> <span data-ttu-id="09151-143">鍵は認証メタデータ ドキュメントに記載されています。</span><span class="sxs-lookup"><span data-stu-id="09151-143">The key is found in the authentication metadata document.</span></span> <span data-ttu-id="09151-144">このドキュメントは、`amurl` クレームで指定された URL でホストされている JSON ファイルです。</span><span class="sxs-lookup"><span data-stu-id="09151-144">This document is a JSON file hosted at the URL specified in the `amurl` claim.</span></span>

<span data-ttu-id="09151-145">認証メタデータ ドキュメントには、次の形式を使用します。</span><span class="sxs-lookup"><span data-stu-id="09151-145">The authentication metadata document uses the following format.</span></span>

```json
{
    "id": "_70b34511-d105-4e2b-9675-39f53305bb01",
    "version": "1.0",
    "name": "Exchange",
    "realm": "*",
    "serviceName": "00000002-0000-0ff1-ce00-000000000000",
    "issuer": "00000002-0000-0ff1-ce00-000000000000@*",
    "allowedAudiences": [
        "00000002-0000-0ff1-ce00-000000000000@*"
    ],
    "keys": [
        {
            "usage": "signing",
            "keyinfo": {
                "x5t": "enh9BJrVPU5ijV1qjZjV-fL2bco"
            },
            "keyvalue": {
                "type": "x509Certificate",
                "value": "MIIHNTCC..."
            }
        }
    ],
    "endpoints": [
        {
            "location": "https://by2pr06mb2229.namprd06.prod.outlook.com:444/autodiscover/metadata/json/1",
            "protocol": "OAuth2",
            "usage": "metadata"
        }
    ]
}
```

<span data-ttu-id="09151-146">使用可能な署名キーは `keys` 配列にあります。</span><span class="sxs-lookup"><span data-stu-id="09151-146">The available signing keys are in the `keys` array.</span></span> <span data-ttu-id="09151-147">`keyinfo` プロパティの `x5t` 値がトークンのヘッダーの `x5t` 値と一致することを確認することにより、正しいキーを選択します。</span><span class="sxs-lookup"><span data-stu-id="09151-147">Select the correct key by ensuring that the `x5t` value in the `keyinfo` property matches the `x5t` value in the header of the token.</span></span> <span data-ttu-id="09151-148">公開キーは `keyvalue` プロパティの `value` プロパティ内にあり、Base64 でエンコードされたバイト配列として格納されます。</span><span class="sxs-lookup"><span data-stu-id="09151-148">The public key is inside the `value` property in the `keyvalue` property, stored as a base64-encoded byte array.</span></span>

<span data-ttu-id="09151-149">正しい公開キーを取得したら、署名を検証します。</span><span class="sxs-lookup"><span data-stu-id="09151-149">After you have the correct public key, verify the signature.</span></span> <span data-ttu-id="09151-150">署名されたデータは、エンコードされたトークンの最初の 2 つの部分です (ピリオドで区切られる)。</span><span class="sxs-lookup"><span data-stu-id="09151-150">The signed data is the first two parts of the encoded token, separated by a period:</span></span>

```json
{header}.{payload}
```

## <a name="compute-the-unique-id-for-an-exchange-account"></a><span data-ttu-id="09151-151">Exchange アカウントの一意 ID を計算する</span><span class="sxs-lookup"><span data-stu-id="09151-151">Compute the unique ID for an Exchange account</span></span>

<span data-ttu-id="09151-152">Exchange アカウントの一意の識別子を作成するには、認証メタデータドキュメントの URL とアカウントの Exchange 識別子を連結します。</span><span class="sxs-lookup"><span data-stu-id="09151-152">You can create a unique identifier for an Exchange account by concatenating the authentication metadata document URL with the Exchange identifier for the account.</span></span> <span data-ttu-id="09151-153">この一意の ID を持っている場合は、その ID を使用して Outlook アドインの Web サービス用のシングル サインオン (SSO) システムを作成できます。</span><span class="sxs-lookup"><span data-stu-id="09151-153">When you have this unique identifier, you can use it to create a single sign-on (SSO) system for your Outlook add-in web service.</span></span> <span data-ttu-id="09151-154">SSO の一意の ID の使用の詳細については、「[Exchange の ID トークンを使用してユーザーを認証する](authenticate-a-user-with-an-identity-token.md)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="09151-154">For details about using the unique identifier for SSO, see [Authenticate a user with an identity token for Exchange](authenticate-a-user-with-an-identity-token.md).</span></span>

## <a name="use-a-library-to-validate-the-token"></a><span data-ttu-id="09151-155">ライブラリを使用してトークンを検証する</span><span class="sxs-lookup"><span data-stu-id="09151-155">Use a library to validate the token</span></span>

<span data-ttu-id="09151-156">一般的な JWT の解析と検証を行うことができるライブラリは数多くあります。</span><span class="sxs-lookup"><span data-stu-id="09151-156">There are a number of libraries that can do general JWT parsing and validation.</span></span> <span data-ttu-id="09151-157">Microsoft では、 `System.IdentityModel.Tokens.Jwt` Exchange のユーザー id トークンの検証に使用できるライブラリを提供しています。</span><span class="sxs-lookup"><span data-stu-id="09151-157">Microsoft provides the `System.IdentityModel.Tokens.Jwt` library that can be used to validate Exchange user identity tokens.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="09151-158">Exchange Web サービスマネージ API の使用は推奨されていません。ただし、現在は使用できません。このため、サポートされていないライブラリに依存しています。</span><span class="sxs-lookup"><span data-stu-id="09151-158">We no longer recommend the Exchange Web Services Managed API because the Microsoft.Exchange.WebServices.Auth.dll, though still available, is now obsolete and relies on unsupported libraries like Microsoft.IdentityModel.Extensions.dll.</span></span>

### <a name="systemidentitymodeltokensjwt"></a><span data-ttu-id="09151-159">System.IdentityModel.Tokens.Jwt</span><span class="sxs-lookup"><span data-stu-id="09151-159">System.IdentityModel.Tokens.Jwt</span></span>

<span data-ttu-id="09151-160">[System.IdentityModels.Tokens.Jwt](https://www.nuget.org/packages/System.IdentityModel.Tokens.Jwt) ライブラリはトークンを解析し、検証も実行できますが、ユーザー自身で `appctx` クレームを解析して公開署名キーを取得する必要があります。</span><span class="sxs-lookup"><span data-stu-id="09151-160">The [System.IdentityModels.Tokens.Jwt](https://www.nuget.org/packages/System.IdentityModel.Tokens.Jwt) library can parse the token and also perform the validation, though you will need to parse the `appctx` claim yourself and retrieve the public signing key.</span></span>

```cs
// Load the encoded token
string encodedToken = "...";
JwtSecurityToken jwt = new JwtSecurityToken(encodedToken);

// Parse the appctx claim to get the auth metadata url
string authMetadataUrl = string.Empty;
var appctx = jwt.Claims.FirstOrDefault(claim => claim.Type == "appctx");
if (appctx != null)
{
    var AppContext = JsonConvert.DeserializeObject<ExchangeAppContext>(appctx.Value);

    // Token version check
    if (string.Compare(AppContext.Version, "ExIdTok.V1", StringComparison.InvariantCulture) != 0) {
        // Fail validation
    }

    authMetadataUrl = AppContext.MetadataUrl;
}

// Use System.IdentityModel.Tokens.Jwt library to validate standard parts
JwtSecurityTokenHandler tokenHandler = new JwtSecurityTokenHandler();
TokenValidationParameters tvp = new TokenValidationParameters();

tvp.ValidateIssuer = false;
tvp.ValidateAudience = true;
tvp.ValidAudience = "{URL to add-in}";
tvp.ValidateIssuerSigningKey = true;
// GetSigningKeys downloads the auth metadata doc and
// returns a List<SecurityKey>
tvp.IssuerSigningKeys = GetSigningKeys(authMetadataUrl);
tvp.ValidateLifetime = true;

try
{
    var claimsPrincipal = tokenHandler.ValidateToken(encodedToken, tvp, out SecurityToken validatedToken);

    // If no exception, all standard checks passed
}
catch (SecurityTokenValidationException ex)
{
    // Validation failed
}
```

<br/>

<span data-ttu-id="09151-161">`ExchangeAppContext` クラスは次のように定義されます。</span><span class="sxs-lookup"><span data-stu-id="09151-161">The `ExchangeAppContext` class is defined as follows:</span></span>

```cs
using Newtonsoft.Json;

/// <summary>
/// Representation of the appctx claim in an Exchange user identity token.
/// </summary>
public class ExchangeAppContext
{
    /// <summary>
    /// The Exchange identifier for the user
    /// </summary>
    [JsonProperty("msexchuid")]
    public string ExchangeUid { get; set; }

    /// <summary>
    /// The token version
    /// </summary>
    public string Version { get; set; }

    /// <summary>
    /// The URL to download authentication metadata
    /// </summary>
    [JsonProperty("amurl")]
    public string MetadataUrl { get; set; }
}
```

<span data-ttu-id="09151-162">このライブラリを使用して Exchange トークンを検証し、`GetSigningKeys` の実装を持つ例については、「[Outlook-Add-In-Token-Viewer](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer)」を参照してください。</span><span class="sxs-lookup"><span data-stu-id="09151-162">For an example that uses this library to validate Exchange tokens and has an implementation of `GetSigningKeys`, see [Outlook-Add-In-Token-Viewer](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer).</span></span>

## <a name="see-also"></a><span data-ttu-id="09151-163">関連項目</span><span class="sxs-lookup"><span data-stu-id="09151-163">See also</span></span>

- [<span data-ttu-id="09151-164">Outlook-Add-In-Token-Viewer</span><span class="sxs-lookup"><span data-stu-id="09151-164">Outlook-Add-In-Token-Viewer</span></span>](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer)
- [<span data-ttu-id="09151-165">Outlook-Add-in-JavaScript-ValidateIdentityToken</span><span class="sxs-lookup"><span data-stu-id="09151-165">Outlook-Add-in-JavaScript-ValidateIdentityToken</span></span>](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ValidateIdentityToken)
