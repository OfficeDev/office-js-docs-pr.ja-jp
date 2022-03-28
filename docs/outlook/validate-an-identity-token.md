---
title: Outlook アドイン ID トークンを検証する
description: 使用している Outlook アドインから Exchange のユーザー ID トークンを送信できますが、要求を信頼する前に、トークンを検証して適切な Exchange サーバーからのものであることを確認する必要があります。
ms.date: 10/11/2021
ms.localizationpriority: medium
ms.openlocfilehash: 6b903b582fee59fd1c5ff0aa949d614c4ee1dff7
ms.sourcegitcommit: b66ba72aee8ccb2916cd6012e66316df2130f640
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484414"
---
# <a name="validate-an-exchange-identity-token"></a>Exchange の ID トークンを検証する

使用している Outlook アドインから Exchange のユーザー ID トークンを送信できますが、要求を信頼する前に、トークンを検証して適切な Exchange サーバーからのものであることを確認する必要があります。 Exchange のユーザー ID トークンは、JSON Web トークン (JWT) です。 JWT の検証に必要な手順は、「[RFC 7519 JSON Web Token (JWT)](https://www.rfc-editor.org/rfc/rfc7519.txt)」に記載されています。

ID トークンの検証およびユーザーの一意識別子の取得は 4 つのステップで進めることをお勧めします。 まず、Base 64 URL 形式でエンコードされた文字列から、JSON Web トークン (JWT) を抽出します。 次に、トークンが整形式であること、使用する Outlook アドイン向けのトークンであること、有効期限が切れていないこと、および認証メタデータ ドキュメントの有効な URL を抽出できることを確認します。 その後、Exchange サーバーから認証メタデータ ドキュメントを取得し、ID トークンに添付されている署名を検証します。 最後に、ユーザーの ID と認証メタデータ ドキュメントの URL をExchangeして、ユーザーの一意の識別子を計算します。

## <a name="extract-the-json-web-token"></a>JSON Web トークンを抽出する

[getUserIdentityTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) から返されたトークンは、トークンのエンコードされた文字列表現です。 この形式では、すべての JWT にピリオドで区切られた 3 つの部分があります (RFC 7519 ごと)。 形式は次のようになります。

```json
{header}.{payload}.{signature}
```

ヘッダーとペイロードは、各部分の JSON 表現を取得するために、Base64 でデコードされる必要があります。 署名は、バイナリ シグネチャを含むバイト配列を取得するために、Base64 でデコードされる必要があります。

トークンの内容の詳細については、「[Exchange の ID トークンの内部](inside-the-identity-token.md)」を参照してください。

デコードされた 3 つのコンポーネントがあれば、トークンの内容の検証を進めることができます。

## <a name="validate-token-contents"></a>トークンの内容を検証する

トークンの内容を検証するには、次の内容を確認する必要があります。

- ヘッダーを確認し、次の項目を確認します。
  - `typ` claim は に設定されています `JWT`。
  - `alg` claim は に設定されています `RS256`。
  - `x5t` クレームが存在します。

- ペイロードを確認し、次の情報を確認します。
  - `amurl` 内のクレーム `appctx` は、承認されたトークン署名キー マニフェスト ファイルの場所に設定されます。 たとえば、この値の期待`amurl`値Microsoft 365ですhttps://outlook.office365.com:443/autodiscover/metadata/json/1。 詳細については、次のセクション [「ドメインを確認](#verify-the-domain) する」を参照してください。
  - 現在の時刻は、クレームで指定された時刻 `nbf` の `exp` 間です。 `nbf` クレームは、トークンが有効と考えられる最も早い時刻を指定し、`exp` クレームはトークンの有効期限を指定します。 サーバー間のクロック設定には、ある程度の変動を許可することをお勧めします。
  - `aud` claim はアドインの予想される URL です。
  - `version` クレーム内のクレームは `appctx` に設定されます `ExIdTok.V1`。

### <a name="verify-the-domain"></a>ドメインを確認する

前のセクションで説明した `amurl` 検証ロジックを実装する場合は、クレームのドメインがユーザーの自動検出ドメインと一致するように要求する必要があります。 これを行うには、自動検出を使用または実装する必要[Exchange](/exchange/client-developer/exchange-web-services/autodiscover-for-exchange)。

- たとえばExchange Online`amurl`、既知のドメイン (https://outlook.office365.com:443/autodiscover/metadata/json/1)または地域固有または特殊なクラウド (Office 365 URL と IP アドレス範囲) に[属しているか確認します](/microsoft-365/enterprise/urls-and-ip-address-ranges?view=o365-worldwide&preserve-view=true)。

- アドイン サービスにユーザーのテナントを持つ既存の構成がある場合は、これが信頼される場合に確立 `amurl` できます。

- ハイブリッド展開[Exchange、](/microsoft-365/enterprise/configure-exchange-server-for-hybrid-modern-authentication?view=o365-worldwide&preserve-view=true)OAuth ベースの自動検出を使用して、ユーザーに必要なドメインを確認します。 ただし、ユーザーは自動検出フローの一部として認証する必要がある一方で、アドインはユーザーの資格情報を収集して基本認証を行う必要があります。

`amurl`アドインでこれらのオプションの使用を確認できない場合は、アドインのワークフローに認証が必要な場合は、ユーザーに適切な通知を受け取ってアドインを正常にシャットダウンできます。

## <a name="validate-the-identity-token-signature"></a>ID トークンの署名を検証する

JWT に必要なクレームが含まれていることを確認したら、トークンの署名の検証を進めることができます。

### <a name="retrieve-the-public-signing-key"></a>公開署名キーを取得する

最初のステップでは、Exchange サーバーがトークンの署名に使用した証明書に対応する公開キーを取得します。 鍵は認証メタデータ ドキュメントに記載されています。 このドキュメントは、`amurl` クレームで指定された URL でホストされている JSON ファイルです。

認証メタデータ ドキュメントには、次の形式を使用します。

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

使用可能な署名キーは `keys` 配列にあります。 `keyinfo` プロパティの `x5t` 値がトークンのヘッダーの `x5t` 値と一致することを確認することにより、正しいキーを選択します。 公開キーは `keyvalue` プロパティの `value` プロパティ内にあり、Base64 でエンコードされたバイト配列として格納されます。

正しい公開キーを取得したら、署名を検証します。 署名されたデータは、エンコードされたトークンの最初の 2 つの部分です (ピリオドで区切られる)。

```json
{header}.{payload}
```

## <a name="compute-the-unique-id-for-an-exchange-account"></a>Exchange アカウントの一意 ID を計算する

認証メタデータ ドキュメント URL とアカウントExchange識別子を連結して、Exchangeアカウントの一意の識別子を作成します。 この一意の識別子を使用して、アドイン Web サービス用のシングル サインオン (SSO) システムOutlook作成します。 SSO の一意の ID の使用の詳細については、「[Exchange の ID トークンを使用してユーザーを認証する](authenticate-a-user-with-an-identity-token.md)」を参照してください。

## <a name="use-a-library-to-validate-the-token"></a>ライブラリを使用してトークンを検証する

一般的な JWT の解析と検証を行うことができるライブラリは数多くあります。 Microsoft は、ユーザー `System.IdentityModel.Tokens.Jwt` ID トークンの検証に使用できるExchange提供しています。

> [!IMPORTANT]
> Microsoft.Exchange.WebServices.Auth.dll は使用できなくなったため、Exchange Web Services マネージ API は廃止され、Microsoft.IdentityModel.Extensions.dll のようなサポートされていないライブラリに依存しています。

### <a name="systemidentitymodeltokensjwt"></a>System.IdentityModel.Tokens.Jwt

[System.IdentityModels.Tokens.Jwt](https://www.nuget.org/packages/System.IdentityModel.Tokens.Jwt) ライブラリはトークンを解析し、検証も実行できますが、ユーザー自身で `appctx` クレームを解析して公開署名キーを取得する必要があります。

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

`ExchangeAppContext` クラスは次のように定義されます。

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

このライブラリを使用して Exchange トークンを検証し、`GetSigningKeys` の実装を持つ例については、「[Outlook-Add-In-Token-Viewer](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer)」を参照してください。

## <a name="see-also"></a>関連項目

- [Outlook-Add-In-Token-Viewer](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer)
- [Outlook-Add-in-JavaScript-ValidateIdentityToken](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ValidateIdentityToken)
