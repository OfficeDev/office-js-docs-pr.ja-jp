---
title: Outlook アドインでの Exchange の ID トークンの内部
description: Outlook アドインから生成される Exchange のユーザー ID トークンの内容について説明します。
ms.date: 10/31/2019
ms.localizationpriority: medium
ms.openlocfilehash: 843bd76b66f784b1e380bdde5e33adf05755e268
ms.sourcegitcommit: b66ba72aee8ccb2916cd6012e66316df2130f640
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484052"
---
# <a name="inside-the-exchange-identity-token"></a>Exchange の ID トークンの内部

[getUserIdentityTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) メソッドによって返された Exchange のユーザー ID トークンは、アドイン コードがバックエンド サービスへの呼び出しでユーザー ID を含めるための方法を提供します。 この記事では、トークンの形式と内容について説明します。

Exchange ユーザー ID トークンとは、そのトークンを送信する Exchange サーバーによって自己署名された、Base 64 URL 形式でエンコードされた文字列です。 トークンは暗号化されていません。署名の検証に使用する公開キーは、トークンを発行した Exchange サーバーに保存されています。 トークンには 3 つのパーツ (ヘッダー、ペイロード、署名) があります。 トークン文字列では、トークンを容易に分割できるように、パーツがピリオド文字 (`.`) で区切られています。

Exchange では ID トークンに、JSON Web トークン (JWT) 形式を使用します。 JWT トークンの詳細については、「[RFC 7519 JSON Web Token (JWT)](https://www.rfc-editor.org/rfc/rfc7519.txt)」を参照してください。

## <a name="identity-token-header"></a>ID トークンのヘッダー

ヘッダーは、トークンの形式に関する情報と、署名情報に関する情報を提供します。 次の例は、トークンのヘッダーの外観を示しています。

```JSON
{
  "typ": "JWT",
  "alg": "RS256",
  "x5t": "Un6V7lYN-rMgaCoFSTO5z707X-4"
}
```

<br/>
 
トークンのヘッダーのパーツについての説明を、次の表に示します。

| クレーム | 値 | 説明 |
|:-----|:-----|:-----|
| `typ` | `JWT` | トークンを JSON Web トークンとして識別します。 Exchange サーバーから提供される ID トークンは、すべて JWT トークンです。 |
| `alg` | `RS256` | 署名の作成に使用されるハッシュ アルゴリズム。 Exchange サーバーから提供されるトークンは、すべて SHA-256 ハッシュ アルゴリズムの RSASSA-PKCS1-v1_5 を使用します。 |
| `x5t` | 証明書の拇印 | トークンの X.509 拇印です。 |

## <a name="identity-token-payload"></a>ID トークンのペイロード

ペイロードには、電子メール アカウントの識別と、トークンを送信した Exchange サーバーの識別を行う認証クレームが含まれます。以下に、ペイロード セクションの例を示します。

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
 
ID トークンのペイロードのパーツを、次の表に示します。

| クレーム | 説明 |
|:-----|:-----|
| `aud` | トークンを要求したアドインの URL。 トークンは、クライアントのブラウザー内で実行されているアドインから送信された場合にのみ有効です。 アドインが Office アドイン マニフェスト スキーマ v1.1 を使用する場合、この URL はフォーム タイプ `ItemRead` または `ItemEdit` (アドイン マニフェスト内で [FormSettings](/javascript/api/manifest/formsettings) 要素の一部として最初に出現する方) の下にある、最初の `SourceLocation` 要素に指定された URL になります。 |
| `iss` | トークンを発行した Exchange サーバーの一意の識別子です。この Exchange サーバーから発行されるトークンはすべて同じ識別子になります。 |
| `nbf` | トークンの有効期間の開始日時です。この値は 1970 年 1 月 1 日を起点とする秒数です。 |
| `exp` | トークンの有効期間の終了日時です。この値も 1970 年 1 月 1 日を起点とする秒数です。 |
| `appctxsender` | アプリケーション コンテキストを送信した Exchange サーバーの一意の識別子。 |
| `isbrowserhostedapp` | アドインがブラウザーでホストされるかどうかを指定します。 |
| `appctx` | トークンのアプリケーション コンテキスト。 |

appctx クレーム内の情報は、アカウントの一意の識別子と、トークンの署名に使用された公開キーの場所を提供します。 `appctx` クレームのパーツを、次の表に示します。

| アプリケーション コンテキスト プロパティ | 説明 |
|:-----|:-----|
| `msexchuid` | 電子メール アカウントと Exchange サーバーに割り当てられた一意の識別子。 |
| `version` | トークンのバージョン番号。 Exchange によって提供されるトークンの値は、すべて `ExIdTok.V1` になります。 |
| `amurl` | トークンに署名するために使用された X.509 証明書の公開キーが含まれる認証メタデータ ドキュメントの URL。<br/><br/>認証メタデータ ドキュメントの使用方法については、「[Exchange の ID トークンを検証する](validate-an-identity-token.md)」を参照してください。 |

## <a name="identity-token-signature"></a>ID トークンの署名

この署名は、ヘッダーおよびペイロード セクションに対して、ヘッダーで指定されたアルゴリズムを使用したハッシュ処理を行うと共に、ペイロード内で指定された場所にあるサーバー上の自己署名された X509 証明書を使用することで作成されます。Web サービスは、この署名を検証して、ID トークンがその送信元として想定されるサーバーから発行されたものであることを確認できます。

## <a name="see-also"></a>関連項目

Exchange のユーザー ID トークンの解析例については、「[Outlook アドイン トークン ビューアー](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer)」を参照してください。
