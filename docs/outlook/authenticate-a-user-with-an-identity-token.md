---
title: アドインで ID トークンを使用してユーザーを認証する
description: サービスで SSO を実装するために、Outlook アドインが提供する ID トークンの使用方法について説明します。
ms.date: 10/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: 4f2a24bb918cc55527adfb7b789004d3dd15d839
ms.sourcegitcommit: b66ba72aee8ccb2916cd6012e66316df2130f640
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484384"
---
# <a name="authenticate-a-user-with-an-identity-token-for-exchange"></a>Exchange の ID トークンを使用してユーザーを認証する

Exchange のユーザー ID トークンは、アドインがアドイン ユーザーを一意に識別する方法を提供します。 ユーザーの ID を確立することで、Outlook アドインを使用している顧客がサインインせずにサービスに接続できる、シングル サインオン (SSO) 認証スキームをバック エンド サービスに実装できます。 このトークンの種類を使用する場合の詳細については、「[Exchange のユーザー ID トークン](authentication.md#exchange-user-identity-token)」を参照してください。 この記事では、Exchange の ID トークンを使用してバックエンドにユーザーを認証する簡単な方法について説明します。

> [!IMPORTANT]
> これは、単なる SSO の簡単な実装例です。 従来どおり、ID と認証を処理するときは、コードが組織のセキュリティ要件を満たしていることを確認する必要があります。

## <a name="send-the-id-token-with-each-request"></a>要求ごとに ID トークンを送信する

最初の手順では、[getUserIdentityTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) を呼び出すことにより、アドインでサーバーから Exchange のユーザー ID トークンを取得します。 その次に、アドインはこのトークンを、バックエンドに対する各要求とともに送信します。 これはヘッダーか、要求の本文の一部として組み込まれます。

## <a name="validate-the-token"></a>トークンを検証する

バックエンドは、トークンを検証してから承諾する必要があります。 これは、トークンがユーザーの Exchange サーバーによって発行されたことを確認する重要な手順です。 Exchange のユーザー ID トークンの検証の詳細については、「[Exchange の ID トークンを検証する](validate-an-identity-token.md)」を参照してください。

検証およびデコードが完了すると、トークンのペイロードは次のようになります。

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

## <a name="map-the-token-to-a-user-in-your-backend"></a>トークンをバックエンドのユーザーにマップする

バックエンド サービスはトークンから一意のユーザー ID を計算し、内部ユーザー システムのユーザーにマップできます。 たとえば、ユーザーの格納にデータベースを使用する場合は、この一意の ID をデータベース内のユーザーのレコードに追加できます。

### <a name="generate-a-unique-id"></a>一意の ID を生成する

プロパティとプロパティの組み合わせを `msexchuid` 使用 `amurl` します。 たとえば、2 つの値を連結して、Base64 でエンコードされた文字列を生成します。 この値は毎回トークンから確実に生成できるので、Exchange のユーザー ID トークンをシステム内のユーザーにマップできます。

### <a name="check-the-user"></a>ユーザーを確認する

次の手順では、生成された一意の ID を使用して、関連付けられた ID でシステム内のユーザーを確認します。

- ユーザーが見つかった場合、バックエンドは要求を認証済みとして処理し、要求の続行を許可します。

- ユーザーが見つからない場合、バックエンドはユーザーがサインインする必要があることを示すエラーを返します。 その後アドインは、既存の認証方法を使用してバックエンドにサインインするように求めるダイアログを表示します。 ユーザーが認証されると、Exchange のユーザー ID トークンとユーザー認証の詳細が送信されます。 バックエンドはシステム内のユーザーのレコードを一意の ID で更新できます。
