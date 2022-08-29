---
title: イベント ベースのアクティブ化を使用する Outlook アドインでシングル サインオン (SSO) を有効にする
description: イベント ベースのアクティブ化アドインで作業するときに SSO を有効にする方法について説明します。
ms.date: 06/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: 10fd973c0476878443d7238e8805aa4db9f62953
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 08/24/2022
ms.locfileid: "67423119"
---
# <a name="enable-single-sign-on-sso-in-outlook-add-ins-that-use-event-based-activation"></a>イベント ベースのアクティブ化を使用する Outlook アドインでシングル サインオン (SSO) を有効にする

Outlook アドインがイベント ベースのアクティブ化を使用する場合、イベントは別のランタイムで実行 [されます](../testing/runtimes.md)。 [「Outlook アドインでシングル サインオン トークンを使用してユーザーを認証する」の手順を](authenticate-a-user-with-an-sso-token.md)完了したら、この記事で説明されている追加の手順に従って、イベント処理コードの SSO を有効にします。 SSO を有効にすると、 [getAccessToken() API](/javascript/api/office-runtime/officeruntime.auth) を呼び出して、ユーザーの ID を持つアクセス トークンを取得できます。

> [!IMPORTANT]
> アクセス トークンを取得するのと`Office.auth.getAccessToken`同じ機能を実行しながら`OfficeRuntime.auth.getAccessToken`、イベント ベースのアドインを呼び出`OfficeRuntime.auth.getAccessToken`することをお勧めします。 この API は、イベント ベースのアクティブ化と SSO をサポートするすべての Outlook クライアント バージョンでサポートされています。 一方、 `Office.auth.getAccessToken` バージョン 2111 (ビルド 14701.20000) 以降の Outlook on Windows でのみサポートされています。

Outlook on Windows の場合、Outlook アドインのマニフェストで、イベント ベースのアクティブ化のために読み込む 1 つの JavaScript ファイルを識別します。 また、このファイルが SSO をサポートできるように Office に指定する必要もあります。 これを行うには、既知の URI を使用して Office に提供するために、すべてのアドインとその JavaScript ファイルの一覧を作成します。

> [!NOTE]
> この記事の手順は、Windows で Outlook アドインを実行する場合にのみ適用されます。 これは、Outlook on Windows では JavaScript ファイルが使用され、Outlook on the webは同じ JavaScript ファイルを参照できる HTML ファイルを使用するためです。

## <a name="list-allowed-add-ins-with-a-well-known-uri"></a>既知の URI を使用して許可されているアドインを一覧表示する

SSO を使用できるアドインを一覧表示するには、アドインごとに各 JavaScript ファイルを識別する JSON ファイルを作成します。 次に、その JSON ファイルを既知の URI でホストします。 既知の URI を使用すると、現在の Web 配信元のトークンを取得する権限を持つ、ホストされているすべての JS ファイルを指定できます。 これにより、配信元の所有者が、どのホスト JS ファイルをアドインで使用するかを完全に制御できるため、偽装に関するセキュリティの脆弱性を防ぎます。

次の例は、2 つのアドイン (メイン バージョンとベータ バージョン) の SSO を有効にする方法を示しています。 Web サーバーから提供するアドインの数に応じて、必要な数のアドインを一覧表示できます。

```json
{
    "allowed":
    [
        "https://addin.contoso.com:8000/main/js/autorun.js",
        "https://addin.contoso.com:8000/beta/js/autorun.js"
    ]
}
```

配信元のルートにある URI で名前が付けられた `.well-known` 場所で JSON ファイルをホストします。 たとえば、配信元が次の場合、 `https://addin.contoso.com:8000/`よく知られた URI は `https://addin.contoso.com:8000/.well-known/microsoft-officeaddins-allowed.json`.

配信元は、スキーム + サブドメイン + ドメイン + ポートのパターンを参照します。 **場所**`.well-known`の名前は 、リソース ファイルの名前にする **必要があります**`microsoft-officeaddins-allowed.json`。 このファイルには、それぞれのアドインの SSO が許可されているすべての JavaScript ファイルの配列である属性を持 `allowed` つ JSON オブジェクトが含まれている必要があります。

## <a name="see-also"></a>関連項目

- [Outlook アドインでシングル サインオン トークンを使用してユーザーを認証する](authenticate-a-user-with-an-sso-token.md)
- [イベント ベースのアクティブ化のために Outlook アドインを構成する](autolaunch.md)
