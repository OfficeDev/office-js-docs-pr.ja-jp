---
title: イベント ベースのライセンス認証を使用するOutlookでシングル サインオン (SSO) を有効にする
description: イベント ベースのアクティブ化アドインで作業するときに SSO を有効にする方法について学習します。
ms.date: 03/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: 38c717e0d626f4c135f76350e30398db26cac24f
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746536"
---
# <a name="enable-single-sign-on-sso-in-outlook-add-ins-that-use-event-based-activation"></a>イベント ベースのライセンス認証を使用するOutlookでシングル サインオン (SSO) を有効にする

イベント Outlookがイベント ベースのアクティブ化を使用する場合、イベントは別の JavaScript ランタイムで実行されます。 [「Outlook](authenticate-a-user-with-an-sso-token.md) アドインでシングル サインオン トークンを使用してユーザーを認証する」の手順を完了した後、この記事で説明する追加の手順に従って、イベント処理コードの SSO を有効にします。 SSO を有効にしたら、API を呼び出 `getAccessToken()` して、ユーザーの ID を持つアクセス トークンを取得できます。

> [!NOTE]
> この記事の手順は、アドインを Outlookで実行する場合にのみWindows。 これは、OutlookのWindows JavaScript ファイルを使用し、Outlook on the webは同じ JavaScript ファイルを参照できる HTML ファイルを使用する場合です。

Outlook Windows Outlook アドインのマニフェストで、イベント ベースのアクティブ化のために読み込む単一の JavaScript ファイルを識別します。 また、このファイルが SSO をサポートOffice許可されている場合は、そのファイルを指定する必要があります。 これを行うには、すべてのアドインとその JavaScript ファイルの一覧を作成して、既知の URI Officeを提供します。

## <a name="list-allowed-add-ins-with-a-well-known-uri"></a>よく知られている URI を使用して許可されているアドインを一覧表示する

SSO を使用できるアドインを一覧表示するには、各アドインの各 JavaScript ファイルを識別する JSON ファイルを作成します。 次に、その JSON ファイルを既知の URI でホストします。 既知の URI を使用すると、現在の Web オリジンのトークンを取得する権限を持つすべてのホストされた JS ファイルを指定できます。 これにより、オリジンの所有者が、アドインで使用するホストされた JS ファイルと使用しないファイルを完全に制御し、偽装に関するセキュリティ上の脆弱性を防止します。

次の例は、2 つのアドイン (メイン バージョンとベータ版) で SSO を有効にする方法を示しています。 Web サーバーから提供する数に応じて、必要な数のアドインを一覧表示できます。

```json
{
    "allowed":
    [
        "https://addin.contoso.com:8000/main/js/autorun.js",
        "https://addin.contoso.com:8000/beta/js/autorun.js"
    ]
}
```

元のルートにある URI で `.well-known` 指定された場所の下に JSON ファイルをホストします。 たとえば、原点がである場合 `https://addin.contoso.com:8000/`、既知の URI はです `https://addin.contoso.com:8000/.well-known/microsoft-officeaddins-allowed.json`。

原点は、スキーム + サブドメイン + ドメイン + ポートのパターンを参照します。 場所の名前は 、 **リソース** `.well-known`ファイルの名前を指定する **必要** があります `microsoft-officeaddins-allowed.json`。 このファイルには、それぞれのアドインの SSO `allowed` に対して承認されたすべての JavaScript ファイルの配列である値という名前の属性を持つ JSON オブジェクトが含まれている必要があります。

## <a name="see-also"></a>関連項目

- [アドイン内のシングル サインオン トークンを使用してユーザー Outlook認証する](authenticate-a-user-with-an-sso-token.md)
- [イベント ベースのOutlookアドインを構成する](autolaunch.md)
