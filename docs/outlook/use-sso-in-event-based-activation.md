---
title: イベント ベースのライセンス認証を使用するOutlookでシングル サインオン (SSO) を有効にする
description: イベント ベースのアクティブ化アドインで作業するときに SSO を有効にする方法について学習します。
ms.date: 11/16/2021
ms.localizationpriority: medium
ms.openlocfilehash: 66d1edb8b7b0092ee107b73af24d5420caee8677
ms.sourcegitcommit: 6e6c4803fdc0a3cc2c1bcd275288485a987551ff
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 11/18/2021
ms.locfileid: "61066654"
---
# <a name="enable-single-sign-on-sso-in-outlook-add-ins-that-use-event-based-activation"></a>イベント ベースのライセンス認証を使用するOutlookでシングル サインオン (SSO) を有効にする

イベント Outlookがイベント ベースのアクティブ化を使用する場合、イベントは別の JavaScript ランタイムで実行されます。 [「Outlook](authenticate-a-user-with-an-sso-token.md)アドインでシングル サインオン トークンを使用してユーザーを認証する」の手順を完了した後、この記事で説明する追加の手順に従って、イベント処理コードの SSO を有効にします。 SSO を有効にしたら、API を呼び出して、ユーザーの ID を持つアクセス `getAccessToken()` トークンを取得できます。

> [!NOTE]
> この記事の手順は、アプリでアドインOutlook実行する場合にのみWindows。 これは、OutlookのWindows JavaScript ファイルを使用し、Outlook on the web同じ JavaScript ファイルを参照できる HTML ファイルを使用する場合です。

Outlook Windows Outlook アドインのマニフェストで、イベント ベースのアクティブ化のために読み込む 1 つの JavaScript ファイルを識別します。 また、このファイルが SSO をサポートOffice許可されている場合は、そのファイルを指定する必要があります。 これを行うには、2 つの方法があります。 すべてのアドインとその JavaScript ファイルの一覧を作成して、既知の URI を使用Officeを提供できます。 または、SSO を有効にするカスタム応答ヘッダーを追加できます。

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

元のルートにある URI で指定された場所の `.well-known` 下に JSON ファイルをホストします。 たとえば、原点がである場合 `https://addin.contoso.com:8000/` 、既知の URI はです `https://addin.contoso.com:8000/.well-known/microsoft-officeaddins-allowed.json` 。

原点は、スキーム + サブドメイン + ドメイン + ポートのパターンを参照します。 場所の名前は 、 **リソース** ファイルの名前を指定する `.well-known` **必要** があります `microsoft-officeaddins-allowed.json` 。 このファイルには、それぞれのアドインの SSO に対して承認されたすべての JavaScript ファイルの配列である値という名前の属性を持つ JSON オブジェクトが `allowed` 含まれている必要があります。

## <a name="add-a-custom-response-header"></a>カスタム応答ヘッダーの追加

2 つ目の方法は、という名前のカスタム応答ヘッダーを追加します `MS-OfficeAddins-Allowed-Origin` 。 ヘッダーの値は、JavaScript ファイルの原点である必要があります。

たとえば、JavaScript ファイルが場所にある場合は `https://addin.contoso.com:8000/main/js/autorun.js` 、次の応答ヘッダーを追加します。

`MS-OfficeAddins-Allowed-Origin : https://addin.contoso.com:8000`

カスタム応答ヘッダーを追加する方法については、特定の Web サーバーのドキュメントを参照する必要があります。

## <a name="see-also"></a>関連項目

- [アドイン内のシングル サインオン トークンを使用してユーザー Outlook認証する](authenticate-a-user-with-an-sso-token.md)
- [イベント ベースのOutlook用にアドインを構成する](autolaunch.md)
