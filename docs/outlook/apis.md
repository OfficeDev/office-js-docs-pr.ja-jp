---
title: Outlook アドインの API
description: Outlook アドインの API を参照して、Outlook アドインにアクセス許可を宣言する方法について説明します。
ms.date: 01/14/2022
ms.localizationpriority: medium
ms.openlocfilehash: 44b5b770d36177307989500db89f1f4f8ca859ec
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745690"
---
# <a name="outlook-add-in-apis"></a>Outlook アドインの API

Outlook アドインで API を使用するには、Office.js ライブラリの場所、要件セット、スキーマ、アクセス許可を指定する必要があります。 メールボックス オブジェクトを通じて公開Office JavaScript API を主に[使用](#mailbox-object)します。

## <a name="officejs-library"></a>Office.js ライブラリ

Outlook アドイン API と対話操作するには、Office.js の JavaScript API を使用する必要があります。 ライブラリのコンテンツ配信ネットワーク (CDN) はです`https://appsforoffice.microsoft.com/lib/1/hosted/Office.js`。 AppSource に送られるアドインは、この CDN で Office.js を参照しなければなりません。ローカル参照は使用できません。

アドインの UI を実行する Web ページ (.html、.aspx、.php のファイル) の `<head>` タグの中の `<script>` タグの中で CDN を参照します。

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

新しい API が追加されても、Office.js への URL は同じままになります。URL 内のバージョンは、既存の API の動作を分割する場合にのみ変更されます。

> [!IMPORTANT]
> 任意のクライアント アプリケーション用のアドインを開発Office、ページのセクションOffice JavaScript API `<head>` を参照します。 これにより、あらゆる body 要素の前に API が完全に初期化されます。

## <a name="requirement-sets"></a>要件セット

すべての Outlook API は `Mailbox` 要件セットに属しています。 `Mailbox` の要件セットにはバージョンがあり、リリースされる新しい API の各セットは、新しいバージョンのセットに属しています。 最新の API セットがリリースされても、すべての Outlook クライアントがそれをサポートするわけではありませんが、ある Outlook クライアントが 1 つの要件セットのサポートを宣言した場合、その要件セットの中のすべての API がサポートされます。

どの Outlook クライアントにアドインを表示するかを制御するには、最小の要件セットのバージョンをマニフェストで指定します。たとえば、要件セットのバージョン 1.3 を指定すると、最小バージョンの 1.3 をサポートしていない Outlook クライアントにはアドインが表示されなくなります。

要件セットを指定しても、そのバージョンの API にアドインを限定することにはなりません。要件セット v1.1 を指定しているアドインが、v1.3 をサポートする Outlook クライアントで実行されると、そのアドインは v1.3 の API を使用できます。要件セットでは、どの Outlook クライアントにアドインを表示するかのみを制御します。

マニフェストで指定した要件セットよりも上位の要件セットの API が使用できるかどうかを確認する場合は、標準の JavaScript を使用できます。

```js
if (item.somePropertyOrFunction) {
   item.somePropertyOrFunction...  
}
```

> [!NOTE]
> このような確認は、マニフェストで指定された要件セットのバージョンに存在する API には必要ありません。

それなしではアドインの機能が機能しないような、シナリオに絶対必要な API のセットをサポートする最低限要件セットを指定します。 要件セットは、`<Requirements>` 要素内のマニフェストで指定します。 詳細は、[Outlook のアドイン マニフェスト](manifests.md)と「[Outlook API 要件セットについて](../reference/requirement-sets/outlook-api-requirement-sets.md)」を参照してください。

`<Methods>` 要素は Outlook アドインには適用されないので、特定のメソッドのサポートは宣言できません。

## <a name="permissions"></a>アクセス許可

アドインには、そのアドインが必要とする API を使用するための適切なアクセス許可が必要になります。アクセス許可には、4 つのレベルがあります。詳細については、「[Outlook アドインのアクセス許可モデルを理解する](understanding-outlook-add-in-permissions.md)」を参照してください。

<br/>

|権限レベル|説明|
|:-----|:-----|
| **制限付き** | エンティティは使用できますが、正規表現は使用できません。 |
| **アイテムの読み取り** | **制限付き** で許可されているものに加えて、以下のものが許可されます。<ul><li>正規表現</li><li>Outlook アドイン API の読み取りアクセス</li><li>アイテムのプロパティとコールバック トークンの取得</li></ul> |
| **読み取り/書き込み** | **アイテムの読み取り** で許可される内容に加えて、次に示す内容が許可されます。<ul><li>`makeEwsRequestAsync` を除いた、完全な Outlook アドイン API のアクセス</li><li>アイテムのプロパティの設定</li></ul> |
| **メールボックスの読み取り/書き込み** | **読み取り/書き込み** で許可されているものに加えて、以下のものが許可されます。<ul><li>アイテムやフォルダーの作成、読み取り、書き込み</li><li>アイテムの送信</li><li>[makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) の呼び出し</li></ul> |

一般的には、アドインに必要な最低限のアクセス許可を指定する必要があります。 アクセス許可は、マニフェスト内の `<Permissions>` 要素で宣言されます。 詳細については、「[Outlook アドインのマニフェスト](manifests.md)」を参照してください。 セキュリティの問題の詳細については、「プライバシーとセキュリティ[」を参照Officeアドインを参照してください](../concepts/privacy-and-security.md)。

## <a name="mailbox-object"></a>Mailbox オブジェクト

[!include[information about Mailbox object](../includes/mailbox-object-desc.md)]

## <a name="see-also"></a>関連項目

- [Outlook アドインのマニフェスト](manifests.md)
- [Outlook API 要件セットについて](../reference/requirement-sets/outlook-api-requirement-sets.md)
- [Office アドインのプライバシーとセキュリティ](../concepts/privacy-and-security.md)
