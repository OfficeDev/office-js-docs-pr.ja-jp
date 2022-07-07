---
title: Outlook アドインの API
description: Outlook アドインの API を参照して、Outlook アドインにアクセス許可を宣言する方法について説明します。
ms.date: 06/30/2022
ms.localizationpriority: medium
ms.openlocfilehash: 583d2b07a0590e7a04b052d5675320b8ea73a61f
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 07/06/2022
ms.locfileid: "66660257"
---
# <a name="outlook-add-in-apis"></a>Outlook アドインの API

Outlook アドインで API を使用するには、Office.js ライブラリの場所、要件セット、スキーマ、アクセス許可を指定する必要があります。 主に [、Mailbox](#mailbox-object) オブジェクトを通じて公開される Office JavaScript API を使用します。

## <a name="officejs-library"></a>Office.js ライブラリ

[Outlook アドイン API](/javascript/api/outlook) を操作するには、Office.jsで JavaScript API を使用する必要があります。 ライブラリのコンテンツ配信ネットワーク (CDN) は `https://appsforoffice.microsoft.com/lib/1/hosted/Office.js`. AppSource に送られるアドインは、この CDN で Office.js を参照しなければなりません。ローカル参照は使用できません。

アドインの UI を実行する Web ページ (.html、.aspx、.php のファイル) の `<head>` タグの中の `<script>` タグの中で CDN を参照します。

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

新しい API が追加されても、Office.js への URL は同じままになります。URL 内のバージョンは、既存の API の動作を分割する場合にのみ変更されます。

> [!IMPORTANT]
> 任意の Office クライアント アプリケーション用のアドインを開発する場合は、ページのセクション内 `<head>` から Office JavaScript API を参照します。 これにより、あらゆる body 要素の前に API が完全に初期化されます。

## <a name="requirement-sets"></a>要件セット

すべての Outlook API は [、メールボックス要件セット](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets)に属しています。 `Mailbox` の要件セットにはバージョンがあり、リリースされる新しい API の各セットは、新しいバージョンのセットに属しています。 最新の API セットがリリースされても、すべての Outlook クライアントがそれをサポートするわけではありませんが、ある Outlook クライアントが 1 つの要件セットのサポートを宣言した場合、その要件セットの中のすべての API がサポートされます。

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

それなしではアドインの機能が機能しないような、シナリオに絶対必要な API のセットをサポートする最低限要件セットを指定します。 要素のマニフェストで要件セットを **\<Requirements\>** 指定します。 詳細は、[Outlook のアドイン マニフェスト](manifests.md)と「[Outlook API 要件セットについて](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets)」を参照してください。

この要素は **\<Methods\>** Outlook アドインには適用されないため、特定のメソッドのサポートを宣言することはできません。

## <a name="permissions"></a>アクセス許可

アドインには、そのアドインが必要とする API を使用するための適切なアクセス許可が必要になります。アクセス許可には、4 つのレベルがあります。詳細については、「[Outlook アドインのアクセス許可モデルを理解する](understanding-outlook-add-in-permissions.md)」を参照してください。

<br/>

|権限レベル|説明|
|:-----|:-----|
| **制限付き** | エンティティは使用できますが、正規表現は使用できません。 |
| **アイテムの読み取り** | **制限付き** で許可されているものに加えて、以下のものが許可されます。<ul><li>正規表現</li><li>Outlook アドイン API の読み取りアクセス</li><li>アイテムのプロパティとコールバック トークンの取得</li></ul> |
| **読み取り/書き込み** | **アイテムの読み取り** で許可される内容に加えて、次に示す内容が許可されます。<ul><li>`makeEwsRequestAsync` を除いた、完全な Outlook アドイン API のアクセス</li><li>アイテムのプロパティの設定</li></ul> |
| **メールボックスの読み取り/書き込み** | **読み取り/書き込み** で許可されているものに加えて、以下のものが許可されます。<ul><li>アイテムやフォルダーの作成、読み取り、書き込み</li><li>アイテムの送信</li><li>[makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) の呼び出し</li></ul> |

一般的には、アドインに必要な最低限のアクセス許可を指定する必要があります。 アクセス許可は、マニフェストの **\<Permissions\>** 要素で宣言されます。 詳細については、「[Outlook アドインのマニフェスト](manifests.md)」を参照してください。 セキュリティの問題については、「 [Office アドインのプライバシーとセキュリティ](../concepts/privacy-and-security.md)」を参照してください。

## <a name="mailbox-object"></a>Mailbox オブジェクト

[!include[information about Mailbox object](../includes/mailbox-object-desc.md)]

## <a name="see-also"></a>関連項目

- [Outlook アドインのマニフェスト](manifests.md)
- [Outlook API 要件セットについて](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets)
- [Office アドインのプライバシーとセキュリティ](../concepts/privacy-and-security.md)
