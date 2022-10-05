---
title: Outlook アドインの API
description: Outlook アドインの API を参照して、Outlook アドインにアクセス許可を宣言する方法について説明します。
ms.date: 10/03/2022
ms.localizationpriority: medium
ms.openlocfilehash: 69043646add5e32502efb0d2a5b1259667e564d9
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467077"
---
# <a name="outlook-add-in-apis"></a>Outlook アドインの API

Outlook アドインで API を使用するには、Office.js ライブラリの場所、要件セット、スキーマ、アクセス許可を指定する必要があります。 主に [、Mailbox](#mailbox-object) オブジェクトを通じて公開される Office JavaScript API を使用します。

## <a name="officejs-library"></a>Office.js ライブラリ

[Outlook アドイン API](/javascript/api/outlook) を操作するには、Office.jsで JavaScript API を使用する必要があります。 ライブラリのコンテンツ配信ネットワーク (CDN) は `https://appsforoffice.microsoft.com/lib/1/hosted/Office.js`. AppSource に送られるアドインは、この CDN で Office.js を参照しなければなりません。ローカル参照は使用できません。

アドインの UI を実行する Web ページ (.html、.aspx、.php のファイル) の `<head>` タグの中の `<script>` タグの中で CDN を参照します。

```HTML
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
```

As we add new APIs, the URL to Office.js will stay the same. We will change the version in the URL only if we break an existing API behavior.

> [!IMPORTANT]
> 任意の Office クライアント アプリケーション用のアドインを開発する場合は、ページのセクション内 `<head>` から Office JavaScript API を参照します。 これにより、あらゆる body 要素の前に API が完全に初期化されます。

## <a name="requirement-sets"></a>要件セット

すべての Outlook API は [、メールボックス要件セット](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets)に属しています。 `Mailbox` の要件セットにはバージョンがあり、リリースされる新しい API の各セットは、新しいバージョンのセットに属しています。 最新の API セットがリリースされても、すべての Outlook クライアントがそれをサポートするわけではありませんが、ある Outlook クライアントが 1 つの要件セットのサポートを宣言した場合、その要件セットの中のすべての API がサポートされます。

To control which Outlook clients the add-in appears in, specify a minimum requirement set version in the manifest. For example, if you specify requirement set version 1.3, the add-in will not show up in any Outlook client that doesn't support a minimum version of 1.3.

Specifying a requirement set doesn't limit your add-in to the APIs in that version. If the add-in specifies requirement set v1.1 but is running in an Outlook client that supports v1.3, the add-in can still use v1.3 APIs. The requirement set only controls which Outlook clients the add-in appears in.

マニフェストで指定した要件セットよりも上位の要件セットの API が使用できるかどうかを確認する場合は、標準の JavaScript を使用できます。

```js
if (item.somePropertyOrFunction) {
   item.somePropertyOrFunction...  
}
```

> [!NOTE]
> このような確認は、マニフェストで指定された要件セットのバージョンに存在する API には必要ありません。

それなしではアドインの機能が機能しないような、シナリオに絶対必要な API のセットをサポートする最低限要件セットを指定します。 マニフェストで要件セットを指定します。 マークアップは、使用しているマニフェストによって異なります。 

- **XML マニフェスト**: 要素を使用します **\<Requirements\>** 。 **\<Methods\>** Outlook アドインでは子要素 **\<Requirements\>** がサポートされていないため、特定のメソッドのサポートを宣言することはできません。
- **Teams マニフェスト (プレビュー)**: "extensions.capabilities" プロパティを使用します。 

詳細については、「 [Outlook アドイン マニフェスト](manifests.md)」と「 [Outlook API 要件セットについて」を](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets)参照してください。

## <a name="permissions"></a>アクセス許可

アドインには、そのアドインが必要とする API を使用するための適切なアクセス許可が必要になります。 一般的には、アドインに必要な最低限のアクセス許可を指定する必要があります。

アクセス許可には 4 つのレベルがあります。 **制限付き**、 **読み取りアイテム**、 **読み取り/書き込みアイテム**、 **および読み取り/書き込みメールボックス**。 詳細については、以下をご覧ください。 詳細については、「[Outlook アドインのアクセス許可モデルを理解する](understanding-outlook-add-in-permissions.md)」を参照してください。

## <a name="mailbox-object"></a>Mailbox オブジェクト

[!include[information about Mailbox object](../includes/mailbox-object-desc.md)]

## <a name="see-also"></a>関連項目

- [Outlook アドインのマニフェスト](manifests.md)
- [Outlook API 要件セットについて](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets)
- [Outlook アドインのアクセス許可について](understanding-outlook-add-in-permissions.md)。
- [Office アドインのプライバシーとセキュリティ](../concepts/privacy-and-security.md)
