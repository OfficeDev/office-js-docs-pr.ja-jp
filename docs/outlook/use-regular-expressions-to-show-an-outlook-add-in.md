---
title: 正規表現アクティブ化ルールを使用してアドインを表示する
description: Outlook コンテキスト アドインで正規表現アクティブ化ルールを使用する方法について説明します。
ms.date: 10/03/2022
ms.localizationpriority: medium
ms.openlocfilehash: ed2fbbfcf7bf55e04f4ec6f225e29fb43ec99639
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: ja-JP
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467091"
---
# <a name="use-regular-expression-activation-rules-to-show-an-outlook-add-in"></a>正規表現アクティブ化ルールを使用して Outlook アドインを表示する

メッセージの特定のフィールドで一致がある場合に[コンテキスト アドイン](contextual-outlook-add-ins.md)をアクティブ化するように正規表現ルールを指定します。 コンテキスト アドインは、読み取りモードでのみアクティブ化されます。 ユーザーがアイテムを作成している場合、Outlook はコンテキスト アドインをアクティブ化しません。 また、デジタル署名されたアイテムなど、Outlook でアドインがアクティブ化されないシナリオもあります。 詳細については、「[Outlook アドインのアクティブ化ルール](activation-rules.md)」を参照してください。

[!include[JSON manifest does not support contextual add-ins](../includes/json-manifest-outlook-contextual-not-supported.md)]

アドイン XML マニフェストでは、[ItemHasRegularExpressionMatch](/javascript/api/manifest/rule#itemhasregularexpressionmatch-rule) ルールまたは [ItemHasKnownEntity](/javascript/api/manifest/rule#itemhasknownentity-rule) ルールの一部として正規表現を指定することができます。 ルールは [DetectedEntity](/javascript/api/manifest/extensionpoint#detectedentity) 拡張点で指定されます。

Outlook では、クライアント コンピューターのブラウザーで使用する JavaScript インタープリターのルールに基づいて正規表現を評価します。 Outlook では、すべての XML プロセッサでもサポートされているものと同じ特殊文字リストをサポートしています。 次の表は、このような特殊文字を示しています。 正規表現でこれらの文字を使用するには、次の表に示すように、対応する文字のエスケープ シーケンスを指定します。

|文字|説明|使用するエスケープ シーケンス|
|:-----|:-----|:-----|
|`"`|二重引用符|`&quot;`|
|`&`|アンパサンド|`&amp;`|
|`'`|アポストロフィ|`&apos;`|
|`<`|より小さい|`&lt;`|
|`>`|より大きい|`&gt;`|

## <a name="itemhasregularexpressionmatch-rule"></a>ItemHasRegularExpressionMatch ルール

`ItemHasRegularExpressionMatch` ルールはサポートされているプロパティの特定の値に基づいてアドインのアクティブ化を制御するのに便利です。 `ItemHasRegularExpressionMatch` ルールには以下の属性があります。

|属性名|説明|
|:-----|:-----|
|`RegExName`|アドインのコードで参照できるように、正規表現の名前を指定します。|
|`RegExValue`|アドインを表示するかどうかを判断するために評価する正規表現を指定します。|
|`PropertyName`|正規表現の評価対象となるプロパティの名前を指定します。 有効な値は `BodyAsHTML`、`BodyAsPlaintext`、`SenderSMTPAddress`、`Subject` です。<br/><br/>`BodyAsHTML` を指定した場合、アイテムの本文が HTML の場合にのみ Outlook は正規表現を適用します。 HTML 以外の場合、Outlook はその正規表現に対して一致を返しません。<br/><br/>`BodyAsPlaintext` を指定すると、Outlook はアイテムの本文に対して正規表現を常に適用します。<br/><br/>**大事な：** 要素の **Highlight** 属性を指定する必要がある場合は、**PropertyName** 属性`BodyAsPlaintext`**\<Rule\>** を . |
|`IgnoreCase`|`RegExName` で指定された正規表現のマッチングで大文字と小文字の違いを無視するかどうかを指定します。|
| `Highlight` | クライアントが一致するテキストを強調表示にする方法を指定します。 この要素は、`ExtensionPoint` 要素内の `Rule` 要素にのみ適用できます。 `all` または `none` のいずれかになります。 指定のない場合、既定値は `all` に設定されます。<br/><br/>**大事な：** 要素で **Highlight** 属性を **\<Rule\>** 指定するには、 **PropertyName** 属性 `BodyAsPlaintext`を . |

### <a name="best-practices-for-using-regular-expressions-in-rules"></a>ルールで正規表現を使用する場合のベスト プラクティス

正規表現を使用する場合は、以下に特に注意してください。

- アイテムの本文にルールを指定する `ItemHasRegularExpressionMatch` 場合、正規表現は本文をさらにフィルター処理する必要があり、アイテムの本文全体を返そうとしないでください。 アイテムの本文全体を取得しようとするなどの `.*` 正規表現を使用しても、期待される結果が返されるとは限りません。
- あるブラウザーで返されたプレーンテキストの本文は、別のブラウザーではわずかに異なることがあります。 `BodyAsPlaintext` を `PropertyName` 属性として `ItemHasRegularExpressionMatch` ルールを使用する場合は、アドインのサポート対象であるすべてのブラウザーで正規表現をテストします。

    さまざまなブラウザーがさまざまな方法で選択したアイテムの本文を取得するため、使用している正規表現が、本文の一部として返される可能性がある微妙な違いをサポートしていることを確認する必要があります。 たとえば、アイテムの本文を取得するために、Internet Explorer 9 などのブラウザーでは DOM の `innerText` プロパティを使用し、Firefox などのその他のブラウザーでは `.textContent()` メソッドを使用します。 また、さまざまなブラウザーが異なる改行を返す場合があります。改行は、Internet Explorer では `\r\n`、Firefox および Chrome では `\n` です。 詳細については、「[W3C DOM の互換性 - HTML](https://quirksmode.org/dom/html/)」(W3C DOM の互換性 - HTML) を参照してください。

- アイテムの HTML 形式の本文は、Outlook リッチ クライアントと、Outlook on the web または Outlook モバイルとでは若干異なります。 正規表現を正確に定義する必要があります。

- 正規表現が適用されている Outlook クライアント、デバイスの種類、またはプロパティに応じて、正規表現をアクティブ化ルールとして設計する際に注意する必要があるクライアントごとに、他のベスト プラクティスと制限があります。 詳細については、「 [Outlook アドインのアクティブ化と JavaScript API の制限](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)」を参照してください。

### <a name="examples"></a>例

次の `ItemHasRegularExpressionMatch` ルールでは、大文字小文字に関係なく、送信者の SMTP メール アドレスが `@contoso` と一致した場合にアドインをアクティブにします。

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    RegExName="addressMatches"
    RegExValue="@[cC][oO][nN][tT][oO][sS][oO]"
    PropertyName="SenderSMTPAddress"
/>
```

次の例では、`IgnoreCase` 属性を使用して同じ正規表現を指定しています。

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    RegExName="addressMatches"
    RegExValue="@contoso"
    PropertyName="SenderSMTPAddress"
    IgnoreCase="true"
/>
```

次の `ItemHasRegularExpressionMatch` ルールでは、現在のアイテムの本文に株式銘柄コードが含まれている場合にアドインをアクティブにします。

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    PropertyName="BodyAsPlaintext"
    RegExName="TickerSymbols"
    RegExValue="\b(NYSE|NASDAQ|AMEX):\s*[A-Za-z]+\b"/>

```

## <a name="itemhasknownentity-rule"></a>ItemHasKnownEntity ルール

`ItemHasKnownEntity` ルールでは、選択したアイテムの件名または本文でのエンティティの存在に基づいてアドインをアクティブにします。 [EntityType](/javascript/api/outlook/office.mailboxenums.entitytype) タイプはサポートされるエンティティを定義します。 `ItemHasKnownEntity` ルールに正規表現を適用すると、アクティブ化がエンティティの値のサブセット (特定の URL セットまたは、特定の市外局番の電話番号など) に基づく点で、利便性が増します。

> [!NOTE]
> マニフェストに指定されている既定のロケールに関係なく、Outlook が抽出できるのは英語のエンティティ文字列だけです。 エンティティの種類をサポートするのは `MeetingSuggestion` メッセージのみです。予定ではサポートされません。 **送信済みアイテム** フォルダー内のアイテムからエンティティを抽出したり、ルールを`ItemHasKnownEntity`使用して送信済み **アイテム** フォルダー内のアイテムのアドインをアクティブ化したりすることはできません。

`ItemHasKnownEntity` ルールでは、以下の表にある属性をサポートしています。 `ItemHasKnownEntity` ルールで正規表現の指定が任意の場合、エンティティ フィルターとして正規表現を使用するには、`RegExFilter` 属性と `FilterName` 属性の両方を指定する必要があります。

|属性名|説明|
|:-----|:-----|
|`EntityType`|このルールが `true` と評価するために見つける必要のあるエンティティの型を指定します。 複数のルールを使用して複数のエンティティの型を指定します。|
|`RegExFilter`|`EntityType` で指定されているエンティティのインスタンスをさらにフィルター処理する正規表現を指定します。|
|`FilterName`|`RegExFilter` で指定されている正規表現の名前を指定し、それ以降にコードでその正規表現を参照できるようにします。|
|`IgnoreCase`|`RegExFilter` で指定された正規表現のマッチングで大文字と小文字の違いを無視するかどうかを指定します。|

### <a name="examples"></a>例

次の `ItemHasKnownEntity` ルールでは、現在のアイテムの件名または本文に URL が存在し、URL に文字列 `youtube` (大文字小文字は区別しない) が含まれている場合、常にアドインをアクティブにします。

```XML
<Rule xsi:type="ItemHasKnownEntity"
    EntityType="Url"
    RegExFilter="youtube"
    FilterName="youtube"
    IgnoreCase="true"/>
```

## <a name="using-regular-expression-results-in-code"></a>コードでの正規表現の結果の使用

現在の項目で次のメソッドを使用して、正規表現に一致する文字列を取得できます。

- [getRegExMatches](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) は、アドインの `ItemHasRegularExpressionMatch` ルールと `ItemHasKnownEntity` ルールで指定されているすべての正規表現について、現在のアイテムで一致するものを返します。

- [getRegExMatchesByName](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) は、アドインの `ItemHasRegularExpressionMatch` ルールで指定されている特定された正規表現について、現在のアイテムで一致するものを返します。

- [getFilteredEntitiesByName](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) は、アドインの `ItemHasKnownEntity` ルールで指定されている正規表現について、一致するものを含むエンティティのインスタンス全体を返します。

正規表現が評価されると、配列オブジェクトに入れてアドインに一致が返されます。 `getRegExMatches` については、そのオブジェクトに正規表現の名前の識別子があります。

> [!NOTE]
> Outlook は、配列内の特定の順序で一致を返しません。 また、同じメールボックス内の同じアイテムでこれらの各クライアントで同じアドインを実行する場合でも、この配列で同じ順序で一致が返されることを想定しないでください。

### <a name="examples"></a>例

`videoURL` という名前の正規表現を使用する `ItemHasRegularExpressionMatch` ルールを含めたコレクションの例を次に示します。

```XML
<Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="videoURL" RegExValue="http://www\.youtube\.com/watch\?v=[a-zA-Z0-9_-]{11}" PropertyName="BodyAsPlaintext"/>
</Rule>
```

次の例では、現在のアイテムの `getRegExMatches` を使用して、変数 `videos` を前の `ItemHasRegularExpressionMatch` ルールの結果に設定します。

```js
const videos = Office.context.mailbox.item.getRegExMatches().videoURL;
```

Multiple matches are stored as array elements in that object. The following code example shows how to iterate over the matches for a regular expression named  `reg1` to build a string to display as HTML.

```js
function initDialer()
{
    let myEntities;
    let myString;
    let myCell;
    myEntities = Office.context.mailbox.item.getRegExMatches();

    myString = "";
    myCell = document.getElementById('dialerholder');
    // Loop over the myEntities collection.
    for (let i in myEntities.reg1) {
        myString += "<p><a href='callto:tel:" + myEntities.reg1[i] + "'>" + myEntities.reg1[i] + "</a></p>";
    }

    myCell.innerHTML = myString;
}
```

`MeetingSuggestion` エンティティと `CampSuggestion` という正規表現を指定する `ItemHasKnownEntity` ルールの例を次に示します。 現在選択されているアイテムに会議の提案が含まれ、件名または本文に `WonderCamp` という用語があると判明した場合、Outlook はこのアドインをアクティブにします。

```XML
<Rule xsi:type="ItemHasKnownEntity"
    EntityType="MeetingSuggestion"
    RegExFilter="WonderCamp"
    FilterName="CampSuggestion"
    IgnoreCase="false"/>
```

次のコード例では、現在のアイテムの `getFilteredEntitiesByName` を使用して変数 `suggestions` を設定し、前の `ItemHasKnownEntity` ルールで検出された会議の提案の配列を取得します。

```js
const suggestions = Office.context.mailbox.item.getFilteredEntitiesByName("CampSuggestion");
```

## <a name="see-also"></a>関連項目

- [Outlook アドイン: Contoso 社の注文番号](https://github.com/OfficeDev/Outlook-Add-In-Contextual-Regex) - 正規表現の一致に基づいてアクティブ化されるコンテキスト アドインのサンプル。
- [閲覧フォーム用の Outlook アドインを作成する](read-scenario.md)
- [Outlook アドインのアクティブ化ルール](activation-rules.md)
- [Outlook アドインのアクティブ化と JavaScript API の制限](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [Outlook アイテム内の文字列を既知のエンティティとして照合する](match-strings-in-an-item-as-well-known-entities.md)
- [.NET Framework での正規表現のベスト プラクティス](/dotnet/standard/base-types/best-practices)
